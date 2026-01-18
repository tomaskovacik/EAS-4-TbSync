/*
  Basic Microsoft Graph client for EAS-4-TbSync migration.

  Responsibilities:
  - Token management (store access/refresh token in accountData)
  - API request wrapper with automatic token refresh
  - Helpers for listing calendars/contactFolders and performing delta queries

  NOTE: This is a starting skeleton. Integrate with your existing OAuth helper
  if you already have one (eas.network.getOAuthObj / OAuth2 class).
*/

"use strict";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

var graphClient = {
  // Ensure we have a valid access token for the account; refresh using refresh token if needed.
  ensureAccessToken: async function(accountData) {
    // accountData storage keys (choose names consistent with your storage)
    let accessToken = accountData.getAccountProperty("graph.accessToken") || "";
    let refreshToken = accountData.getAccountProperty("graph.refreshToken") || "";
    let expiresAt = accountData.getAccountProperty("graph.expiresAt") || 0; // ms epoch

    let now = Date.now();
    if (accessToken && expiresAt > (now + 60000)) { // token valid for >= 60s
      return accessToken;
    }

    // If there's an OAuth helper available in your codebase, prefer reusing it.
    // Fallback: use refresh token grant against v2 token endpoint.
    if (!refreshToken) {
      // No refresh token — signal that interactive sign-in is required.
      throw new Error("graph_no_refresh_token");
    }

    // Token endpoint
    const tokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

    // Client id -- prefer reading from settings, or reuse existing client id.
    // Make sure the app is registered and supports the requested scopes.
    const clientId = accountData.getAccountProperty("graph.clientId") || "replace-with-your-client-id";

    // If you're using a confidential client you must include client_secret; in extensions, PKCE is recommended.
    // Here we assume refresh token flow only.
    const body = new URLSearchParams();
    body.append("client_id", clientId);
    body.append("grant_type", "refresh_token");
    body.append("refresh_token", refreshToken);
    body.append("scope", "offline_access openid profile Calendars.ReadWrite Contacts.ReadWrite Tasks.ReadWrite");

    let resp = await fetch(tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString()
    });

    if (!resp.ok) {
      // Refresh failed; clear tokens so caller can prompt interactive sign-in
      accountData.setAccountProperty("graph.accessToken", "");
      accountData.setAccountProperty("graph.refreshToken", "");
      accountData.setAccountProperty("graph.expiresAt", 0);
      throw new Error("graph_refresh_failed");
    }

    let tok = await resp.json();
    if (!tok.access_token) {
      throw new Error("graph_refresh_invalid_response");
    }

    // Save tokens and expiry
    let expiresIn = tok.expires_in || 3600;
    accessToken = tok.access_token;
    refreshToken = tok.refresh_token || refreshToken;
    expiresAt = Date.now() + (expiresIn * 1000);

    accountData.setAccountProperty("graph.accessToken", accessToken);
    accountData.setAccountProperty("graph.refreshToken", refreshToken);
    accountData.setAccountProperty("graph.expiresAt", expiresAt);

    return accessToken;
  },

  // Generic Graph API request
  apiRequest: async function(accountData, method, path, body = null, extraHeaders = {}) {
    let token = await graphClient.ensureAccessToken(accountData);

    let headers = Object.assign({
      "Authorization": "Bearer " + token,
      "Accept": "application/json"
    }, extraHeaders);

    if (body && !(body instanceof FormData)) {
      headers["Content-Type"] = "application/json";
      body = JSON.stringify(body);
    }

    const url = GRAPH_BASE + path;
    let resp = await fetch(url, {
      method: method,
      headers: headers,
      body: body
    });

    // Basic handling of auth failure to surface refresh needs to caller
    if (resp.status === 401) {
      // try one refresh + retry
      await accountData.setAccountProperty("graph.accessToken", "");
      token = await graphClient.ensureAccessToken(accountData);
      headers["Authorization"] = "Bearer " + token;
      resp = await fetch(url, { method, headers, body });
    }

    if (!resp.ok) {
      let text = await resp.text();
      let err = text;
      try { err = JSON.parse(text); } catch (e) {}
      let error = new Error("GraphRequestFailed: " + resp.status);
      error.status = resp.status;
      error.body = err;
      throw error;
    }

    // For responses that have no content (DELETE), return empty
    if (resp.status === 204) return null;
    return resp.json();
  },

  // Calendars list
  getCalendars: async function(accountData) {
    // returns list of calendars
    return await graphClient.apiRequest(accountData, "GET", "/me/calendars");
  },

  // Contact folders list
  getContactFolders: async function(accountData) {
    return await graphClient.apiRequest(accountData, "GET", "/me/contactFolders");
  },

  // Calendar events delta. Supply calendarId and optionally a deltaLink (previously saved).
  eventsDelta: async function(accountData, calendarId, deltaLink) {
    // If you have a stored deltaLink, it is a full URL; Graph allows calling it directly.
    if (deltaLink) {
      // deltaLink may point to graph.microsoft.com/... so call it directly with token
      let token = await graphClient.ensureAccessToken(accountData);
      let resp = await fetch(deltaLink, {
        method: "GET",
        headers: {
          "Authorization": "Bearer " + token,
          "Accept": "application/json"
        }
      });
      if (!resp.ok) {
        let err = await resp.text();
        throw new Error("Graph delta fetch failed: " + resp.status + " " + err);
      }
      return await resp.json(); // contains value[], @odata.nextLink or @odata.deltaLink
    } else {
      // Start a new delta query
      return await graphClient.apiRequest(accountData, "GET", `/me/calendars/${encodeURIComponent(calendarId)}/events/delta`);
    }
  },

  // Contacts delta
  contactsDelta: async function(accountData, deltaLink) {
    if (deltaLink) {
      let token = await graphClient.ensureAccessToken(accountData);
      let resp = await fetch(deltaLink, {
        method: "GET",
        headers: {
          "Authorization": "Bearer " + token,
          "Accept": "application/json"
        }
      });
      if (!resp.ok) {
        let err = await resp.text();
        throw new Error("Graph delta fetch failed: " + resp.status + " " + err);
      }
      return await resp.json();
    } else {
      return await graphClient.apiRequest(accountData, "GET", `/me/contacts/delta`);
    }
  },

  // Example create/update/delete for calendar events and contacts
  createEvent: async function(accountData, calendarId, eventObj) {
    return await graphClient.apiRequest(accountData, "POST", `/me/calendars/${encodeURIComponent(calendarId)}/events`, eventObj);
  },

  updateEvent: async function(accountData, eventId, eventObj) {
    return await graphClient.apiRequest(accountData, "PATCH", `/me/events/${encodeURIComponent(eventId)}`, eventObj);
  },

  deleteEvent: async function(accountData, eventId) {
    return await graphClient.apiRequest(accountData, "DELETE", `/me/events/${encodeURIComponent(eventId)}`);
  },

  createContact: async function(accountData, contactObj) {
    return await graphClient.apiRequest(accountData, "POST", `/me/contacts`, contactObj);
  },

  updateContact: async function(accountData, contactId, contactObj) {
    return await graphClient.apiRequest(accountData, "PATCH", `/me/contacts/${encodeURIComponent(contactId)}`, contactObj);
  },

  deleteContact: async function(accountData, contactId) {
    return await graphClient.apiRequest(accountData, "DELETE", `/me/contacts/${encodeURIComponent(contactId)}`);
  }
};

this.graphClient = graphClient;