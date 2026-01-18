# Copilot Comments

This fork contains an initial migration scaffold to move EAS-based sync logic to Microsoft Graph.

What I added
- content/includes/graphClient.js: a starter Graph client with token refresh and helper methods for:
  - calendars list
  - contact folders list
  - calendar events delta
  - contacts delta
  - create/update/delete event/contact

Notes and next steps
- Replace EAS WBXML sync usage with Graph delta endpoints.
- Implement mapping helpers (local <-> Graph JSON) for events and contacts.
- Persist Graph folder IDs and deltaLinks per folder in accountData.
- Replace old OAuth scope usage with Graph scopes (Calendars.ReadWrite, Contacts.ReadWrite, Tasks.ReadWrite, offline_access).
- Test thoroughly: timezone handling, attachments, token refresh, rate limits.

If you'd like I can prepare mapping helpers (graphMappings.js) and a PR patch that wires contacts or calendars end-to-end.