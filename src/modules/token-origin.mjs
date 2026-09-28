/**
 * Strip our own `Origin` header from Microsoft token-endpoint requests.
 *
 * `fetch()` from the background page stamps `Origin: moz-extension://<uuid>`
 * on the POST to `…/oauth2/v2.0/token`. Entra reads an Origin as "a browser
 * page is redeeming this code" and, for a redirect URI registered as
 * "Mobile and desktop applications" such as `http://localhost`, refuses:
 *
 *   AADSTS9002326: Cross-origin token redemption is permitted only for the
 *   'Single-Page Application' client-type.
 *
 * Registering the URI as an SPA instead is no way out: SPA refresh tokens
 * live 24 hours, so every account would have to sign in daily. The request
 * is a native client's in every way that matters (Thunderbird's own OAuth2
 * sends none), so the header is dropped - only ours, only on the token
 * endpoint.
 */

const FILTER_URLS = ["https://login.microsoftonline.com/*/oauth2/v2.0/token"];

export function installTokenOriginStripper() {
  if (!browser.webRequest?.onBeforeSendHeaders) {
    console.warn(
      "[eas-4-tbsync] webRequest API unavailable; token Origin stripping disabled",
    );
    return;
  }
  browser.webRequest.onBeforeSendHeaders.addListener(
    stripOwnOrigin,
    { urls: FILTER_URLS },
    ["blocking", "requestHeaders"],
  );
}

/** Our own origin, e.g. `moz-extension://a3ff7645-…` (no trailing slash). */
function ownOrigin() {
  return browser.runtime.getURL("").replace(/\/$/, "");
}

export function stripOwnOrigin(details) {
  const headers = details.requestHeaders;
  if (!Array.isArray(headers)) return {};
  const ours = ownOrigin();
  const kept = headers.filter(
    (h) => !(h.name?.toLowerCase() === "origin" && h.value === ours),
  );
  // Somebody else's request (another add-on, a web page) → untouched.
  if (kept.length === headers.length) return {};
  return { requestHeaders: kept };
}
