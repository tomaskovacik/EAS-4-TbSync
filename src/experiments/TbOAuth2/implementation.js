/*
 * Interactive OAuth sign-in through Thunderbird's own OAuth2 module.
 *
 * Thunderbird already runs Microsoft sign-in in the system browser for its
 * mail accounts: `OAuth2` in comm-central's mailnews/base/src/OAuth2.sys.mjs
 * picks `ExternalRequest` when `mailnews.oauth.useExternalBrowser` is on,
 * opens a loopback listener, launches the browser, checks `state`, and
 * exchanges the code with PKCE. None of that is reachable from a
 * MailExtension, so this experiment only builds an `OAuth2` object with the
 * add-on's own issuer details (never Thunderbird's client ID) and hands back
 * the tokens. Refreshing stays in the add-on.
 *
 * Internal API: OAuth2.sys.mjs has no stability promise; check it against
 * each Thunderbird version this add-on supports.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

/* global ExtensionCommon, Cc, Ci, Cr */

"use strict";

var { ExtensionUtils } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionUtils.sys.mjs",
);
var { ExtensionError } = ExtensionUtils;

/** OAuth2 rejects with a JSON string, the redirect URL itself, or an Error.
 *  Turn whichever it was into one readable sentence. */
function describe(reason) {
  const text = String(reason?.message ?? reason ?? "unknown error");
  if (/^https?:\/\//.test(text)) {
    try {
      const url = new URL(text);
      const error = url.searchParams.get("error");
      if (error) {
        const desc = url.searchParams.get("error_description") ?? "";
        return `Microsoft returned: ${error} ${desc}`.trim();
      }
    } catch {}
    return "authorization failed";
  }
  try {
    const parsed = JSON.parse(text);
    if (parsed?.error) return String(parsed.error);
  } catch {}
  return text;
}

var TbOAuth2 = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const { OAuth2 } = ChromeUtils.importESModule(
      "resource:///modules/OAuth2.sys.mjs",
    );
    /** requestId → { oauth, timer } */
    const pending = new Map();

    function abort(requestId, why) {
      const p = pending.get(requestId);
      if (!p) return;
      pending.delete(requestId);
      p.timer.cancel();
      try {
        // Closes the loopback listener (or the internal window).
        p.oauth.finishAuthorizationRequest();
      } catch {}
      // "cancelled" is the reason OAuth2 itself uses; it is ignored once a
      // code has arrived, so a late Cancel cannot spoil a finished sign-in.
      p.oauth.onAuthorizationFailed(
        Cr.NS_ERROR_ABORT,
        JSON.stringify({ error: why }),
        "cancelled",
      );
    }

    context.callOnClose({
      close() {
        for (const id of [...pending.keys()]) abort(id, "cancelled");
      },
    });

    return {
      TbOAuth2: {
        async authorize(requestId, details) {
          if (pending.has(requestId)) {
            throw new ExtensionError("Duplicate requestId");
          }
          const oauth = new OAuth2(details.scope, {
            authorizationEndpoint: details.authorizationEndpoint,
            tokenEndpoint: details.tokenEndpoint,
            clientId: details.clientId,
            redirectionEndpoint: details.redirectionEndpoint,
            useExternalBrowser: true,
            usePKCE: true,
          });
          // `telemetryData` lives on the prototype; give this instance its
          // own so nothing is recorded under a built-in issuer's name.
          oauth.telemetryData = {};
          if (details.loginHint) {
            oauth.extraAuthParams.push(["login_hint", details.loginHint]);
          }

          const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
          pending.set(requestId, { oauth, timer });
          timer.initWithCallback(
            () => abort(requestId, "timeout"),
            details.timeoutMs,
            Ci.nsITimer.TYPE_ONE_SHOT,
          );

          try {
            await oauth.connect(true, false);
          } catch (e) {
            throw new ExtensionError(describe(e));
          } finally {
            pending.get(requestId)?.timer.cancel();
            pending.delete(requestId);
          }
          return {
            accessToken: oauth.accessToken,
            refreshToken: oauth.refreshToken,
            expiresIn: Math.max(
              0,
              Math.round((oauth.tokenExpires - Date.now()) / 1000),
            ),
          };
        },

        async cancel(requestId) {
          abort(requestId, "cancelled");
        },
      },
    };
  }
};
