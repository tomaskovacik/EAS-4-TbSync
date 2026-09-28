# PoC: Microsoft sign-in in the system browser without copy/paste

**This branch: own loopback listener** (`src/experiments/OAuthLoopback/`). The other variant: [`oauth-tb-oauth2-poc`](https://github.com/tomaskovacik/EAS-4-TbSync/tree/oauth-tb-oauth2-poc).

> ⚠️ **Proof of concept, AI-generated.** This code was written with an AI
> assistant (Claude Code) and tested by hand on one machine. Treat it as a
> starting point for discussion, not as merge-ready code: **a full review is
> required** before any of it is used.

Follow-up to [#357](https://github.com/jobisoft/EAS-4-TbSync/pull/357). Both
PoC branches are built on top of #357 and keep its paste dialog as the
fallback.

## Idea

#357 opens Microsoft sign-in in the system browser, but because the redirect
goes to the `nativeclient` page, the user has to copy the final URL back into
Thunderbird. Thunderbird itself avoids that for its own Microsoft accounts
(default since TB 153, [bug 2019793](https://bugzilla.mozilla.org/show_bug.cgi?id=2019793),
[bug 2024728](https://bugzilla.mozilla.org/show_bug.cgi?id=2024728)): it
redirects to `http://localhost:<random port>` and catches the redirect on a
short-lived loopback listener (`ExternalRequest` in
`mailnews/base/src/OAuth2.sys.mjs`). These branches do the same for EAS.

## Requirement: `http://localhost` redirect URI

**This only works when the client ID in use has `http://localhost`
registered as a redirect URI.** It was tested with a custom app registration
in a private tenant, not with the community client ID.

App registration settings that were needed:

- Authentication → platform **"Mobile and desktop applications"** →
  redirect URI `http://localhost` (Entra ignores the port for loopback URIs).
  Under the "Web" platform it fails with AADSTS7000218 (client secret required).
- Authentication → Advanced settings → **Allow public client flows: Yes**.
- API permissions → Office 365 Exchange Online → `EAS.AccessAsUser.All`
  (delegated), admin consent granted.

For the community client ID `2980deeb-7460-4723-864a-f9b0f10cd992`, the
maintainer would add `http://localhost` under "Mobile and desktop
applications". The client ID, its permissions and existing refresh tokens
stay as they are. Until then the PoC uses the loopback route **only for a
custom client ID** and keeps #357's paste dialog for the community one
(`DEFAULT_CLIENT_HAS_LOOPBACK_REDIRECT = false` in `oauth.mjs`).

## How to try it

1. Build (`npm run build`) and install `dist/dev.xpi` (TB 153-157).
2. Add-on options → set **Custom OAuth client ID** to an app configured as
   above, and tick **Sign in using your default browser**.
3. Sign in or re-authenticate an Office 365 EAS account. The browser opens,
   a small "waiting" window stays in Thunderbird (Cancel), and the account
   signs in without any paste.

## The two variants

| | `oauth-loopback-poc` (own listener) | `oauth-tb-oauth2-poc` (Thunderbird's `OAuth2`) |
|---|---|---|
| Works on TB156 (tested) | ✅ | ✅ |
| Size vs #357 (incl. tests) | +714 / −30 | +587 / −9 |
| Local listener, browser page, tab closing, focus return | TbSync's own copy (`OAuthLoopback` experiment) | Thunderbird's own code, maintained by Mozilla (`TbOAuth2` experiment) |
| Reopen button while waiting | yes | no |
| Depends on | only the stable socket components | the internal `OAuth2.sys.mjs`, which can change between releases |
| Follows Thunderbird's `mailnews.oauth.*` prefs | no | yes |
| Experiment deprecation risk | same | same |

Both use TbSync's own client ID; Thunderbird's client ID is never used.
Both add PKCE (S256) on the browser route.

## Problems found while testing (and fixed)

1. **`URL` is not defined in experiment scope.** The `ext-*.js` sandbox only
   exposes `ChromeUtils` plus a fixed list (`_createExtGlobal()` in
   `ExtensionCommon.sys.mjs`), so `new URL()` threw and the listener never
   answered; the browser sat on "Waiting for localhost…". Replaced with a
   regex, and the listener now always answers.
2. **AADSTS9002326** (cross-origin token redemption is permitted only for the
   SPA client type). `fetch()` from the background page sends
   `Origin: moz-extension://…`, which Entra refuses for a "Mobile and desktop"
   redirect URI. Registering the URI as SPA is not an option: SPA refresh
   tokens expire after 24 hours. `src/modules/token-origin.mjs` removes our
   own `Origin` from requests to the token endpoint, the same way
   `anchor-mailbox.mjs` rewrites headers. This also affects refreshes in the
   Thunderbird-`OAuth2` variant, whose first exchange runs in Thunderbird.
3. **AADSTS7000218** (client secret required): app registration issue, see
   the requirement above.

## Open points

- Token refresh after the first hour was not yet verified in a long run.
- Both variants add an Experiment API. Thunderbird plans to restrict
  Experiments (postponed by a year from TB 153), so the long-term fix would
  be a native API, for example `identity.launchWebAuthFlow` using the system
  browser with a loopback redirect, reusing Thunderbird's `ExternalRequest`.
- Recent Firefox (Local Network Access, see
  [bug 2071842](https://bugzilla.mozilla.org/show_bug.cgi?id=2071842)) may ask
  the user to allow the redirect to `localhost`. Content blockers such as
  uBlock Origin may block it too.
- `strict_max_version` is raised to `157.*` (as on master) so the PoC
  installs on TB156; #357 is based on an older master.
