/*
 * Loopback redirect listener for OAuth sign-in in the system browser.
 *
 * A MailExtension cannot open a server socket, and `tabs.onUpdated` cannot
 * see a browser that is not Thunderbird, so this is the one piece of the
 * external-browser flow that needs an Experiment. It is a trimmed copy of
 * what Thunderbird itself does for mail accounts - `ExternalRequest` in
 * comm-central's mailnews/base/src/OAuth2.sys.mjs: an nsIServerSocket on a
 * random loopback port, the first `GET /…` taken as the redirect, a short
 * HTML answer to the browser, focus handed back to Thunderbird.
 *
 * Everything OAuth (state, PKCE, token exchange) stays in oauth.mjs; this
 * file only turns "the browser hit http://localhost:<port>/?…" into a URL.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

/* global Services, ExtensionCommon, Cc, Ci, Cr */

"use strict";

var { ExtensionUtils } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionUtils.sys.mjs",
);
var { ExtensionError } = ExtensionUtils;

const MAX_REQUEST_LINE_BYTES = 8192;

var OAuthLoopback = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    /** port → listener. One per sign-in in flight. */
    const listeners = new Map();

    const page = (key, fallback) => {
      const text = context.extension.localizeMessage(key) || fallback;
      const escaped = text.replace(
        /[&<>"]/g,
        (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c],
      );
      return (
        "<!doctype html><meta charset=utf-8><title>Thunderbird</title>" +
        "<body style='font-family:system-ui,sans-serif;margin:3em;'>" +
        `<p>${escaped}</p></body>`
      );
    };

    const returnFocus = () => {
      if (Services.focus.activeWindow) return;
      const win = Services.wm.getMostRecentWindow("mail:3pane");
      if (!win) return;
      win.restore();
      win.focus();
    };

    function createListener(extensionId) {
      const socket = Cc["@mozilla.org/network/server-socket;1"].createInstance(
        Ci.nsIServerSocket,
      );
      // Random port, loopback only, default backlog.
      socket.init(-1, true, -1);
      const port = socket.port;

      const { promise, resolve, reject } = Promise.withResolvers();
      // Nobody may be awaiting yet; keep an unobserved rejection quiet.
      promise.catch(() => {});

      const connections = new Set();
      let closed = false;
      let timer = null;

      const listener = {
        port,
        promise,
        QueryInterface: ChromeUtils.generateQI(["nsIServerSocketListener"]),

        close(reason = "Loopback listener closed") {
          if (closed) return;
          closed = true;
          if (timer) timer.cancel();
          socket.close();
          for (const c of connections) c.close();
          connections.clear();
          listeners.delete(port);
          reject(new ExtensionError(reason));
        },

        armTimeout(ms) {
          if (timer || closed) return;
          timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
          timer.initWithCallback(
            () => listener.close("Sign-in timed out"),
            ms,
            Ci.nsITimer.TYPE_ONE_SHOT,
          );
        },

        onSocketAccepted(_socket, transport) {
          if (closed) {
            transport.close(Cr.NS_ERROR_ABORT);
            return;
          }
          const input = transport
            .openInputStream(0, 0, 0)
            .QueryInterface(Ci.nsIAsyncInputStream);
          const output = transport.openOutputStream(0, 0, 0);
          let buffer = "";

          const connection = {
            QueryInterface: ChromeUtils.generateQI(["nsIInputStreamCallback"]),
            close() {
              try {
                input.close();
                output.close();
                transport.close(Cr.NS_OK);
              } catch {}
              connections.delete(connection);
            },
            respond(status, body) {
              const bytes = unescape(encodeURIComponent(body));
              const response =
                `HTTP/1.1 ${status}\r\n` +
                "Content-Type: text/html; charset=utf-8\r\n" +
                "Content-Security-Policy: default-src 'none'; style-src 'unsafe-inline'\r\n" +
                "Cache-Control: no-store\r\n" +
                `Content-Length: ${bytes.length}\r\n` +
                "Connection: close\r\n\r\n" +
                bytes;
              try {
                output.write(response, response.length);
              } catch {}
              connection.close();
            },
            onInputStreamReady(stream) {
              if (closed) {
                connection.close();
                return;
              }
              let available;
              try {
                available = stream.available();
              } catch {
                connection.close();
                return;
              }
              if (available > 0) {
                const s = Cc[
                  "@mozilla.org/scriptableinputstream;1"
                ].createInstance(Ci.nsIScriptableInputStream);
                s.init(stream);
                buffer += s.read(available);
              }
              const eol = buffer.indexOf("\r\n");
              if (eol < 0) {
                if (buffer.length > MAX_REQUEST_LINE_BYTES) {
                  connection.respond("400 Bad Request", "");
                  return;
                }
                stream.asyncWait(connection, 0, 0, Services.tm.mainThread);
                return;
              }
              const match = /^GET\s+(\/\S*)(?:\s|$)/.exec(buffer.slice(0, eol));
              // Only the redirect itself ends the sign-in. A favicon or any
              // other stray request gets a 404 and the listener stays up.
              if (!match || !/^\/(\?|$)/.test(match[1])) {
                connection.respond("404 Not Found", "");
                return;
              }
              const url = `http://localhost:${port}${match[1]}`;
              const failed = new URL(url).searchParams.has("error");
              connection.respond(
                "200 OK",
                failed
                  ? page(
                      "oauthLoopback.page.failed",
                      "Sign-in did not succeed. You can close this tab and return to Thunderbird.",
                    )
                  : page(
                      "oauthLoopback.page.done",
                      "Signed in. You can close this tab and return to Thunderbird.",
                    ),
              );
              // Settle before close(), which would otherwise reject.
              resolve(url);
              listener.close();
              returnFocus();
            },
          };
          connections.add(connection);
          input.asyncWait(connection, 0, 0, Services.tm.mainThread);
        },

        onStopListening() {},
      };

      socket.asyncListen(listener);
      listeners.set(port, listener);
      console.debug(`[${extensionId}] OAuth loopback listening on ${port}`);
      return listener;
    }

    // A reload or disable of the add-on must not leave sockets open.
    context.callOnClose({
      close() {
        for (const l of [...listeners.values()]) l.close("Add-on shut down");
      },
    });

    return {
      OAuthLoopback: {
        async listen() {
          try {
            return createListener(context.extension.id).port;
          } catch (e) {
            throw new ExtensionError(
              `Could not open a loopback listener: ${e.message}`,
            );
          }
        },
        async waitForRedirect(port, timeoutMs) {
          const listener = listeners.get(port);
          if (!listener) throw new ExtensionError("Unknown loopback listener");
          listener.armTimeout(timeoutMs);
          return listener.promise;
        },
        async close(port) {
          listeners.get(port)?.close();
        },
      },
    };
  }
};
