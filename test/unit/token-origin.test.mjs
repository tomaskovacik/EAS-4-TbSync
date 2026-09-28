/**
 * Unit tests for dropping our own Origin from token-endpoint requests
 * (AADSTS9002326 with a loopback redirect URI).
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

const OWN = "moz-extension://a3ff7645-d446-4512-b0ad-d53215d32ff5";
globalThis.browser.runtime = {
  ...globalThis.browser.runtime,
  getURL: (p) => `${OWN}/${p}`,
};

const { stripOwnOrigin } = await import("../../src/modules/token-origin.mjs");

test("our own Origin is removed, everything else kept", () => {
  const out = stripOwnOrigin({
    requestHeaders: [
      { name: "Content-Type", value: "application/x-www-form-urlencoded" },
      { name: "Origin", value: OWN },
      { name: "Accept", value: "*/*" },
    ],
  });
  assert.deepEqual(out.requestHeaders, [
    { name: "Content-Type", value: "application/x-www-form-urlencoded" },
    { name: "Accept", value: "*/*" },
  ]);
});

test("another origin is left alone", () => {
  const out = stripOwnOrigin({
    requestHeaders: [{ name: "Origin", value: "https://outlook.office.com" }],
  });
  assert.deepEqual(out, {}, "no change requested");
});

test("a request without Origin is left alone", () => {
  assert.deepEqual(
    stripOwnOrigin({ requestHeaders: [{ name: "Accept", value: "*/*" }] }),
    {},
  );
});
