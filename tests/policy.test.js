"use strict";

// What a request is allowed to do. These are the checks that keep an endpoint
// which runs arbitrary JavaScript from being reachable by anything other than
// the person at the keyboard.

const test = require("node:test");
const assert = require("node:assert");
const path = require("node:path");

const {
  checkCompileRequest,
  parseCompileBody,
  resolveStaticPath,
  isOwnOrigin,
  mimeFor,
} = require("../src/policy");

const JSON_CT = { "content-type": "application/json", host: "localhost:3000" };

test("a same-origin JSON post is allowed", () => {
  assert.strictEqual(
    checkCompileRequest({ ...JSON_CT, origin: "http://localhost:3000" }).ok,
    true,
  );
});

test("localhost and 127.0.0.1 count as the same origin", () => {
  assert.strictEqual(
    checkCompileRequest({
      "content-type": "application/json",
      host: "127.0.0.1:3000",
      origin: "http://localhost:3000",
    }).ok,
    true,
  );
});

test("a request with no Origin at all is allowed", () => {
  // curl and the tests send none; a browser always sends one cross-origin.
  assert.strictEqual(checkCompileRequest(JSON_CT).ok, true);
});

// The reason this check exists: a page on any website can POST to
// http://localhost:3000 without a preflight, and Crepuscule Slog would have run the code.
test("a post from another site is refused", () => {
  const verdict = checkCompileRequest({ ...JSON_CT, origin: "https://evil.example" });
  assert.strictEqual(verdict.ok, false);
  assert.strictEqual(verdict.status, 403);
  assert.match(verdict.error, /another site/);
});

test("a loopback origin on a different port is refused", () => {
  // Another local dev server is still someone else.
  const verdict = checkCompileRequest({ ...JSON_CT, origin: "http://localhost:8080" });
  assert.strictEqual(verdict.ok, false);
});

test("an origin that merely looks like localhost is refused", () => {
  for (const origin of [
    "http://localhost.evil.example:3000",
    "http://127.0.0.1.evil.example:3000",
    "https://localhost:3000",
    "null",
  ]) {
    assert.strictEqual(
      checkCompileRequest({ ...JSON_CT, origin }).ok,
      false,
      `${origin} should be refused`,
    );
  }
});

// text/plain is what makes a cross-origin POST a "simple request" that needs
// no permission from the server. Demanding JSON takes that route away.
test("a non-JSON content type is refused", () => {
  for (const type of ["text/plain;charset=UTF-8", "application/x-www-form-urlencoded", ""]) {
    const verdict = checkCompileRequest({ "content-type": type, host: "localhost:3000" });
    assert.strictEqual(verdict.ok, false, `${type} should be refused`);
    assert.strictEqual(verdict.status, 415);
  }
});

test("a JSON content type with a charset is accepted", () => {
  assert.strictEqual(
    checkCompileRequest({ "content-type": "application/json; charset=utf-8", host: "x:1" }).ok,
    true,
  );
});

test("isOwnOrigin handles a missing port as port 80", () => {
  assert.strictEqual(isOwnOrigin("http://localhost", "localhost"), true);
  assert.strictEqual(isOwnOrigin("http://localhost", "localhost:3000"), false);
});

// ─── Body ────────────────────────────────────────────────────────────────────

test("a well-formed body is accepted", () => {
  const job = parseCompileBody(JSON.stringify({ format: "docx", code: "doc;" }));
  assert.deepStrictEqual(job, { ok: true, format: "docx", code: "doc;" });
});

test("malformed JSON, missing fields and wrong types are all refused", () => {
  assert.strictEqual(parseCompileBody("{oops").status, 400);
  assert.strictEqual(parseCompileBody("null").status, 400);
  assert.strictEqual(parseCompileBody('"a string"').status, 400);
  assert.match(parseCompileBody(JSON.stringify({ format: "docx" })).error, /Missing/);
  assert.match(parseCompileBody(JSON.stringify({ code: "x" })).error, /Missing/);
  assert.match(
    parseCompileBody(JSON.stringify({ format: "docx", code: 42 })).error,
    /must both be strings/,
  );
});

test("an unknown format is refused before any code runs", () => {
  const job = parseCompileBody(JSON.stringify({ format: "rtf", code: "x;" }));
  assert.strictEqual(job.ok, false);
  assert.match(job.error, /Unknown format: rtf/);
});

// ─── Static files ────────────────────────────────────────────────────────────

const ROOT = path.resolve("public");

test("the bare path serves the page", () => {
  assert.strictEqual(resolveStaticPath("/", ROOT), path.join(ROOT, "index.html"));
  assert.strictEqual(resolveStaticPath("/?x=1", ROOT), path.join(ROOT, "index.html"));
});

test("an ordinary asset resolves inside the public directory", () => {
  assert.strictEqual(resolveStaticPath("/style.css", ROOT), path.join(ROOT, "style.css"));
  assert.strictEqual(resolveStaticPath("/crepuscule-slog.js?v=2", ROOT), path.join(ROOT, "crepuscule-slog.js"));
});

test("traversal out of the public directory is refused", () => {
  for (const url of [
    "/../server.js",
    "/../../CLAUDE.md",
    "/%2e%2e/server.js",
    "/..%2fserver.js",
    "/..\\server.js",
    "/subdir/../../package.json",
  ]) {
    assert.strictEqual(resolveStaticPath(url, ROOT), null, `${url} should be refused`);
  }
});

test("a malformed percent-encoding or a null byte is refused", () => {
  assert.strictEqual(resolveStaticPath("/%zz", ROOT), null);
  assert.strictEqual(resolveStaticPath("/index.html%00.png", ROOT), null);
});

test("content types come from the extension", () => {
  assert.match(mimeFor("/a/index.html"), /^text\/html/);
  assert.match(mimeFor("crepuscule-slog.JS"), /javascript/);
  assert.strictEqual(mimeFor("thing.bin"), "application/octet-stream");
});
