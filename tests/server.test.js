"use strict";

// End to end, over a real socket: the checks that only mean anything once
// there is a server listening.

const test = require("node:test");
const assert = require("node:assert");
const http = require("node:http");
const net = require("node:net");
const os = require("node:os");

const { createCrepusculeSlogServer, HOST } = require("../server");
const { createRunner } = require("../src/runner");
const { readZip, entryText } = require("./zip");

const DOC = `
const { Document, Paragraph } = require('docx');
const doc = new Document({ sections: [{ children: [new Paragraph('Over HTTP')] }] });
doc;
`;

// One server for the whole file; a fresh worker is spawned per compile anyway.
let server;
let port;

// This was 4000, and that is what made this file flake.
//
// Only one test below wants a short timeout -- the runaway-script one, which
// has to sit through it -- and that test now stands up its own server. Every
// other test here compiles something that finishes or throws immediately, so
// for them the timeout is not a deadline being measured, only a backstop
// against hanging the suite. Setting that backstop to four seconds meant a
// worker that was merely slow to be *scheduled* tripped it, and `server.js`
// answers a timeout with 504 rather than the 500 a throwing script produces.
//
// That is the flake `review.md` recorded three times without ever naming the
// assertion: "a script that fails comes back as an error the page can show",
// expecting 500 and getting 504. `node --test` runs the test files in
// parallel and `tests/runner.test.js` spins `while (true) {}` in a worker for
// three seconds, which pins a core, so the two files race. The machine only
// had to be busy enough for the losing side to take four seconds to get going.
const runner = createRunner({ timeoutMs: 30000 });

test.before(async () => {
  server = createCrepusculeSlogServer({ runner });
  await new Promise((resolve) => server.listen(0, HOST, resolve));
  port = server.address().port;
});

test.after(async () => {
  await new Promise((resolve) => server.close(resolve));
  await runner.shutdown();
});

function request(options, body, atPort = port) {
  return new Promise((resolve, reject) => {
    const req = http.request(
      {
        host: "127.0.0.1",
        port: atPort,
        method: "GET",
        ...options,
        headers: { host: `127.0.0.1:${atPort}`, ...(options.headers || {}) },
      },
      (res) => {
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () =>
          resolve({
            status: res.statusCode,
            headers: res.headers,
            body: Buffer.concat(chunks),
          }),
        );
      },
    );
    req.on("error", reject);
    if (body !== undefined) req.end(body);
    else req.end();
  });
}

function compile(payload, headers = {}, atPort = port) {
  const body = typeof payload === "string" ? payload : JSON.stringify(payload);
  return request(
    {
      method: "POST",
      path: "/compile",
      headers: {
        "content-type": "application/json",
        "content-length": Buffer.byteLength(body),
        ...headers,
      },
    },
    body,
    atPort,
  );
}

// ─── Where it listens ────────────────────────────────────────────────────────

// It used to be listen(PORT) with no address, which binds every interface —
// so anyone on the same network could POST code to it and have it run.
test("listens on loopback only", async () => {
  assert.match(server.address().address, /^(127\.0\.0\.1|::1)$/);

  const lan = Object.values(os.networkInterfaces())
    .flat()
    .find((i) => i.family === "IPv4" && !i.internal);
  if (!lan) return; // nothing to reach it by on this machine

  await assert.rejects(
    () =>
      new Promise((resolve, reject) => {
        const socket = net.connect({ host: lan.address, port }, () => {
          socket.destroy();
          resolve();
        });
        socket.setTimeout(2000, () => socket.destroy(new Error("timed out")));
        socket.on("error", reject);
      }),
    `the server answered on ${lan.address}, which is not loopback`,
  );
});

// ─── Compiling ───────────────────────────────────────────────────────────────

test("a pasted script comes back as a download", async () => {
  const res = await compile({ format: "docx", code: DOC });
  assert.strictEqual(res.status, 200);
  assert.match(res.headers["content-type"], /wordprocessingml\.document/);
  assert.match(res.headers["content-disposition"], /filename="output\.docx"/);
  assert.strictEqual(Number(res.headers["content-length"]), res.body.length);
  assert.match(entryText(readZip(res.body), "word/document.xml"), /Over HTTP/);
});

test("a script that fails comes back as an error the page can show", async () => {
  const res = await compile({ format: "xlsx", code: "throw new Error('bad sheet');\n" });
  assert.strictEqual(res.status, 500);
  assert.match(JSON.parse(res.body).error, /bad sheet/);
});

// The one test that genuinely wants a short timeout, since it has to sit
// through it. It gets its own server and runner so that four seconds is a
// deadline for this test alone and not for the whole file -- which is what
// used to make the throwing-script test above answer 504 under load.
test("a script that never finishes gets a 504, and the server keeps working", async () => {
  const impatient = createRunner({ timeoutMs: 4000 });
  const ownServer = createCrepusculeSlogServer({ runner: impatient });
  await new Promise((resolve) => ownServer.listen(0, HOST, resolve));
  const ownPort = ownServer.address().port;

  try {
    const res = await compile({ format: "docx", code: "while (true) {}\n" }, {}, ownPort);
    assert.strictEqual(res.status, 504);
    assert.match(JSON.parse(res.body).error, /did not finish/);

    const after = await compile({ format: "docx", code: DOC }, {}, ownPort);
    assert.strictEqual(after.status, 200, "the server stopped answering after a runaway script");
  } finally {
    await new Promise((resolve) => ownServer.close(resolve));
    await impatient.shutdown();
  }
});

// ─── What is refused ─────────────────────────────────────────────────────────

test("a cross-origin post is refused without running anything", async () => {
  const res = await compile({ format: "docx", code: DOC }, { origin: "https://evil.example" });
  assert.strictEqual(res.status, 403);
});

test("a text/plain post is refused", async () => {
  // The shape a hostile page would have to use to avoid a preflight.
  const res = await compile({ format: "docx", code: DOC }, { "content-type": "text/plain" });
  assert.strictEqual(res.status, 415);
});

test("an oversized paste is refused rather than buffered", async () => {
  const body = JSON.stringify({ format: "docx", code: "x".repeat(9 * 1024 * 1024) });
  const res = await compile(body);
  assert.strictEqual(res.status, 413);
});

test("a bad body is refused", async () => {
  assert.strictEqual((await compile("{not json")).status, 400);
  assert.strictEqual((await compile({ format: "docx" })).status, 400);
  assert.strictEqual((await compile({ format: "rtf", code: "x;" })).status, 400);
});

test("an unsupported method is refused", async () => {
  const res = await request({ method: "DELETE", path: "/compile" });
  assert.strictEqual(res.status, 405);
});

// ─── Static files ────────────────────────────────────────────────────────────

test("the page and its assets are served", async () => {
  const page = await request({ path: "/" });
  assert.strictEqual(page.status, 200);
  assert.match(page.headers["content-type"], /^text\/html/);
  assert.match(page.body.toString(), /<title>/i);

  const script = await request({ path: "/crepuscule-slog.js" });
  assert.strictEqual(script.status, 200);
  assert.match(script.headers["content-type"], /javascript/);
});

// The static root used to be the project directory, so everything in the
// repository was downloadable — including .git and anything under node_modules.
test("nothing outside the public directory is served", async () => {
  for (const path of [
    "/server.js",
    "/package.json",
    "/.git/config",
    "/node_modules/docx/package.json",
    "/../server.js",
    "/%2e%2e/server.js",
  ]) {
    const res = await request({ path });
    assert.ok(
      res.status === 403 || res.status === 404,
      `${path} returned ${res.status}; it should not be served`,
    );
  }
});

test("a missing file is a 404", async () => {
  assert.strictEqual((await request({ path: "/nope.css" })).status, 404);
});
