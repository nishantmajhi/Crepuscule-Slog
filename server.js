"use strict";

// The HTTP end of Crepuscule Slog: serve the page, take a POST of pasted code, hand it
// to a worker thread, send back the bytes that come out.
//
// Everything that decides anything lives elsewhere — src/policy.js for what a
// request is allowed to do, src/compile.js for what a script means, and
// src/runner.js for keeping a bad script from taking the process with it.

const http = require("http");
const fs = require("fs");
const path = require("path");

const {
  MAX_BODY_BYTES,
  OUTPUT_MIME,
  checkCompileRequest,
  parseCompileBody,
  resolveStaticPath,
  mimeFor,
} = require("./src/policy");
const { createRunner } = require("./src/runner");

const PORT = Number(process.env.CREPUSCULE_SLOG_PORT) || 3000;

// Loopback only, and deliberately so. Crepuscule Slog runs whatever JavaScript it is
// given; on 0.0.0.0 — which is what listen(PORT) alone gives you — that is
// remote code execution for anyone who can reach the machine.
const HOST = process.env.CREPUSCULE_SLOG_HOST || "127.0.0.1";

const PUBLIC_DIR = path.join(__dirname, "public");

function createCrepusculeSlogServer(options = {}) {
  const runner = options.runner || createRunner();

  const server = http.createServer(async (req, res) => {
    if (req.method === "POST" && req.url.split("?")[0] === "/compile") {
      await handleCompile(req, res, runner);
      return;
    }

    if (req.method === "GET" || req.method === "HEAD") {
      serveStatic(req, res);
      return;
    }

    res.writeHead(405, { "Content-Type": "text/plain" });
    res.end("Method not allowed");
  });

  server.on("close", () => runner.shutdown());
  return server;
}

async function handleCompile(req, res, runner) {
  const allowed = checkCompileRequest(req.headers);
  if (!allowed.ok) return sendJson(res, allowed.status, { error: allowed.error });

  let text;
  try {
    text = await readBody(req, MAX_BODY_BYTES);
  } catch (err) {
    if (err.code === "CREPUSCULE_SLOG_BODY_TOO_LARGE") {
      return sendJson(res, 413, {
        error: `That paste is larger than the ${Math.round(MAX_BODY_BYTES / (1024 * 1024))} MB limit.`,
      });
    }
    return sendJson(res, 400, { error: "Could not read the request body." });
  }

  const job = parseCompileBody(text);
  if (!job.ok) return sendJson(res, job.status, { error: job.error });

  try {
    const buf = await runner.run(job.format, job.code);
    res.writeHead(200, {
      "Content-Type": OUTPUT_MIME[job.format],
      "Content-Disposition": `attachment; filename="output.${job.format}"`,
      "Content-Length": buf.length,
    });
    res.end(buf);
  } catch (err) {
    // A failed compile is nearly always a fault in the pasted script, so the
    // message and the stack go back to the page, where they are shown as-is.
    console.error(`[crepuscule-slog] compile error (${job.format}):`, err.message);
    const status = err.code === "CREPUSCULE_SLOG_TIMEOUT" ? 504 : 500;
    sendJson(res, status, { error: err.message, stack: err.scriptStack || err.stack });
  }
}

function serveStatic(req, res) {
  const filePath = resolveStaticPath(req.url, PUBLIC_DIR);
  if (!filePath) {
    res.writeHead(403, { "Content-Type": "text/plain" });
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" });
      res.end("Not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": mimeFor(filePath),
      "Content-Length": data.length,
    });
    res.end(req.method === "HEAD" ? undefined : data);
  });
}

// Read the body, giving up once it passes the limit rather than after.
//
// Past the limit the request is refused immediately but still read to its
// end, with every further chunk dropped on the floor. The point of the limit
// is to bound memory, and that is met by letting go of what was buffered;
// closing the socket underneath a client that is still uploading only earns
// it a connection reset in place of the explanation.
function readBody(req, limit) {
  return new Promise((resolve, reject) => {
    let chunks = [];
    let total = 0;
    let refused = false;
    req.on("data", (chunk) => {
      if (refused) return;
      total += chunk.length;
      if (total > limit) {
        refused = true;
        chunks = [];
        const err = new Error("body too large");
        err.code = "CREPUSCULE_SLOG_BODY_TOO_LARGE";
        reject(err);
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => {
      if (!refused) resolve(Buffer.concat(chunks).toString("utf8"));
    });
    req.on("error", reject);
  });
}

function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  const headers = {
    "Content-Type": "application/json",
    "Content-Length": Buffer.byteLength(body),
  };
  res.writeHead(status, headers);
  res.end(body);
}

module.exports = { createCrepusculeSlogServer, PORT, HOST, PUBLIC_DIR };

// Only listen when started directly, so the tests can use a port of their own.
if (require.main === module) {
  createCrepusculeSlogServer().listen(PORT, HOST, () => {
    console.log(`\nCrepuscule Slog server running at http://localhost:${PORT}`);
    console.log(`Open that URL in your browser.\n`);
  });
}
