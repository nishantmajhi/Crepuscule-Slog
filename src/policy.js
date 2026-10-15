"use strict";

// The decisions the server makes about a request, with no sockets in them.
//
// Crepuscule Slog runs arbitrary JavaScript on demand. That is the whole point of it,
// and it is safe only for as long as the one asking is the person sitting at
// the machine. Everything in this file exists to keep that true.

const path = require("path");

// A paste is source code. Eight megabytes is more than any script Claude
// writes and small enough that a stray upload cannot exhaust the heap.
const MAX_BODY_BYTES = 8 * 1024 * 1024;

const FORMATS = ["docx", "xlsx", "pptx"];

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".ico": "image/x-icon",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".json": "application/json",
};

const OUTPUT_MIME = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

const LOOPBACK_HOSTS = new Set(["localhost", "127.0.0.1", "[::1]", "::1"]);

// The page is served by this same process, so the only origin with any
// business posting to /compile is this server's own. It is compared against
// the Host header rather than a configured port, so it holds whichever port
// the server was started on — and localhost and 127.0.0.1 count as the same
// place, since the user may have typed either.
function isOwnOrigin(origin, hostHeader) {
  let url;
  try {
    url = new URL(origin);
  } catch {
    return false;
  }
  if (url.protocol !== "http:") return false;
  if (!LOOPBACK_HOSTS.has(url.hostname) && !LOOPBACK_HOSTS.has(`[${url.hostname}]`)) {
    return false;
  }
  const originPort = url.port || "80";
  const hostPort = String(hostHeader || "").split(":").pop();
  return originPort === (hostPort && /^\d+$/.test(hostPort) ? hostPort : "80");
}

// Is this POST /compile allowed to run code on this machine?
//
// Binding to loopback is not enough on its own. Any web page open in the
// user's browser can POST to http://localhost:3000 — and a POST with
// Content-Type: text/plain is a "simple request", which needs no preflight,
// so nothing stops it being sent. The page cannot read the reply, but it does
// not need to: the code has already run. Two checks close that door. A
// browser attaches Origin to every cross-origin POST, and a page cannot forge
// or omit it; and demanding JSON puts the request outside the simple-request
// rules, so the browser must ask permission first and will be refused.
function checkCompileRequest(headers) {
  const origin = headerValue(headers, "origin");
  if (origin && !isOwnOrigin(origin, headerValue(headers, "host"))) {
    return {
      ok: false,
      status: 403,
      error:
        "Refused: this request came from another site. Crepuscule Slog compiles code, " +
        "so it only accepts requests from the page it serves itself.",
    };
  }

  const contentType = headerValue(headers, "content-type") || "";
  if (!/^application\/json\b/i.test(contentType.trim())) {
    return {
      ok: false,
      status: 415,
      error: "Refused: /compile expects Content-Type: application/json.",
    };
  }

  return { ok: true };
}

// What the body has to contain to be a compile job.
function parseCompileBody(text) {
  let body;
  try {
    body = JSON.parse(text);
  } catch {
    return { ok: false, status: 400, error: "Invalid JSON body." };
  }
  if (!body || typeof body !== "object") {
    return { ok: false, status: 400, error: "Invalid JSON body." };
  }

  const { format, code } = body;
  if (!format || !code) {
    return { ok: false, status: 400, error: "Missing format or code." };
  }
  if (typeof format !== "string" || typeof code !== "string") {
    return { ok: false, status: 400, error: "format and code must both be strings." };
  }
  if (!FORMATS.includes(format)) {
    return { ok: false, status: 400, error: `Unknown format: ${format}` };
  }

  return { ok: true, format, code };
}

// Turn a request URL into a file inside the static directory, or nothing.
//
// The directory is the one holding the page, not the project root: serving
// the root meant /.git/config, /package.json and everything under
// node_modules were downloadable from the same server.
function resolveStaticPath(requestUrl, rootDir) {
  let urlPath = String(requestUrl).split("?")[0].split("#")[0];
  try {
    urlPath = decodeURIComponent(urlPath);
  } catch {
    return null; // malformed percent-encoding
  }
  if (urlPath.includes("\0")) return null;
  if (urlPath === "/" || urlPath === "") urlPath = "/index.html";
  if (!urlPath.startsWith("/")) return null;

  // Backslashes are path separators on Windows but not in URLs, so a request
  // for /..\..\secret must not become a traversal once path.resolve sees it.
  if (urlPath.includes("\\")) return null;

  const root = path.resolve(rootDir);
  const resolved = path.resolve(root, `.${urlPath}`);
  if (resolved !== root && !resolved.startsWith(root + path.sep)) return null;
  return resolved;
}

function mimeFor(filePath) {
  return MIME[path.extname(filePath).toLowerCase()] || "application/octet-stream";
}

function headerValue(headers, name) {
  if (!headers) return undefined;
  const value = headers[name] ?? headers[name.toLowerCase()];
  return Array.isArray(value) ? value[0] : value;
}

module.exports = {
  MAX_BODY_BYTES,
  FORMATS,
  OUTPUT_MIME,
  isOwnOrigin,
  checkCompileRequest,
  parseCompileBody,
  resolveStaticPath,
  mimeFor,
};
