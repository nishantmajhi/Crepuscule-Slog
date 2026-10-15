"use strict";

// Running the pasted script, with no HTTP in sight.
//
// Crepuscule Slog's trick lives here: a script written for a machine that does not exist
// — it writes to /home/claude/output.docx — is run against a fake `fs` that
// keeps the bytes and throws the path away. Everything below is about
// accepting the several shapes such a script can take and getting one
// document out of it.
//
// This module is *not* a security boundary and does not pretend to be one.
// It compiles the pasted text into a function in this realm and calls it, so
// the code runs with exactly the privileges of the process. What the module
// provides is convenience: the libraries pre-loaded, a fake filesystem, and
// output collected whichever way the script chooses to produce it. Isolation,
// where it matters, is the caller's job — src/runner.js runs each compilation
// in a worker thread it can terminate.

const vm = require("vm");
const path = require("path");

// Pre-load the three libraries once. These are exactly what Claude's skill
// files install and require.
const docxLib = require("docx");
const ExcelJS = require("exceljs");
const PptxGenJS = require("pptxgenjs");

// How long to keep waiting after the script itself has finished, for output
// that a floating promise has not delivered yet. Only reached when the script
// has produced nothing usable, so a well-formed paste never pays it.
const DRAIN_MS = 10000;

// ─── DOCX ────────────────────────────────────────────────────────────────────
async function compileDocx(code, options = {}) {
  let captured = null;
  const fakeFs = makeFakeFs((buf) => {
    captured = toBuffer(buf);
  });

  // A finished document can arrive as a Document to pack, as packed bytes, or
  // through the fake fs. Any of the three means there is nothing left to wait
  // for.
  const usable = (result) =>
    captured !== null ||
    result instanceof docxLib.Document ||
    result instanceof Uint8Array ||
    Buffer.isBuffer(result);

  const result = await runPastedScript(
    code,
    {
      require: makeRequire({ docx: docxLib }, fakeFs),
      fs: fakeFs,
      // Spread the docx exports as globals as well, so scripts that destructure
      // from require() and scripts that just use `Document` both work.
      ...docxLib,
    },
    { hasOutput: usable, drainMs: options.drainMs },
  );

  if (result instanceof docxLib.Document) {
    return toBuffer(await docxLib.Packer.toBuffer(result));
  }
  if (result instanceof Uint8Array || Buffer.isBuffer(result)) {
    return toBuffer(result);
  }
  if (captured) return captured;

  throw new Error(
    "Script finished but produced no output. " +
      "Make sure it ends with `doc;`, calls `Packer.toBuffer(doc).then(...)`, " +
      "or uses `fs.writeFileSync`.",
  );
}

// ─── XLSX ────────────────────────────────────────────────────────────────────
async function compileXlsx(code, options = {}) {
  let captured = null;
  const fakeFs = makeFakeFs((buf) => {
    captured = toBuffer(buf);
  });

  // exceljs writes to disk itself rather than handing back bytes, so every
  // workbook the script can get hold of has its writeFile replaced.
  const intercept = (wb) => {
    wb.xlsx.writeFile = async (_path) => {
      const buf = await wb.xlsx.writeBuffer();
      captured = toBuffer(buf);
      return buf;
    };
    return wb;
  };

  class CapturingWorkbook extends ExcelJS.Workbook {
    constructor(...args) {
      super(...args);
      intercept(this);
    }
  }

  // The spread has to come first: `{ ...ExcelJS, Workbook: Capturing }`. The
  // other order puts the library's own Workbook back on top and the script
  // ends up with an unintercepted one, which writes to the real disk.
  const exceljsModule = { ...ExcelJS, Workbook: CapturingWorkbook };

  // A ready-made workbook, matching the skill's pattern of using a `workbook`
  // that is already there.
  const workbook = intercept(new ExcelJS.Workbook());

  const usable = (result) => captured !== null || isWorkbook(result);

  const result = await runPastedScript(
    code,
    {
      require: makeRequire({ exceljs: exceljsModule }, fakeFs),
      ExcelJS: exceljsModule,
      workbook,
      fs: fakeFs,
    },
    { hasOutput: usable, drainMs: options.drainMs },
  );

  // Precedence matters. The workbook the script handed back is the one it
  // meant; the provided global is only right if the script actually used it.
  // Returning the global unconditionally produced a valid file with no rows
  // in it whenever a script built its own workbook.
  if (isWorkbook(result)) return toBuffer(await result.xlsx.writeBuffer());
  if (captured) return captured;
  if (workbook.worksheets.length > 0) {
    return toBuffer(await workbook.xlsx.writeBuffer());
  }

  throw new Error(
    "Script finished but produced no output — no worksheets were added. " +
      "End your script with the workbook variable (e.g. `wb;`) or call " +
      "`workbook.xlsx.writeFile(...)`.",
  );
}

function isWorkbook(value) {
  return Boolean(value && value.xlsx && typeof value.xlsx.writeBuffer === "function");
}

// ─── PPTX ────────────────────────────────────────────────────────────────────
async function compilePptx(code, options = {}) {
  let captured = null;
  const fakeFs = makeFakeFs((buf) => {
    captured = toBuffer(buf);
  });

  // Interception by subclass rather than by patching PptxGenJS.prototype.
  // Patching the prototype patches it for the whole process: two compilations
  // overlapping in one process installed their capture callbacks over each
  // other, so one request was answered with the other's presentation and the
  // loser's writeFile fell through to the real filesystem.
  class CapturingPptx extends PptxGenJS {
    async write(opts) {
      const props = opts && typeof opts === "object" ? { ...opts } : {};
      props.outputType = "nodebuffer";
      const buf = await super.write(props);
      captured = toBuffer(buf);
      return buf;
    }

    async writeFile(_opts) {
      // The path is meaningless here; the bytes are the point.
      const buf = await super.write({ outputType: "nodebuffer" });
      captured = toBuffer(buf);
      return buf;
    }
  }

  const usable = (result) =>
    captured !== null || Boolean(result && typeof result.write === "function");

  const result = await runPastedScript(
    code,
    {
      require: makeRequire({ pptxgenjs: CapturingPptx }, fakeFs),
      PptxGenJS: CapturingPptx,
      fs: fakeFs,
    },
    { hasOutput: usable, drainMs: options.drainMs },
  );

  // A presentation returned without ever being written still needs packing.
  if (!captured && result && typeof result.write === "function") {
    captured = toBuffer(await result.write({ outputType: "nodebuffer" }));
  }

  if (captured) return captured;

  throw new Error(
    "Script finished but produced no output. " +
      "Make sure it calls `pptx.writeFile(...)`, `pptx.write(...)`, or ends with `pptx;`.",
  );
}

const COMPILERS = {
  docx: compileDocx,
  xlsx: compileXlsx,
  pptx: compilePptx,
};

function compile(format, code, options) {
  const compiler = COMPILERS[format];
  if (!compiler) return Promise.reject(new Error(`Unknown format: ${format}`));
  return compiler(code, options);
}

// ─── Running the script ──────────────────────────────────────────────────────

// Compile the pasted text into a function of this realm and call it with the
// supplied globals as arguments.
//
// It used to run in a vm context, which looked tidier but was quietly wrong:
// a vm context has its own intrinsics, so an array literal written inside it
// is not an instance of this realm's Array. exceljs decides how to read a row
// with `value instanceof Array`, so every `sheet.addRow(['a', 1])` fell down
// the object branch and added an empty row — producing a spreadsheet with the
// right sheet names, the right formatting and none of the data. Compiling in
// this realm makes the pasted code's objects the same objects the libraries
// expect. It gives up nothing in isolation, because a vm context was never
// providing any.
async function runPastedScript(code, globals, options = {}) {
  const drainMs = options.drainMs === undefined ? DRAIN_MS : options.drainMs;
  const hasOutput = options.hasOutput || (() => true);

  const timers = new TimerSet();
  const names = [];
  const values = [];

  const provided = {
    console,
    Buffer,
    setTimeout: timers.setTimeout,
    clearTimeout: timers.clearTimeout,
    setInterval: timers.setInterval,
    clearInterval: timers.clearInterval,
    Promise,
    ...globals,
  };
  for (const name of Object.keys(provided)) {
    if (!IDENTIFIER.test(name) || RESERVED.has(name)) continue;
    names.push(name);
    values.push(provided[name]);
  }
  // Shadow the host globals a pasted script has no business reaching for by
  // accident. This is tidiness, not containment — see the note at the top.
  for (const name of SHADOWED) {
    if (!names.includes(name)) {
      names.push(name);
      values.push(undefined);
    }
  }

  const fn = compilePasted(code, names);

  try {
    let result = fn(...values);
    if (result && typeof result.then === "function") result = await result;

    // A script may hand its output to a promise nobody awaited —
    // `Packer.toBuffer(doc).then(buf => fs.writeFileSync(...))` is the common
    // shape — so if nothing usable has appeared yet, give that callback time
    // to run. A script that returned its document does not wait at all.
    if (!hasOutput(result) && drainMs > 0) {
      await waitUntil(() => hasOutput(result), drainMs);
    }
    return result;
  } finally {
    // Anything the script scheduled and did not cancel would otherwise keep
    // running for the life of the process.
    timers.clearAll();
  }
}

// Build the callable. The return injection is a guess about the last line, so
// if the guess does not parse, fall back to the script as pasted — and if
// that does not parse either, report that error, which is the user's.
function compilePasted(code, names) {
  const wrap = (body) => `return (async () => {\n${body}\n})();`;
  const options = { filename: "pasted-script.js" };
  try {
    return vm.compileFunction(wrap(injectReturn(code)), names, options);
  } catch (injectedError) {
    try {
      return vm.compileFunction(wrap(code), names, options);
    } catch {
      throw injectedError;
    }
  }
}

function waitUntil(done, timeoutMs) {
  return new Promise((resolve) => {
    const deadline = Date.now() + timeoutMs;
    const check = () => {
      if (done() || Date.now() >= deadline) resolve();
      else setTimeout(check, 10).unref?.();
    };
    // Let the microtask queue run first: output is usually one tick away.
    Promise.resolve().then(check);
  });
}

// Timers handed to the pasted script, remembered so they can be cancelled when
// the compilation is over.
class TimerSet {
  constructor() {
    this.handles = new Set();
    this.setTimeout = (...args) => this.track(setTimeout(...args));
    this.setInterval = (...args) => this.track(setInterval(...args));
    this.clearTimeout = (h) => {
      this.handles.delete(h);
      clearTimeout(h);
    };
    this.clearInterval = (h) => {
      this.handles.delete(h);
      clearInterval(h);
    };
  }

  track(handle) {
    this.handles.add(handle);
    return handle;
  }

  clearAll() {
    for (const handle of this.handles) {
      clearTimeout(handle);
      clearInterval(handle);
    }
    this.handles.clear();
  }
}

// ─── Return injection ────────────────────────────────────────────────────────

const IDENTIFIER = /^[A-Za-z_$][A-Za-z0-9_$]*$/;

// Names that cannot be function parameters, and names it would be actively
// unhelpful to shadow.
const RESERVED = new Set(["eval", "arguments", "this"]);

const SHADOWED = ["process", "global", "globalThis", "module", "exports", "__dirname", "__filename"];

// A line that begins a statement which cannot be returned.
const NOT_AN_EXPRESSION =
  /^(const|let|var|function|class|if|else|for|while|do|try|catch|finally|switch|case|default|return|throw|break|continue|debugger|import|export|with)\b/;

// Inject `return` before the last bare expression, so a script ending in
// `doc;` hands the document back.
function injectReturn(code) {
  const lines = code.split("\n");
  for (let i = lines.length - 1; i >= 0; i--) {
    const trimmed = lines[i].trim();
    // Skip blank lines and trailing comments.
    if (
      trimmed === "" ||
      trimmed.startsWith("//") ||
      trimmed.startsWith("/*") ||
      trimmed.startsWith("*")
    ) {
      continue;
    }
    if (NOT_AN_EXPRESSION.test(trimmed)) break;
    // A line starting with a closing delimiter is the tail of a statement that
    // began further up — a multi-line call, array or `.then()` block. Putting
    // `return` in front of it is a syntax error.
    if (/^[)\]}]/.test(trimmed)) break;

    // Splice at the index rather than using String.replace, whose replacement
    // string treats $&, $` and $' as back-references: a last line containing
    // one came out corrupted.
    const at = lines[i].indexOf(trimmed);
    lines[i] = `${lines[i].slice(0, at)}return ${lines[i].slice(at)}`;
    break;
  }
  return lines.join("\n");
}

// ─── The fake world the script runs in ───────────────────────────────────────

// Only the three packages Claude's skills use resolve; everything else gets a
// message saying so. The table has a null prototype because a plain object
// resolved its inherited keys — require('constructor') handed back
// Object.prototype.constructor instead of throwing.
function makeRequire(modules, fakeFs) {
  const table = Object.assign(Object.create(null), modules);
  return function require(id) {
    if (typeof id === "string" && id in table) return table[id];
    // Some scripts require('fs') explicitly rather than using the global.
    if (id === "fs") return fakeFs;
    // path is only ever used to build an output path, which is discarded.
    if (id === "path") return path;
    throw new Error(
      `require('${id}') is not available in Crepuscule Slog's module allowlist. ` +
        `Only 'docx', 'exceljs', and 'pptxgenjs' are pre-installed.`,
    );
  };
}

// Intercepts writeFileSync and writeFile, keeps the bytes, discards the path —
// which is usually /home/claude/something and means nothing here.
function makeFakeFs(onCapture) {
  return {
    writeFileSync(_filePath, data, _opts) {
      onCapture(data);
    },
    writeFile(_filePath, data, optsOrCb, maybeCb) {
      onCapture(data);
      const cb = typeof optsOrCb === "function" ? optsOrCb : maybeCb;
      if (typeof cb === "function") cb(null);
    },
    promises: {
      async writeFile(_filePath, data) {
        onCapture(data);
      },
      async mkdir() {},
    },
    // Stubs for anything else fs-shaped a script might call on its way to
    // writing the file.
    mkdirSync() {},
    existsSync() {
      return true;
    },
    readFileSync(p) {
      throw new Error(`readFileSync('${p}') is not supported in Crepuscule Slog.`);
    },
  };
}

function toBuffer(data) {
  if (Buffer.isBuffer(data)) return data;
  if (data instanceof Uint8Array) return Buffer.from(data);
  if (data instanceof ArrayBuffer) return Buffer.from(data);
  return Buffer.from(data);
}

module.exports = {
  COMPILERS,
  compile,
  compileDocx,
  compileXlsx,
  compilePptx,
  runPastedScript,
  injectReturn,
  makeRequire,
  makeFakeFs,
  toBuffer,
  DRAIN_MS,
};
