"use strict";

// The DOCX compiler, driven with the script shapes Claude's docx skill
// actually emits. Each test asserts on the contents of the produced archive,
// not just its size, because an empty-but-valid .docx is the failure mode
// worth catching.

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const { compileDocx } = require("../src/compile");
const { readZip, entryText } = require("./zip");

function assertIsDocx(buf) {
  assert.ok(Buffer.isBuffer(buf), "compiler must return a Buffer");
  assert.strictEqual(buf.subarray(0, 2).toString(), "PK", "not a zip");
  const files = readZip(buf);
  assert.ok(files["word/document.xml"], "archive has no word/document.xml");
  assert.ok(files["[Content_Types].xml"], "archive has no content types part");
  return entryText(files, "word/document.xml");
}

// Pattern A — the most common ending from the docx skill: a floating promise
// whose callback writes to a path that does not exist on this machine.
test("pattern A: Packer.toBuffer().then(fs.writeFileSync)", async () => {
  const xml = assertIsDocx(
    await compileDocx(`
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require('docx');
const fs = require('fs');

const doc = new Document({
  sections: [{
    children: [
      new Paragraph({ text: 'Quarterly Report', heading: HeadingLevel.HEADING_1 }),
      new Paragraph({ children: [new TextRun({ text: 'Revenue rose.', bold: true })] }),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/output.docx', buf);
});
`),
  );
  assert.match(xml, /Quarterly Report/);
  assert.match(xml, /Revenue rose\./);
});

// Pattern B — the document as a bare trailing expression.
test("pattern B: bare `doc;` as the last expression", async () => {
  const xml = assertIsDocx(
    await compileDocx(`
const { Document, Paragraph } = require('docx');
const doc = new Document({
  sections: [{ children: [new Paragraph('Bare expression body')] }],
});
doc;
`),
  );
  assert.match(xml, /Bare expression body/);
});

// The bare-expression script has already produced its document by the time the
// script ends; waiting ten seconds for a write that is never coming made the
// commonest paste the slowest one.
test("pattern B returns promptly rather than waiting out the drain timeout", async () => {
  const started = Date.now();
  await compileDocx(`
const { Document, Paragraph } = require('docx');
const doc = new Document({ sections: [{ children: [new Paragraph('x')] }] });
doc;
`);
  const elapsed = Date.now() - started;
  assert.ok(elapsed < 2000, `took ${elapsed} ms; a returned document needs no wait`);
});

// Pattern C — await the buffer, then write it.
test("pattern C: await Packer.toBuffer then fs.writeFileSync", async () => {
  const xml = assertIsDocx(
    await compileDocx(`
const { Document, Packer, Paragraph } = require('docx');
const fs = require('fs');
const doc = new Document({ sections: [{ children: [new Paragraph('Awaited body')] }] });
const buffer = await Packer.toBuffer(doc);
fs.writeFileSync('/home/claude/output.docx', buffer);
`),
    );
  assert.match(xml, /Awaited body/);
});

// The docx exports are also injected as globals, because some scripts skip the
// require and use `Document` directly.
test("the docx exports are available as globals without require", async () => {
  const xml = assertIsDocx(
    await compileDocx(`
const doc = new Document({ sections: [{ children: [new Paragraph('No require here')] }] });
doc;
`),
  );
  assert.match(xml, /No require here/);
});

test("a returned buffer is accepted directly", async () => {
  const xml = assertIsDocx(
    await compileDocx(`
const { Document, Packer, Paragraph } = require('docx');
const doc = new Document({ sections: [{ children: [new Paragraph('Returned buffer')] }] });
await Packer.toBuffer(doc);
`),
  );
  assert.match(xml, /Returned buffer/);
});

// A long argument list split across lines leaves `);` as the final line. The
// return-injection must not turn that into a syntax error.
test("a multi-line fs.writeFileSync call as the last statement", async () => {
  const xml = assertIsDocx(
    await compileDocx(`
const { Document, Packer, Paragraph } = require('docx');
const fs = require('fs');
const doc = new Document({ sections: [{ children: [new Paragraph('Split call')] }] });
const buf = await Packer.toBuffer(doc);
fs.writeFileSync(
  '/home/claude/output.docx',
  buf
);
`),
  );
  assert.match(xml, /Split call/);
});

test("fs.writeFile with a callback is intercepted and the callback runs", async () => {
  const xml = assertIsDocx(
    await compileDocx(`
const { Document, Packer, Paragraph } = require('docx');
const fs = require('fs');
const doc = new Document({ sections: [{ children: [new Paragraph('Callback form')] }] });
const buf = await Packer.toBuffer(doc);
let called = false;
fs.writeFile('/home/claude/output.docx', buf, () => { called = true; });
if (!called) throw new Error('callback was not invoked');
doc;
`),
  );
  assert.match(xml, /Callback form/);
});

test("fs.promises.writeFile is intercepted as well", async () => {
  const xml = assertIsDocx(
    await compileDocx(`
const { Document, Packer, Paragraph } = require('docx');
const fs = require('fs');
const doc = new Document({ sections: [{ children: [new Paragraph('Promise form')] }] });
await fs.promises.writeFile('/home/claude/output.docx', await Packer.toBuffer(doc));
`),
  );
  assert.match(xml, /Promise form/);
});

test("no output produces a message naming the three endings", async () => {
  await assert.rejects(
    // A short drain: the wait only exists for output that might still be
    // coming, and this script has none coming.
    () => compileDocx("const x = 1;\n", { drainMs: 200 }),
    /produced no output/,
  );
});

test("a thrown error reaches the caller unchanged", async () => {
  await assert.rejects(
    () => compileDocx("throw new Error('deliberate failure');\n"),
    /deliberate failure/,
  );
});

// The path a script writes to is meaningless here, so it must never reach the
// real filesystem — that interception is the reason the project exists.
test("nothing is written to the real filesystem", async () => {
  const target = path.join(os.tmpdir(), `crepuscule-slog-docx-escape-${process.pid}.docx`);
  try {
    fs.unlinkSync(target);
  } catch {}
  await compileDocx(`
const { Document, Packer, Paragraph } = require('docx');
const fs = require('fs');
const doc = new Document({ sections: [{ children: [new Paragraph('x')] }] });
const buf = await Packer.toBuffer(doc);
fs.writeFileSync(${JSON.stringify(target)}, buf);
`);
  assert.strictEqual(fs.existsSync(target), false, `${target} was really written`);
});

test("require of an unknown module explains what is available", async () => {
  await assert.rejects(
    () => compileDocx("require('axios');\n"),
    (err) => {
      assert.match(err.message, /axios/);
      assert.match(err.message, /docx/);
      // "sandbox" claims containment this does not provide; the accurate word
      // for what the list is, is an allowlist.
      assert.doesNotMatch(err.message, /sandbox/i);
      return true;
    },
  );
});

// `modules` is a plain object, so an inherited key used to resolve instead of
// throwing: require('constructor') handed back Object.prototype.constructor.
test("require of an inherited Object key throws rather than resolving", async () => {
  await assert.rejects(
    () => compileDocx("require('constructor');\n"),
    /not available/,
  );
  await assert.rejects(
    () => compileDocx("require('hasOwnProperty');\n"),
    /not available/,
  );
});

test("fake fs.readFileSync says plainly that it is unsupported", async () => {
  await assert.rejects(
    () => compileDocx("require('fs').readFileSync('/etc/passwd');\n"),
    /not supported in Crepuscule Slog/,
  );
});
