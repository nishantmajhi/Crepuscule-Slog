"use strict";

// The runner is what stands between a pasted script and the server's ability
// to answer the next request. Its job is to come back — with a document, with
// the script's error, or with a refusal — but always to come back.

const test = require("node:test");
const assert = require("node:assert");

const { createRunner } = require("../src/runner");
const { readZip, entryText } = require("./zip");

test("compiles through a worker thread", async (t) => {
  const runner = createRunner();
  t.after(() => runner.shutdown());

  const buf = await runner.run(
    "docx",
    `
const { Document, Paragraph } = require('docx');
const doc = new Document({ sections: [{ children: [new Paragraph('Through a thread')] }] });
doc;
`,
  );
  assert.strictEqual(buf.subarray(0, 2).toString(), "PK");
  assert.match(entryText(readZip(buf), "word/document.xml"), /Through a thread/);
});

// The whole point. Before this existed, one such paste ended the server: the
// event loop never came back and every later request hung.
test("a script that loops forever is stopped, and the next one still works", async (t) => {
  const runner = createRunner({ timeoutMs: 3000 });
  t.after(() => runner.shutdown());

  await assert.rejects(
    () => runner.run("docx", "while (true) {}\n"),
    (err) => {
      assert.strictEqual(err.code, "CREPUSCULE_SLOG_TIMEOUT");
      assert.match(err.message, /did not finish within 3 seconds/);
      return true;
    },
  );

  const buf = await runner.run(
    "docx",
    `
const { Document, Paragraph } = require('docx');
const doc = new Document({ sections: [{ children: [new Paragraph('Still here')] }] });
doc;
`,
  );
  assert.match(entryText(readZip(buf), "word/document.xml"), /Still here/);
});

test("a script that awaits something that never arrives is stopped too", async (t) => {
  const runner = createRunner({ timeoutMs: 2000 });
  t.after(() => runner.shutdown());

  await assert.rejects(
    () => runner.run("docx", "await new Promise(() => {});\n"),
    /did not finish within 2 seconds/,
  );
});

test("the script's own error is passed back", async (t) => {
  const runner = createRunner();
  t.after(() => runner.shutdown());

  await assert.rejects(
    () => runner.run("pptx", "throw new Error('no slides for you');\n"),
    /no slides for you/,
  );
});

test("an unknown format is refused", async (t) => {
  const runner = createRunner();
  t.after(() => runner.shutdown());

  await assert.rejects(() => runner.run("rtf", "x;\n"), /Unknown format/);
});

// Threads are not shared, so nothing one script does can reach another.
test("concurrent compilations are independent", async (t) => {
  const runner = createRunner();
  t.after(() => runner.shutdown());

  const deck = (slides, label) => `
const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
for (let i = 0; i < ${slides}; i++) pptx.addSlide().addText('${label}', { x: 1, y: 1 });
pptx;
`;
  const count = (buf) =>
    Object.keys(readZip(buf)).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;

  const [big, small] = await Promise.all([
    runner.run("pptx", deck(9, "BIG")),
    runner.run("pptx", deck(2, "SMALL")),
  ]);
  assert.strictEqual(count(big), 9);
  assert.strictEqual(count(small), 2);
});

// A script can patch the libraries, or leave timers behind, or corrupt its own
// heap. None of that may survive into the next compilation.
test("a script cannot affect the next compilation", async (t) => {
  const runner = createRunner();
  t.after(() => runner.shutdown());

  await assert.rejects(
    () =>
      runner.run(
        "docx",
        `
const docx = require('docx');
docx.Packer.toBuffer = () => { throw new Error('sabotaged'); };
setInterval(() => {}, 10);
throw new Error('done meddling');
`,
      ),
    /done meddling/,
  );

  const buf = await runner.run(
    "docx",
    `
const { Document, Packer, Paragraph } = require('docx');
const fs = require('fs');
const doc = new Document({ sections: [{ children: [new Paragraph('Unsabotaged')] }] });
fs.writeFileSync('/home/claude/out.docx', await Packer.toBuffer(doc));
`,
  );
  assert.match(entryText(readZip(buf), "word/document.xml"), /Unsabotaged/);
});
