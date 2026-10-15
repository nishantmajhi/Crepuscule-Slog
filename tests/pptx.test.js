"use strict";

// The PPTX compiler. pptxgenjs has no buffer-returning writeFile, so output
// has to be intercepted; the question each test here asks is whether the
// interception belongs to one compilation or to the whole process.

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const { compilePptx } = require("../src/compile");
const { readZip, entryText } = require("./zip");

function tempTarget(tag) {
  const target = path.join(os.tmpdir(), `crepuscule-slog-pptx-${tag}-${process.pid}.pptx`);
  try {
    fs.unlinkSync(target);
  } catch {}
  return target;
}

function slideText(buf, n = 1) {
  assert.ok(Buffer.isBuffer(buf), "compiler must return a Buffer");
  assert.strictEqual(buf.subarray(0, 2).toString(), "PK", "not a zip");
  const files = readZip(buf);
  assert.ok(files[`ppt/slides/slide${n}.xml`], `archive has no slide ${n}`);
  return entryText(files, `ppt/slides/slide${n}.xml`);
}

function slideCount(buf) {
  return Object.keys(readZip(buf)).filter((n) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(n),
  ).length;
}

// Pattern A — writeFile to a path that does not exist here.
test("pattern A: pptx.writeFile to an absent path", async () => {
  const target = tempTarget("patternA");
  const buf = await compilePptx(`
const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
const slide = pptx.addSlide();
slide.addText('Kickoff', { x: 1, y: 1, fontSize: 32 });
await pptx.writeFile({ fileName: ${JSON.stringify(target)} });
`);
  assert.strictEqual(fs.existsSync(target), false, `${target} was really written`);
  assert.match(slideText(buf), /Kickoff/);
});

// Pattern B — the presentation as a bare trailing expression.
test("pattern B: bare `pptx;` as the last expression", async () => {
  const buf = await compilePptx(`
const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.addSlide().addText('Bare expression slide', { x: 1, y: 1 });
pptx;
`);
  assert.match(slideText(buf), /Bare expression slide/);
});

// Pattern C — write() with a nodebuffer, then fs.writeFileSync.
test("pattern C: pptx.write then fs.writeFileSync", async () => {
  const buf = await compilePptx(`
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const pptx = new PptxGenJS();
pptx.addSlide().addText('Written through fs', { x: 1, y: 1 });
const data = await pptx.write({ outputType: 'nodebuffer' });
fs.writeFileSync('/home/claude/output.pptx', data);
`);
  assert.match(slideText(buf), /Written through fs/);
});

test("the injected PptxGenJS global is intercepted too", async () => {
  const target = tempTarget("global");
  const buf = await compilePptx(`
const pptx = new PptxGenJS();
pptx.addSlide().addText('Global constructor', { x: 1, y: 1 });
await pptx.writeFile({ fileName: ${JSON.stringify(target)} });
`);
  assert.strictEqual(fs.existsSync(target), false, `${target} was really written`);
  assert.match(slideText(buf), /Global constructor/);
});

test("multiple slides all arrive", async () => {
  const buf = await compilePptx(`
const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
['One', 'Two', 'Three'].forEach((t, i) => {
  const s = pptx.addSlide();
  s.addText(t, { x: 1, y: 1 + i });
});
pptx;
`);
  assert.strictEqual(slideCount(buf), 3);
  assert.match(slideText(buf, 3), /Three/);
});

// Two compilations overlapping in one process must not see each other's
// interception. The presentations here are deliberately different sizes and
// write at different moments, so a swap is unmistakable.
test("concurrent compilations keep their own output", async () => {
  const bigTarget = tempTarget("big");
  const smallTarget = tempTarget("small");

  const script = (slides, delayMs, label, target) => `
const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
for (let i = 0; i < ${slides}; i++) {
  pptx.addSlide().addText('${label} ' + i, { x: 1, y: 1 });
}
await new Promise(r => setTimeout(r, ${delayMs}));
await pptx.writeFile({ fileName: ${JSON.stringify(target)} });
`;

  const [big, small] = await Promise.all([
    compilePptx(script(12, 150, "BIG", bigTarget)),
    compilePptx(script(1, 450, "SMALL", smallTarget)),
  ]);

  assert.strictEqual(slideCount(big), 12, "the 12-slide compile got the wrong bytes");
  assert.strictEqual(slideCount(small), 1, "the 1-slide compile got the wrong bytes");
  assert.match(slideText(small), /SMALL 0/);
  assert.strictEqual(fs.existsSync(bigTarget), false, `${bigTarget} was really written`);
  assert.strictEqual(fs.existsSync(smallTarget), false, `${smallTarget} was really written`);
});

// Whatever the interception does to pptxgenjs, it must leave the library as it
// found it: a later import in the host process is not Crepuscule Slog's to modify.
test("the pptxgenjs prototype is not left modified", async () => {
  const PptxGenJS = require("pptxgenjs");
  const before = {
    write: PptxGenJS.prototype.write,
    writeFile: PptxGenJS.prototype.writeFile,
  };
  await compilePptx(`
const pptx = new PptxGenJS();
pptx.addSlide().addText('x', { x: 1, y: 1 });
pptx;
`);
  assert.strictEqual(PptxGenJS.prototype.write, before.write);
  assert.strictEqual(PptxGenJS.prototype.writeFile, before.writeFile);
});

test("no output produces a message naming the three endings", async () => {
  await assert.rejects(
    () => compilePptx("const x = 1;\n", { drainMs: 200 }),
    /produced no output/,
  );
});

test("a thrown error reaches the caller unchanged", async () => {
  await assert.rejects(
    () => compilePptx("throw new Error('deck exploded');\n"),
    /deck exploded/,
  );
});
