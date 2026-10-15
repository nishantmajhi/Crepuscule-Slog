"use strict";

// injectReturn turns a script whose last statement is a bare expression
// (`doc;`) into one that returns it, so the compiler can pick the value up.
// It is a text transform applied to code nobody here wrote, so the thing that
// matters is that it never turns a working script into a broken one.

const test = require("node:test");
const assert = require("node:assert");
const vm = require("node:vm");

const { injectReturn } = require("../src/compile");

// Every transformed script must still parse the way the compiler runs it.
function parses(code) {
  try {
    new vm.Script(`(async () => { ${injectReturn(code)} })()`);
    return true;
  } catch {
    return false;
  }
}

test("returns the last bare expression", () => {
  const out = injectReturn("const doc = 1;\ndoc;\n");
  assert.match(out, /return doc;/);
});

test("ignores trailing blank lines and comments", () => {
  const out = injectReturn("const doc = 1;\ndoc;\n// all done\n\n");
  assert.match(out, /return doc;/);
  assert.doesNotMatch(out, /return \/\//);
});

test("leaves a declaration alone", () => {
  const code = "const doc = makeIt();\n";
  assert.strictEqual(injectReturn(code), code);
});

test("leaves an explicit return alone", () => {
  const code = "return doc;\n";
  assert.strictEqual(injectReturn(code), code);
});

test("does not break a multi-line call as the last statement", () => {
  // Claude formats long argument lists across lines; the last line is then a
  // bare `);` which must not become `return );`.
  const code = [
    "const buf = await Packer.toBuffer(doc);",
    "fs.writeFileSync(",
    '  "/home/claude/out.docx",',
    "  buf",
    ");",
    "",
  ].join("\n");
  assert.ok(parses(code), "injected script must still parse");
});

test("does not break a script ending in a closing brace", () => {
  const code = ["function unused() {", "  return 1;", "}", ""].join("\n");
  assert.ok(parses(code));
});

test("does not break a script ending in a closing bracket", () => {
  const code = ["const rows = [", "  [1, 2],", "];", ""].join("\n");
  assert.ok(parses(code));
});

test("keeps a floating .then() block working", () => {
  const code = [
    "Packer.toBuffer(doc).then(buf => {",
    '  fs.writeFileSync("/home/claude/out.docx", buf);',
    "});",
    "",
  ].join("\n");
  assert.ok(parses(code));
});

test("does not corrupt a last line containing $ replacement patterns", () => {
  // String.prototype.replace treats $&, $` and $' in the replacement as
  // back-references. A last line holding one of those must survive verbatim.
  const code = 'const o = { "$&": doc };\no["$&"];\n';
  const out = injectReturn(code);
  assert.ok(parses(code), "injected script must still parse");
  assert.match(out, /return o\["\$&"\];/);
});

test("is a no-op on an empty script", () => {
  assert.strictEqual(injectReturn(""), "");
  assert.strictEqual(injectReturn("\n\n"), "\n\n");
});

test("preserves the original indentation of the injected line", () => {
  const out = injectReturn("if (x) {\n}\n  doc;\n");
  assert.match(out, /^ {2}return doc;$/m);
});
