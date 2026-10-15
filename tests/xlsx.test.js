"use strict";

// The XLSX compiler. Two things are easy to get wrong here and both produce a
// file that opens without complaint: returning a workbook that is not the one
// the script built, and letting exceljs write to the real disk because the
// interception was attached to an object the script never used.

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const ExcelJS = require("exceljs");

const { compileXlsx } = require("../src/compile");
const { readZip } = require("./zip");

function tempTarget(tag) {
  const target = path.join(os.tmpdir(), `crepuscule-slog-xlsx-${tag}-${process.pid}.xlsx`);
  try {
    fs.unlinkSync(target);
  } catch {}
  return target;
}

// Read the produced bytes back with exceljs, which is the only honest way to
// ask what is in a spreadsheet: sheet names live in a deflated part.
async function readBack(buf) {
  assert.ok(Buffer.isBuffer(buf), "compiler must return a Buffer");
  assert.strictEqual(buf.subarray(0, 2).toString(), "PK", "not a zip");
  const files = readZip(buf);
  assert.ok(files["xl/workbook.xml"], "archive has no xl/workbook.xml");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}

function cellValues(sheet) {
  const out = [];
  sheet.eachRow((row) => out.push(row.values.slice(1)));
  return out;
}

// Pattern A — the shape the exceljs skill emits: its own workbook, written to
// a path on a machine that is not this one.
test("pattern A: require('exceljs'), own workbook, writeFile to an absent path", async () => {
  const target = tempTarget("patternA");
  const wb = await readBack(
    await compileXlsx(`
const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet('Summary');
sheet.addRow(['Region', 'Revenue']);
sheet.addRow(['North', 42000]);
await wb.xlsx.writeFile(${JSON.stringify(target)});
`),
  );
  assert.strictEqual(fs.existsSync(target), false, `${target} was really written`);
  const sheet = wb.getWorksheet("Summary");
  assert.ok(sheet, `workbook has no Summary sheet, only: ${wb.worksheets.map((w) => w.name)}`);
  assert.deepStrictEqual(cellValues(sheet), [
    ["Region", "Revenue"],
    ["North", 42000],
  ]);
});

// The same script written against the injected `ExcelJS` global rather than
// require(). Both forms must reach the intercepted constructor.
test("the injected ExcelJS global is intercepted too", async () => {
  const target = tempTarget("global");
  const wb = await readBack(
    await compileXlsx(`
const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet('Inventory');
sheet.addRow(['SKU', 'Qty']);
sheet.addRow(['A-1', 7]);
await wb.xlsx.writeFile(${JSON.stringify(target)});
`),
  );
  assert.strictEqual(fs.existsSync(target), false, `${target} was really written`);
  assert.ok(wb.getWorksheet("Inventory"), "the script's own workbook was not returned");
});

// Pattern B — the ready-made `workbook` global, ending in a bare expression.
test("pattern B: the provided workbook global, returned bare", async () => {
  const wb = await readBack(
    await compileXlsx(`
const sheet = workbook.addWorksheet('Data');
sheet.addRow(['a', 1]);
workbook;
`),
  );
  const sheet = wb.getWorksheet("Data");
  assert.ok(sheet);
  assert.deepStrictEqual(cellValues(sheet), [["a", 1]]);
});

// Pattern C — writeBuffer() followed by fs.writeFileSync.
test("pattern C: writeBuffer then fs.writeFileSync", async () => {
  const wb = await readBack(
    await compileXlsx(`
const ExcelJS = require('exceljs');
const fs = require('fs');
const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet('Ledger');
sheet.addRow(['entry', 1]);
const buffer = await wb.xlsx.writeBuffer();
fs.writeFileSync('/home/claude/output.xlsx', buffer);
`),
  );
  assert.ok(wb.getWorksheet("Ledger"), "the buffer written to fake fs was not returned");
});

// A returned workbook must win over the untouched global, and an untouched
// global must never be returned in place of the script's own work.
test("the script's workbook is returned, not the empty global", async () => {
  const wb = await readBack(
    await compileXlsx(`
const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();
wb.addWorksheet('Mine').addRow(['x']);
wb;
`),
  );
  assert.deepStrictEqual(
    wb.worksheets.map((w) => w.name),
    ["Mine"],
  );
});

test("a workbook with no sheets is reported rather than returned", async () => {
  // exceljs will happily write a workbook with no worksheets; Excel will not
  // open it. Better to say so than to hand over a file that fails later.
  await assert.rejects(
    () => compileXlsx("const x = 1;\n", { drainMs: 200 }),
    /no output|no worksheets/i,
  );
});

test("a thrown error reaches the caller unchanged", async () => {
  await assert.rejects(
    () => compileXlsx("throw new Error('spreadsheet exploded');\n"),
    /spreadsheet exploded/,
  );
});

test("number formats and styles survive the round trip", async () => {
  const wb = await readBack(
    await compileXlsx(`
const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet('Money');
sheet.columns = [{ header: 'Item', key: 'item' }, { header: 'Cost', key: 'cost' }];
const row = sheet.addRow({ item: 'Widget', cost: 9.5 });
row.getCell('cost').numFmt = '$#,##0.00';
wb;
`),
  );
  const sheet = wb.getWorksheet("Money");
  assert.strictEqual(sheet.getRow(2).getCell(2).numFmt, "$#,##0.00");
});
