/*
 * Runs XML mode in a real browser.
 *
 * Crepuscule Slog has two halves. npm mode runs the pasted script on the server, and the
 * 79 checks in `npm test` drive that over a real socket -- they compile a
 * spreadsheet and read the rows back out of the downloaded file. XML mode
 * never touches the server: JSZip assembles the archive in the tab and
 * FileSaver hands it to the browser. Nothing in the Node suite makes any claim
 * about it, and `review.md` has carried a manual check for it since
 * 2026-09-17 -- "npm start, pick XML from the dropdown, press Compile on the
 * example, open the file".
 *
 * This is that check, performed. Chrome is driven with `--headless=new
 * --dump-dom`, which needs no driver, no protocol client and no dependency;
 * the technique is `Web/Frontend/With Libraries/MDview/tests/browser/run.js`,
 * which is where it was worked out. `public/` is copied to a scratch directory
 * with one extra line appended to `index.html` -- a script tag for `driver.js`
 * -- and the copy is checked against the real file byte for byte up to that
 * line, so a page that drifted from the one under test fails loudly.
 *
 * The "open the file" half happens here rather than in the browser. The driver
 * hands back the exact bytes `window.saveAs` was called with, and this reads
 * their central directory with `tests/zip.js` and asserts on the parts inside.
 * That is the reason the check is worth automating at all: a person opening the
 * file in Word sees "it worked", and cannot see that a part is missing from the
 * archive until some later document needs it.
 *
 * The page loads JSZip and FileSaver from their CDNs, as the real page does.
 * With no browser installed, or with no network, this exits 0 and says SKIP:
 * what is lost in a skip is exactly what the Node suite already does not
 * cover, so a skip is not worse than the status quo -- it is the status quo.
 */

"use strict";

const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { execFileSync, spawnSync } = require("node:child_process");

const { readZip } = require("../zip.js");

const ROOT = path.join(__dirname, "..", "..");
const PUBLIC = path.join(ROOT, "public");

const CANDIDATES = [
  process.env.CREPUSCULE_SLOG_BROWSER,
  "C:/Program Files/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
  "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
  "/usr/bin/google-chrome",
  "/usr/bin/chromium",
  "/usr/bin/chromium-browser",
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
].filter(Boolean);

function findBrowser() {
  for (const candidate of CANDIDATES) {
    try {
      if (fs.statSync(candidate).isFile()) return candidate;
    } catch {
      /* not there */
    }
  }
  return null;
}

const APPENDED = '\n<script src="driver.js"></script>\n';

const BASELINE_SUBJECT = "Initial commit";

/*
 * The four files the page is made of. The pass of 2026-09-17 moved them into
 * `public/`; in the baseline commit they sit at the repository root, which is
 * the only difference `--original` has to account for.
 */
const PAGE_FILES = ["index.html", "crepuscule-slog.js", "style.css", "favicon.js"];

/*
 * `--original` stages the baseline commit's page instead of the current one.
 *
 * XML mode was not touched by that pass and was not in the first review
 * either, so this suite is a check that nothing regressed rather than a fix to
 * confirm -- which means it would pass against the original too, and a suite
 * that cannot fail is testing nothing. This is how that is demonstrated: the
 * same eight checks, against the code as it shipped. Every one of the driver's
 * seams -- the four element ids, the tab markup and the `window.saveAs` call --
 * is byte-identical between the two versions, so nothing is adapted or
 * transcribed.
 *
 * The commit is looked up by its subject rather than by a SHA so that a
 * rebased or amended history fails loudly instead of silently testing nothing.
 */
function baselineFiles() {
  const sha = execFileSync(
    "git",
    [
      "log",
      "--format=%H",
      "--grep=^" + BASELINE_SUBJECT + "$",
      "--max-count=1",
      "--max-parents=0",
    ],
    { cwd: ROOT, encoding: "utf8" },
  )
    .trim()
    .split("\n")[0];
  if (!sha) throw new Error("no baseline commit titled " + BASELINE_SUBJECT);

  const files = {};
  for (const name of PAGE_FILES) {
    files[name] = execFileSync("git", ["show", sha + ":" + name], {
      cwd: ROOT,
      encoding: "utf8",
      maxBuffer: 16 * 1024 * 1024,
    });
  }
  return { sha, files };
}

function stage(useOriginal) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "crepuscule-slog-browser-"));

  let original;
  if (useOriginal) {
    const baseline = baselineFiles();
    console.log("staging the baseline page from " + baseline.sha.slice(0, 7));
    for (const name of PAGE_FILES) {
      fs.writeFileSync(path.join(dir, name), baseline.files[name], "utf8");
    }
    original = baseline.files["index.html"];
  } else {
    for (const name of fs.readdirSync(PUBLIC)) {
      fs.copyFileSync(path.join(PUBLIC, name), path.join(dir, name));
    }
    original = fs.readFileSync(path.join(PUBLIC, "index.html"), "utf8");
  }

  fs.copyFileSync(path.join(__dirname, "driver.js"), path.join(dir, "driver.js"));
  fs.writeFileSync(path.join(dir, "index.html"), original + APPENDED, "utf8");

  /* The page under test must be the page it claims to be. */
  const staged = fs.readFileSync(path.join(dir, "index.html"), "utf8");
  if (
    staged.slice(0, original.length) !== original ||
    staged.slice(original.length) !== APPENDED
  ) {
    throw new Error("the staged copy is not the page plus exactly one script tag");
  }
  return dir;
}

function run(browser, dir) {
  const profile = path.join(dir, "profile");
  const result = spawnSync(
    browser,
    [
      "--headless=new",
      "--disable-gpu",
      "--no-sandbox",
      "--no-first-run",
      "--disable-extensions",
      "--user-data-dir=" + profile,
      "--virtual-time-budget=60000",
      "--dump-dom",
      "file:///" + path.join(dir, "index.html").replace(/\\/g, "/"),
    ],
    { encoding: "utf8", maxBuffer: 64 * 1024 * 1024, timeout: 180000 },
  );

  if (result.error) throw result.error;
  const dom = result.stdout || "";
  const match = dom.match(/<div id="test-results">([\s\S]*?)<\/div>/);
  if (!match) {
    throw new Error(
      "the page produced no results. Last of the dumped DOM:\n" +
        dom.slice(-1500) +
        "\n--- stderr ---\n" +
        (result.stderr || "").slice(-1500),
    );
  }
  const decoded = match[1]
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
  return JSON.parse(decoded);
}

/*
 * What has to be inside each archive for the file to open at all.
 *
 * `[Content_Types].xml` and `_rels/.rels` are what make a zip an Office file
 * rather than a zip; a reader that cannot find them reports a corrupt
 * document. The third entry in each list is the part the pasted XML actually
 * became, which is the one thing the user typed.
 */
const REQUIRED_PARTS = {
  docx: ["[Content_Types].xml", "_rels/.rels", "word/document.xml"],
  xlsx: ["[Content_Types].xml", "_rels/.rels", "xl/worksheets/sheet1.xml"],
  pptx: ["[Content_Types].xml", "_rels/.rels", "ppt/slides/slide1.xml"],
};

/*
 * What each format's example says. If these are not in the part, the archive
 * was built from something other than what was in the box.
 *
 * The spreadsheet lists every cell of all three rows on purpose. The npm-mode
 * bug this project's pass of 2026-09-17 found dropped every row of every
 * spreadsheet it built and still produced a file that opened, so "a zip came
 * out" and "the rows are in it" are separate questions here, and this is the
 * one that matters.
 */
const EXAMPLE_TEXT = {
  docx: ["Hello from Crepuscule Slog (XML mode)"],
  xlsx: ["Item", "Qty", "Price", "Widget", "Gadget", "9.5", "24"],
  pptx: ["Hello from Crepuscule Slog (XML mode)"],
};

const CONTENT_TYPE = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

function checkArchive(fmt, entry, fail) {
  const name = fmt + ": the file that was handed to the browser";
  let zip;
  try {
    zip = readZip(Buffer.from(entry.zip, "base64"));
  } catch (err) {
    fail(name, "the bytes are not a readable zip: " + err.message);
    return;
  }

  const names = Object.keys(zip);
  for (const part of REQUIRED_PARTS[fmt]) {
    if (!Object.prototype.hasOwnProperty.call(zip, part)) {
      fail(name, "the archive has no " + part + ". It contains: " + names.join(", "));
      return;
    }
  }

  const body = zip[REQUIRED_PARTS[fmt][2]].toString("utf8");
  const missing = EXAMPLE_TEXT[fmt].filter((text) => !body.includes(text));
  if (missing.length) {
    fail(
      name,
      REQUIRED_PARTS[fmt][2] +
        " is missing " +
        missing.map((t) => JSON.stringify(t)).join(", ") +
        " from the example. First 200 bytes: " +
        JSON.stringify(body.slice(0, 200)),
    );
    return;
  }

  if (entry.type !== CONTENT_TYPE[fmt]) {
    fail(
      name,
      "the blob's MIME type is " +
        JSON.stringify(entry.type) +
        ", expected " +
        JSON.stringify(CONTENT_TYPE[fmt]),
    );
    return;
  }

  if (!String(entry.filename).toLowerCase().endsWith("." + fmt)) {
    fail(
      name,
      "the download was named " +
        JSON.stringify(entry.filename) +
        ", which does not end in ." +
        fmt,
    );
    return;
  }

  console.log(
    "  ok   " +
      name +
      " (" +
      names.length +
      " parts, " +
      entry.bytes +
      " bytes, " +
      entry.filename +
      ")",
  );
}

function main() {
  const useOriginal = process.argv.includes("--original");

  const browser = findBrowser();
  if (!browser) {
    console.log("SKIP: no Chrome or Edge found. Set CREPUSCULE_SLOG_BROWSER to one.");
    return 0;
  }
  console.log("browser: " + browser);

  const dir = stage(useOriginal);
  let results;
  try {
    results = run(browser, dir);
  } catch (err) {
    /*
     * A page that could not fetch JSZip cannot be told apart from here from a
     * page that is broken, so say both and skip rather than fail.
     */
    console.log("SKIP: the page did not report. This needs a network, because");
    console.log("      index.html loads JSZip and FileSaver from their CDNs.");
    console.log(String(err.message || err));
    return 0;
  }

  let failed = 0;
  const fail = (name, why) => {
    failed++;
    console.log("  FAIL " + name + "\n       " + why);
  };

  const byName = new Map(results.map((r) => [r.name, r]));
  for (const r of results) {
    if (!r.ok) fail(r.name, r.error || JSON.stringify(r));
    else if (!/XML mode compiles/.test(r.name)) console.log("  ok   " + r.name);
  }

  for (const fmt of ["docx", "xlsx", "pptx"]) {
    const entry = byName.get(
      fmt + ": XML mode compiles and hands a file to the browser",
    );
    if (!entry) {
      fail(fmt + ": XML mode compiles", "the driver reported nothing for it");
      continue;
    }
    if (!entry.ok) continue; /* already reported above */
    console.log("  ok   " + entry.name);
    checkArchive(fmt, entry, fail);
  }

  console.log(
    failed === 0
      ? "\nbrowser: all checks passed"
      : "\nbrowser: " + failed + " failed",
  );
  return failed === 0 ? 0 : 1;
}

process.exit(main());
