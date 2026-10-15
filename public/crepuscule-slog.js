(function () {
  // ─── XML example snippets ────────────────────────────────────────────────
  const xmlExamples = {
    docx: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Hello from Crepuscule Slog (XML mode)</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>This document was assembled from raw Office XML in your browser.</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`,

    xlsx: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Item</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Qty</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Price</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Widget</t></is></c>
      <c r="B2"><v>4</v></c>
      <c r="C2"><v>9.5</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Gadget</t></is></c>
      <c r="B3"><v>2</v></c>
      <c r="C3"><v>24</v></c>
    </row>
  </sheetData>
</worksheet>`,

    pptx: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Hello from Crepuscule Slog (XML mode)</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`,
  };

  // ─── NPM example snippets ─────────────────────────────────────────────────
  const examples = {
    docx: {
      filename: "output.docx",
      code: `// Paste a Node-style script written against the "docx" package.
// Crepuscule Slog calls Packer.toBuffer(doc) for you - just build "doc" and
// leave it as the LAST expression, or call Packer.toBuffer(doc)
// yourself (the .toBlob()/save part is handled automatically).

const { Document, Paragraph, TextRun, HeadingLevel } = docx;

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun("Hello from Crepuscule Slog")]
      }),
      new Paragraph({
        children: [new TextRun("This document was compiled entirely in your browser from pasted code.")]
      })
    ]
  }]
});

doc;`,
    },
    xlsx: {
      filename: "output.xlsx",
      code: `// Paste a Node-style script written against "exceljs".
// Crepuscule Slog provides a ready "workbook = new ExcelJS.Workbook()" instance
// as a global called workbook - or create your own; either is fine.
// Leave the workbook as the LAST expression.

const wb = new ExcelJS.Workbook();
const sheet = wb.addWorksheet("Sheet1");

sheet.columns = [
  { header: "Item", key: "item", width: 20 },
  { header: "Qty", key: "qty", width: 10 },
  { header: "Price", key: "price", width: 12 }
];

sheet.addRow({ item: "Widget", qty: 4, price: 9.5 });
sheet.addRow({ item: "Gadget", qty: 2, price: 24 });

sheet.getRow(1).font = { bold: true };

wb;`,
    },
    pptx: {
      filename: "output.pptx",
      code: `// Paste a Node-style script written against "pptxgenjs".
// Crepuscule Slog intercepts pptx.writeFile() / pptx.write() and triggers
// a real download instead. Just build "pptx" normally.

const pptx = new PptxGenJS();
pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pptx.layout = "WIDE";

const slide = pptx.addSlide();
slide.addText("Hello from Crepuscule Slog", {
  x: 0.5, y: 0.4, w: 12, h: 1, fontSize: 36, bold: true, color: "1D1B16"
});
slide.addText("Compiled in-browser from pasted code.", {
  x: 0.5, y: 1.6, w: 12, h: 0.6, fontSize: 18, color: "5B574C"
});

pptx;`,
    },
  };

  const pillLabels = {
    docx: { npm: "expects: docx (npm)", xml: "expects: docx (XML)" },
    xlsx: { npm: "expects: exceljs (npm)", xml: "expects: xlsx (XML)" },
    pptx: { npm: "expects: pptxgenjs (npm)", xml: "expects: pptx (XML)" },
  };

  const accent = {
    docx: "var(--word)",
    xlsx: "var(--excel)",
    pptx: "var(--ppt)",
  };
  let current = "docx";
  let mode = "npm"; // "npm" | "xml"

  const tabs = document.querySelectorAll(".tab");
  const input = document.getElementById("code-input");
  const modeSelect = document.getElementById("mode-select");
  const compileBtn = document.getElementById("compile-btn");
  const filenameInput = document.getElementById("filename-input");
  const logEl = document.getElementById("log");
  const resultEl = document.getElementById("result");
  const resultMsg = document.getElementById("result-msg");
  const retryBtn = document.getElementById("retry-btn");

  let lastError = "";

  function buildSelectOptions(fmt) {
    modeSelect.innerHTML = "";
    const opts = [
      { value: "npm", label: pillLabels[fmt].npm },
      { value: "xml", label: pillLabels[fmt].xml },
    ];
    opts.forEach((o) => {
      const el = document.createElement("option");
      el.value = o.value;
      el.textContent = o.label;
      if (o.value === mode) el.selected = true;
      modeSelect.appendChild(el);
    });
  }

  function applyMode() {
    const isXml = mode === "xml";
    modeSelect.classList.toggle("xml-mode", isXml);
    modeSelect.style.setProperty("--accent", accent[current]);

    // Placeholder text direction
    input.placeholder = isXml
      ? "Paste the raw XML Claude generated (document.xml, sheet1.xml, or slide1.xml)…"
      : "Paste the Node-style script Claude generated…";
  }

  function setFormat(fmt, keepMode) {
    current = fmt;
    if (!keepMode) mode = "npm";

    tabs.forEach((t) => {
      const active = t.dataset.fmt === fmt;
      t.classList.toggle("active", active);
      t.setAttribute("aria-selected", active ? "true" : "false");
    });

    buildSelectOptions(fmt);
    applyMode();

    // Load appropriate example
    input.value = mode === "xml" ? xmlExamples[fmt] : examples[fmt].code;
    filenameInput.value = examples[fmt].filename;

    compileBtn.style.setProperty("--accent", accent[fmt]);
    input.style.setProperty("--accent", accent[fmt]);
    resultEl.classList.remove("show", "ok", "err");
    logEl.classList.remove("show");
    logEl.innerHTML = "";
  }

  modeSelect.addEventListener("change", () => {
    mode = modeSelect.value;
    applyMode();
    input.value =
      mode === "xml" ? xmlExamples[current] : examples[current].code;
    resultEl.classList.remove("show", "ok", "err");
    logEl.classList.remove("show");
    logEl.innerHTML = "";
  });

  tabs.forEach((t) =>
    t.addEventListener("click", () => setFormat(t.dataset.fmt, false)),
  );

  document.getElementById("example-btn").addEventListener("click", () => {
    input.value =
      mode === "xml" ? xmlExamples[current] : examples[current].code;
  });
  document.getElementById("clear-btn").addEventListener("click", () => {
    input.value = "";
    input.focus();
  });

  // ─── Filename auto-extraction ─────────────────────────────────────────────
  // npm mode: read "// filename.docx" from the very first line of the script.
  // xml mode (or npm mode with no matching first-line comment): extract the
  // text of the first Title-styled paragraph from the XML.
  //
  // In either case the result is written into filenameInput so the user can
  // still edit it before compiling.

  function extractFilenameFromCode(code) {
    if (mode === "npm") {
      // Check first line for a comment like:  // my-assessment.docx
      var firstLine = code.split(/\r?\n/)[0].trim();
      var commentMatch = firstLine.match(/^\/\/\s*(.+\.[a-zA-Z0-9]+)\s*$/);
      if (commentMatch) {
        return commentMatch[1].trim();
      }
    }

    // XML mode OR npm mode with no valid first-line comment:
    // look for the Title style in the XML pasted (document.xml format).
    // Pattern: <w:pStyle w:val="Title"/> inside a paragraph that has a run.
    var titleMatch = code.match(
      /<w:pStyle[^>]+w:val=["']Title["'][^>]*\/>[\s\S]*?<w:t[^>]*>([\s\S]*?)<\/w:t>/
    );
    if (titleMatch) {
      var titleText = titleMatch[1].trim();
      if (titleText) {
        // Turn the title into a safe filename
        var safe = titleText
          .replace(/[\\/:*?"<>|]/g, "")
          .replace(/\s+/g, "_")
          .substring(0, 80);
        var ext = current || "docx";
        return safe + "." + ext;
      }
    }

    return null;
  }

  function maybeUpdateFilename() {
    var code = input.value;
    if (!code.trim()) return;
    var detected = extractFilenameFromCode(code);
    if (detected) {
      filenameInput.value = detected;
    }
  }

  input.addEventListener("input", maybeUpdateFilename);
  input.addEventListener("paste", function () {
    // paste fires before value updates, so defer by one tick
    setTimeout(maybeUpdateFilename, 0);
  });

  // Also re-run extraction when mode changes, in case code is already present
  modeSelect.addEventListener("change", function () {
    setTimeout(maybeUpdateFilename, 0);
  });

  // ─── Logging helpers ──────────────────────────────────────────────────────
  function logLine(text, cls) {
    const div = document.createElement("div");
    div.className = "log-line " + (cls || "");
    const time = document.createElement("span");
    time.className = "t";
    time.textContent = new Date().toLocaleTimeString([], {
      hour12: false,
    });
    div.appendChild(time);
    const msg = document.createElement("span");
    msg.textContent = text;
    div.appendChild(msg);
    logEl.appendChild(div);
    logEl.classList.add("show");
    logEl.scrollTop = logEl.scrollHeight;
  }

  function showResult(ok, message, detail) {
    resultEl.classList.add("show");
    resultEl.classList.toggle("ok", ok);
    resultEl.classList.toggle("err", !ok);
    resultMsg.innerHTML = message;
    lastError = detail || "";
    retryBtn.style.display = detail ? "inline-flex" : "none";
  }

  retryBtn.addEventListener("click", () => {
    logLine(lastError, "err");
  });

  // ─── NPM compilers - POST to local Node server ───────────────────────────
  // The server runs the script in a real Node.js environment with docx,
  // exceljs, and pptxgenjs pre-installed. It intercepts fs.writeFileSync /
  // writeFile so any hardcoded path (e.g. /home/claude/output.docx) is
  // ignored - only the bytes matter.

  // Relative, so the request always goes back to the origin that served the
  // page. An absolute http://localhost:3000 made the POST cross-origin for
  // anyone who opened the UI as 127.0.0.1:3000, and the browser then refused
  // to hand the reply to the page.
  const SERVER = "";

  async function postToServer(format, code) {
    logLine(`Sending to local Crepuscule Slog server (${format})…`, "work");
    let res;
    try {
      res = await fetch(`${SERVER}/compile`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ format, code }),
      });
    } catch (networkErr) {
      throw new Error(
        "Could not reach the Crepuscule Slog server at localhost:3000. " +
          "Make sure you ran `node server.js` in the project folder.",
      );
    }

    if (!res.ok) {
      let msg = `Server error ${res.status}`;
      try {
        const json = await res.json();
        msg = json.error || msg;
      } catch (_) {}
      throw new Error(msg);
    }

    return res.blob();
  }

  async function compileDocxNpm(code) {
    logLine("Running script in Node.js (docx)…", "work");
    return postToServer("docx", code);
  }

  async function compileXlsxNpm(code) {
    logLine("Running script in Node.js (exceljs)…", "work");
    return postToServer("xlsx", code);
  }

  async function compilePptxNpm(code) {
    logLine("Running script in Node.js (pptxgenjs)…", "work");
    return postToServer("pptx", code);
  }

  // ─── XML compilers ────────────────────────────────────────────────────────
  // Each builds a minimal but valid Office Open XML zip from the raw XML
  // that Claude left in the chat. These mirror the minimal required parts
  // that Office apps need to open the file without complaints.

  const CT_DOCX =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const CT_XLSX =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const CT_PPTX =
    "application/vnd.openxmlformats-officedocument.presentationml.presentation";

  async function compileDocxXml(xml) {
    logLine("Building .docx zip from raw XML…", "work");
    if (typeof JSZip === "undefined")
      throw new Error(
        "JSZip failed to load. Check your connection and reload.",
      );

    const zip = new JSZip();

    // [Content_Types].xml
    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`,
    );

    // _rels/.rels
    zip.folder("_rels").file(
      ".rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
    );

    // word/document.xml - user's XML
    zip.folder("word").file("document.xml", xml);

    // word/_rels/document.xml.rels
    zip
      .folder("word")
      .folder("_rels")
      .file(
        "document.xml.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`,
      );

    // word/settings.xml
    zip.folder("word").file(
      "settings.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>`,
    );

    logLine("Zipping document…", "work");
    const buf = await zip.generateAsync({
      type: "arraybuffer",
      mimeType: CT_DOCX,
    });
    return new Blob([buf], { type: CT_DOCX });
  }

  async function compileXlsxXml(xml) {
    logLine("Building .xlsx zip from raw XML…", "work");
    if (typeof JSZip === "undefined")
      throw new Error(
        "JSZip failed to load. Check your connection and reload.",
      );

    const zip = new JSZip();

    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`,
    );

    zip.folder("_rels").file(
      ".rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    );

    const xlDir = zip.folder("xl");

    xlDir.file(
      "workbook.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
    );

    xlDir.folder("_rels").file(
      "workbook.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`,
    );

    // User's XML goes in as sheet1
    xlDir.folder("worksheets").file("sheet1.xml", xml);

    // Minimal sharedStrings so Excel doesn't complain
    xlDir.file(
      "sharedStrings.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>`,
    );

    logLine("Zipping workbook…", "work");
    const buf = await zip.generateAsync({
      type: "arraybuffer",
      mimeType: CT_XLSX,
    });
    return new Blob([buf], { type: CT_XLSX });
  }

  async function compilePptxXml(xml) {
    logLine("Building .pptx zip from raw XML…", "work");
    if (typeof JSZip === "undefined")
      throw new Error(
        "JSZip failed to load. Check your connection and reload.",
      );

    const zip = new JSZip();

    zip.file(
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
    );

    zip.folder("_rels").file(
      ".rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`,
    );

    const pptDir = zip.folder("ppt");

    pptDir.file(
      "presentation.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
</p:presentation>`,
    );

    pptDir.folder("_rels").file(
      "presentation.xml.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
    );

    // User's XML as slide1
    pptDir.folder("slides").file("slide1.xml", xml);

    pptDir
      .folder("slides")
      .folder("_rels")
      .file(
        "slide1.xml.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,
      );

    // Minimal slide master
    const masterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`;
    pptDir.folder("slideMasters").file("slideMaster1.xml", masterXml);

    pptDir
      .folder("slideMasters")
      .folder("_rels")
      .file(
        "slideMaster1.xml.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`,
      );

    // Minimal slide layout
    const layoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="blank">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;
    pptDir.folder("slideLayouts").file("slideLayout1.xml", layoutXml);

    // slideLayout1.xml.rels - the back-reference to the slide master must be
    // present or PowerPoint rejects the file with "can't read". The correct
    // rel type for this direction is slideLayoutIdLst → slideMaster.
    pptDir
      .folder("slideLayouts")
      .folder("_rels")
      .file(
        "slideLayout1.xml.rels",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`,
      );

    // Minimal theme (required by the layout relationship above)
    pptDir.folder("theme").file(
      "theme1.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1>
      <a:lt1><a:sysClr lastClr="ffffff" val="window"/></a:lt1>
      <a:dk2><a:srgbClr val="1F3864"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5A96D1"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`,
    );

    logLine("Zipping presentation…", "work");
    const buf = await zip.generateAsync({
      type: "arraybuffer",
      mimeType: CT_PPTX,
    });
    return new Blob([buf], { type: CT_PPTX });
  }

  // ─── Main compile dispatcher ──────────────────────────────────────────────
  compileBtn.addEventListener("click", async () => {
    const code = input.value.trim();
    resultEl.classList.remove("show", "ok", "err");
    logEl.innerHTML = "";
    logEl.classList.remove("show");

    if (!code) {
      showResult(
        false,
        "<b>Nothing to compile.</b> Paste Claude's generated script first.",
      );
      return;
    }

    compileBtn.disabled = true;
    compileBtn.textContent = "Compiling…";
    logLine(`Target: ${current.toUpperCase()} · mode: ${mode}`, "work");

    try {
      let blob;
      if (mode === "xml") {
        if (current === "docx") blob = await compileDocxXml(code);
        else if (current === "xlsx") blob = await compileXlsxXml(code);
        else blob = await compilePptxXml(code);
      } else {
        if (current === "docx") blob = await compileDocxNpm(code);
        else if (current === "xlsx") blob = await compileXlsxNpm(code);
        else blob = await compilePptxNpm(code);
      }

      let filename = filenameInput.value.trim() || examples[current].filename;
      if (!filename.toLowerCase().endsWith("." + current)) {
        filename = filename.replace(/\.[a-zA-Z0-9]+$/, "") + "." + current;
        filenameInput.value = filename;
      }

      logLine(`Compiled ${(blob.size / 1024).toFixed(1)} KB`, "ok");
      window.saveAs(blob, filename);
      showResult(
        true,
        // The filename comes from an input box and from the pasted script's
        // first line, and this string is assigned to innerHTML.
        `<b>Done.</b> "${escapeHtml(filename)}" was compiled and downloaded - check your downloads folder.`,
      );
    } catch (err) {
      console.error(err);
      logLine("Compile failed", "err");
      showResult(
        false,
        `<b>Compile failed.</b> ${escapeHtml(err.message || String(err))}`,
        err.stack || String(err),
      );
    } finally {
      compileBtn.disabled = false;
      compileBtn.textContent = "▶ Compile";
    }
  });

  function escapeHtml(s) {
    return String(s).replace(
      /[&<>"']/g,
      (c) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&#39;",
        })[c],
    );
  }

  // init
  setFormat("docx");
})();
