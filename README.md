# Crepuscule Slog

Paste the script an LLM wrote to build a `.docx`, `.xlsx` or `.pptx`. Get the
file. It runs on your machine, takes about a second, and costs nothing.

## The problem it solves

Ask a model for a Word document and it does not hand you a document. It hands
you a short Node script using `docx`, `exceljs` or `pptxgenjs` — the same
libraries its own document skills are written around. The script is usually
correct on the first try, because building an Office file from a description is
the part models are good at.

Turning that script into a file is the part that goes wrong. It needs a machine
with those three packages installed, and the script writes to a path like
`/home/claude/output.docx` that exists on no machine you own. So the usual next
step is a cloud sandbox: spin one up, watch the script fail on something
environmental, ask the model to fix it, run it again. Every one of those turns
is another inference call spent on a step with no judgement in it. The code is
already right; it just needs running.

Crepuscule Slog runs it. The libraries are installed here, the output path is thrown away
and the bytes come back as a download.

## Getting started

Node 18 or newer.

```bash
npm install
npm start
```

Then open <http://localhost:3000>. Leave the terminal running while you use it.

Paste a script, pick the format it builds, press Compile. The file downloads.

## The two modes

**npm mode** is the one described above: a Node script, run by the server in
this project. Whatever the script does with its output — returns the document,
awaits a buffer, calls `fs.writeFileSync`, hands it to a `.then()` — Crepuscule Slog picks
it up. The three shapes a model commonly produces all work:

```js
// the document as a trailing expression
const doc = new Document({ sections: [...] });
doc;

// a buffer written to a path that does not exist here
Packer.toBuffer(doc).then(buf => fs.writeFileSync('/home/claude/out.docx', buf));

// awaited and written
const buf = await Packer.toBuffer(doc);
fs.writeFileSync('/home/claude/out.docx', buf);
```

`require('docx')`, `require('exceljs')` and `require('pptxgenjs')` resolve, and
so do the same names used as bare globals, because different models write it
differently. `require('fs')` gets the fake filesystem. Anything else says so
plainly rather than failing somewhere confusing.

**XML mode** is for when the model wrote raw Office Open XML — a `document.xml`,
a `sheet1.xml`, a `slide1.xml` — instead of a script. The browser wraps the
fragment in the minimum set of parts an Office application needs and zips it
with JSZip. The server is not involved at all.

## What to know before you paste something

**Pasted code runs with your user's privileges.** It is not sandboxed and
cannot be. It can read your files, reach the network and start processes,
exactly as `node script.js` would. The `vm` module Crepuscule Slog used to lean on is not
a security boundary — Node's own documentation says so — and no configuration
of it makes one.

That is a reasonable position for what Crepuscule Slog is for: running a script your own
model wrote a minute ago, in a conversation you were part of, on your own
machine. It is not a reasonable position for code from a forum, a shared chat
log or a stranger. The rule is the one you would apply to a `.js` file
someone sent you: if you would not run it with `node`, do not paste it here.

What Crepuscule Slog does provide is containment against accident and interference:

- The server listens on `127.0.0.1` only. Nothing else on your network can
  reach it.
- `/compile` refuses requests from any other origin, and requires a JSON
  content type. Without both, a page open in another browser tab could post
  code here and have it run — it could not read the reply, but it would not
  need to.
- Each compilation runs in its own worker thread, so a script that loops
  forever is killed after a minute instead of stopping the server, and no
  script can see or disturb another.
- Writes to the filesystem are intercepted. A script writing to
  `/home/claude/output.docx`, or to `C:\Windows\System32`, produces bytes and
  no file.

## Layout

```
crepuscule-slog/
├── server.js          HTTP: static files, POST /compile, and nothing else
├── src/
│   ├── compile.js     the three compilers, the fake fs, the injected globals
│   ├── runner.js      a worker thread per compilation, with a deadline
│   ├── worker.js      the thread's entry point
│   └── policy.js      what a request is allowed to do
├── public/            the page: index.html, crepuscule-slog.js, style.css, favicon.js
└── tests/             node --test, no framework
```

The interesting file is `src/compile.js`. Two decisions in it are worth
knowing about before changing anything:

**The fake `fs` discards the path and keeps the bytes.** That one idea is what
makes a script written for someone else's machine work on yours.

**Pasted code is compiled into this realm, not into a `vm` context.** A `vm`
context has its own intrinsics, so an array literal written inside one is not
an instance of the outer realm's `Array`. exceljs decides how to read a row
with `value instanceof Array`, so under a `vm` context every
`sheet.addRow(['North', 42000])` silently added an empty row and the
spreadsheet arrived with its sheet names, its formatting and none of its data.
Running in one realm is what stops that. It costs nothing in isolation,
because a `vm` context never provided any.

## Tests

```bash
npm test
```

79 checks, no test framework beyond `node --test`, and no server needs to be
running. They cover:

| File | What it holds to account |
| --- | --- |
| `tests/docx.test.js` | the docx compiler, each script-ending shape, the fake fs, the module allowlist |
| `tests/xlsx.test.js` | the exceljs compiler, including that rows actually arrive |
| `tests/pptx.test.js` | the pptxgenjs compiler, and that two compilations cannot swap outputs |
| `tests/injectreturn.test.js` | the guess about the last line, and that it never breaks a working script |
| `tests/policy.test.js` | origin, content type, body and static-path rules |
| `tests/runner.test.js` | the worker: deadlines, crashes, isolation between runs |
| `tests/server.test.js` | the whole thing over a socket, including what it binds to |

The assertions read the produced archive rather than checking its length:
`tests/zip.js` is a small ZIP reader written for the purpose, so a test can ask
whether the words the script wrote are in `word/document.xml`. A file that
starts with `PK` is not evidence of much — an empty document starts with `PK`
too.

## Libraries

| Package | Version | For |
| --- | --- | --- |
| docx | 8.5.0 | `.docx` |
| exceljs | 4.4.0 | `.xlsx` |
| pptxgenjs | 3.12.0 | `.pptx` |

XML mode additionally loads JSZip and FileSaver from a CDN in the browser.

## Configuration

| Variable | Default | Effect |
| --- | --- | --- |
| `CREPUSCULE_SLOG_PORT` | `3000` | Port to listen on |
| `CREPUSCULE_SLOG_HOST` | `127.0.0.1` | Address to bind. Change it only if you have a reason, and understand that the reason is to let other machines run code on this one. |

## Licence

Free for personal, non-commercial use. You may use, study and modify this
software for yourself, but you may not use it commercially, and you may not
share or publish it, in its original form or modified. See [LICENSE](LICENSE)
for the full terms.
