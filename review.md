# Review record — Crepuscule Slog (Node edition)

First reviewed 2026-09-04. Remediated 2026-09-17 on branch `cleanup/2026-09-17`,
not pushed. 79 tests, all passing, `npm test`.

**Touched again on 2026-09-19, after the pass that finished it,** to commit the
`package-lock.json` that pass should have regenerated. No source file changed and
the suite passes at the same 79. Details under "The lockfile" below. It is on the
same branch, so `git diff main` is still the whole of the work on this project.

**Renamed from Kiln to Crepuscule Slog on 2026-10-03**, because the name Kiln
is already in use by other projects on the web. Everything below has been
updated to the new name, including this record, except the remote URL, which
is still `https://github.com/nishantmajhi/Kiln.git`; renaming the GitHub
repository is the author's to do. What changed besides the prose: the page
title and heading, `public/kiln.js` (now `public/crepuscule-slog.js`), the
package name, `createKilnServer` (now `createCrepusculeSlogServer`), and every
`KILN_` name. **The environment variables are a breaking change for anyone who
set them:** `KILN_PORT`, `KILN_HOST` and `KILN_BROWSER` are now
`CREPUSCULE_SLOG_PORT`, `CREPUSCULE_SLOG_HOST` and `CREPUSCULE_SLOG_BROWSER`,
and the old names are no longer read.

This file was a to-do list. It is now a record: what was wrong, what was done,
what was deliberately left, and what the next person should know.

## Corrections to the first review

Two things in the original review need amending, and one needs strengthening.

**"The three test files are scripts, not assertions" understated it.** They do
assert — each checks the ZIP magic bytes. The real problem is worse on both
counts. They reimplement the server's helpers under a comment reading "mirrors
server.js", so they exercise a copy and could not catch a regression in the
original; the copy in `test_xlsx.js` had already diverged into being *more*
correct than the code it mirrored, which is how the interception bug below
survived. And none of the three has ever run: all use `require` alongside
top-level `await`, which CommonJS does not allow, so on Node 26 they fail with
`ERR_AMBIGUOUS_MODULE_SYNTAX` before the first test and on older Node they fail
to parse. The `✓`/`✗` output the README describes was never produced by
anything. They have been deleted and replaced.

**"`fs` is intercepted rather than provided… that is the single insight that
makes the tool practical" is right about the insight and wrong about the
implementation.** The interception missed the most idiomatic exceljs script
entirely. See finding 2.

**"The sandbox is not a security boundary" was correct and, if anything,
generous.** The review's recommendation — change the word, keep the design —
has been followed, but two of the three findings under "Other findings" turned
out to be a live remote-code-execution path rather than a hardening item. See
findings 7 and 8.

## What was found and fixed

Every one of these was reproduced against the original code before it was
touched. Where a number is quoted it came from a probe run against the
unmodified server.

### 1. Rows written as array literals were silently dropped (P1, not in the first review)

The headline bug, and the one that would waste the most of someone's time,
because the output looks fine until you open it.

Scripts ran under `vm.createContext`. A `vm` context has its own intrinsics, so
an array literal created inside it is not an instance of the *host* realm's
`Array`. exceljs chooses how to interpret a row with `value instanceof Array`,
so every `sheet.addRow(['North', 42000])` fell through to the object branch and
added an empty row. A spreadsheet came out with the right sheet names, the
right column widths and number formats, and no data at all.

Scripts are now compiled with `vm.compileFunction`, which builds a function in
this realm and takes the injected globals as parameters. Nothing is lost:
a `vm` context was never providing isolation here, only a way to supply
globals, and it was providing that badly.

### 2. The exceljs interception was bypassed, and wrote to the real disk (P1, not in the first review)

`{ Workbook: PatchedWorkbook, ...ExcelJS }` spreads the library's own
`Workbook` back on top of the intercepted one. Both `require('exceljs')` and
the injected `ExcelJS` global therefore handed scripts an unintercepted
constructor, and `await wb.xlsx.writeFile('/home/claude/output.xlsx')` really
wrote to the real filesystem — on this machine, to `C:\home\claude\`. Crepuscule Slog then
answered the request with the empty global workbook. Verified: the file
appeared on disk, and the download contained no worksheet.

The spread order is now correct, and `tests/xlsx.test.js` asserts on both forms
that no file appears at the path the script names.

### 3. The wrong workbook was returned (P1, not in the first review)

`const wb = result?.xlsx ? result : workbook` falls back to the ready-made
global `workbook`, which always has an `.xlsx`, so the fallback always matched.
A script that built its own workbook and ended with a call rather than a bare
expression got back a valid, empty file. Precedence is now: the workbook the
script returned, then anything captured through the fake `fs`, then the global
workbook *if the script actually added sheets to it*, then an error saying no
worksheets were added.

### 4. Two concurrent .pptx compiles swapped their output (P1, not in the first review)

`compilePptx` patched `PptxGenJS.prototype`, which is process-wide, for the
duration of a compilation. Two overlapping requests installed their capture
callbacks over each other. Reproduced with a 300-slide deck and a 1-slide deck
written 600 ms apart: both requests returned the *same* 1,531,655 bytes — the
one-slide request received the three-hundred-slide deck — and the loser's
`writeFile` fell through to the real filesystem, leaving a `.pptx` in the temp
directory.

Interception is now a subclass, so it belongs to one compilation. A test drives
two interleaved compiles and checks each gets its own slide count, and another
checks the library's prototype is left as it was found.

### 5. Common script shapes failed to compile (P1, not in the first review)

`injectReturn` put `return` in front of the last non-blank line, whatever it
was. A statement split across lines ends in `);` or `];`, and `return );` is a
syntax error, so this failed outright:

```js
fs.writeFileSync(
  '/home/claude/output.docx',
  buf
);
```

So did any script ending in `throw`. A last line containing `$&` came back
corrupted, because the replacement went through `String.prototype.replace`,
whose replacement string treats `$&`, ``$` `` and `$'` as back-references.

It now stops at closing delimiters and statement keywords, splices at an index
instead of by pattern, and — the part that makes it robust rather than
merely less wrong — falls back to the script exactly as pasted if the injected
form does not parse. A guess about the last line can no longer break a working
script.

### 6. The commonest paste took ten seconds (P2, not in the first review)

After a script finished, the code waited for output to arrive through the fake
`fs`, polling for up to ten seconds. For a script ending in `doc;` — the shape
the README itself recommends — the document was already in hand and nothing was
ever coming, so it waited out the full ten seconds every time. Measured: 10,023
ms. The wait is now conditional on there being nothing usable yet. The same
compile takes about 360 ms end to end, most of which is the worker thread.

### 7. The server listened on every interface (P1, first review finding 1)

`listen(PORT)` with no address binds `0.0.0.0`. Confirmed by posting to
`/compile` on this machine's LAN address and receiving a compiled document:
anyone on the network could execute code here. It now binds `127.0.0.1`, with
`CREPUSCULE_SLOG_HOST` for anyone who has a reason to change it.

### 8. Any website could compile on your machine (P1, not in the first review)

This is the one that matters most, because binding to loopback does not fix it
and it is easy to believe it does.

Any page open in the user's browser can POST to `http://localhost:3000`. A POST
with `Content-Type: text/plain` is a CORS "simple request" — no preflight, so
nothing stops it being sent. The page cannot read the reply, but it does not
need to: the code has already run. Verified against the original server with
`Origin: https://evil.example` and a text/plain body — 200, and a document.

`/compile` now requires `application/json`, which puts the request outside the
simple-request rules so the browser must ask permission first and is refused,
and separately rejects any `Origin` that is not the server's own. The origin is
compared against the `Host` header rather than a fixed port, so it survives a
port change and treats `localhost` and `127.0.0.1` as the same place.

### 9. The whole repository was downloadable (P2, not in the first review)

The static root was the project directory. `GET /.git/config`, `/server.js`,
`/package.json` and anything under `node_modules` all returned 200 — over the
network, given finding 7. The page and its assets have moved to `public/`,
which is now the only directory served. The traversal check was already
correct and has been kept, with tests for backslashes, percent-encoded dots and
null bytes.

### 10. An infinite loop ended the server (P2, first review finding 2)

`while (true) {}` in a pasted script hung the single Node process. Not the
request — the process: every later request from anyone went unanswered, which
was confirmed by sending a valid compile afterwards and getting nothing back.

Each compilation now runs in its own worker thread with a 60-second deadline;
overrunning it costs that request a 504 and nothing else. A spare thread is
kept warm because loading the three Office libraries takes about 780 ms. As a
side effect this also isolates compilations from each other completely — no
patched library, no leftover timer and no exhausted heap can reach the next
one.

### 11. No limit on the request body (P3, first review finding 3)

A 40 MB paste was buffered whole. There is now an 8 MB limit, applied as the
body arrives rather than after. The remainder of the upload is read and
discarded rather than the socket being closed, so the client gets the
explanation instead of a connection reset.

### 12. Smaller things

- `require()` resolved inherited keys: `require('constructor')` returned
  `Object.prototype.constructor` instead of throwing. The table now has a null
  prototype.
- The error message said "not available in Crepuscule Slog's sandbox". It now says
  "module allowlist", as the first review asked.
- Timers a script left running were never cancelled; they are now cleared when
  the compilation ends.
- `fs.promises.writeFile` was not stubbed, so a script using it crashed rather
  than being intercepted.
- The page posted to an absolute `http://localhost:3000/compile`. Anyone who
  opened the UI as `127.0.0.1:3000` was making a cross-origin request, and with
  no CORS headers the browser refused to give the reply to the page — the
  download silently failed. It now posts to a relative `/compile`.
- The filename was interpolated into `innerHTML` unescaped. Self-inflicted
  only, but it is one function call to do it properly.
- Added a `LICENSE` (MIT, matching the other projects in this archive).
  `package.json` said ISC and the README pointed at it; both now say MIT.
- `package-lock.json` was never committed. It is now.

## Structure

`server.js` was 430 lines holding the HTTP handling, the vm context, the fake
filesystem, three compilers and the return-injection heuristic, none of which
could be reached without opening a listening socket. It is now 160 lines of
plumbing, with the decisions in `src/`:

```
server.js        sockets, status codes, headers
src/compile.js   the compilers, the fake fs, the injected globals
src/runner.js    a worker thread per compilation, with a deadline
src/worker.js    the thread's entry point
src/policy.js    origin, content type, body size, static paths
public/          index.html, crepuscule-slog.js, style.css, favicon.js
tests/           node --test, no framework, plus a small ZIP reader
```

The first commit on the branch is a pure move with no behaviour change, so the
diff that follows is only the fixes.

## Deliberately not done

**~~XML mode is untested.~~ It is tested now — see "XML mode, in a real
browser" below.** This paragraph said that testing it meant "either extracting
the zip builders from the browser file or driving a headless browser, and
nothing here can do the second". The second turned out to be possible, and the
note was three days stale when it was checked. Left here as written so the
correction has something to correct.

The rest of it was accurate: XML mode lives in `public/crepuscule-slog.js` inside an IIFE,
uses JSZip from a CDN and needs a DOM. It was not in the first review's findings
and it was not touched.

**The UI is unchanged**, per the standing constraint. Layout, styling, copy and
interaction are as they were. Three changes were made inside `public/crepuscule-slog.js`,
all behaviour rather than appearance: the relative compile URL, the escaped
filename, and nothing else.

**The 60-second compile deadline is a guess.** Long enough for a large deck,
short enough to notice. If a legitimate 500-slide presentation ever hits it,
raise it in `src/runner.js`; the worker makes the cost of being wrong small.

**A script can still occupy a thread for its full deadline**, and an
adversarial one could open a thread per request. That is a denial of service
against a server that only listens to you, which is not worth the complexity of
a queue.

**Git history is untouched.** Nothing sensitive is in it — no credentials, no
keys, nothing in `.gitignore` that was committed anyway.

## Notes for a future session

- `npm install` first; the three libraries are not vendored. `npm test` needs
  no server running. As of 2026-09-19 `npm install` leaves the working tree
  clean — before then it did not, and that was this project's fault rather than
  yours. See "The lockfile".
- `tests/zip.js` is a 60-line ZIP reader written for this suite. If you are
  testing anything else that produces an Office file or a `.zip`, take it — it
  reads the central directory properly and inflates with `zlib`.
- The probe scripts used to characterise the original behaviour are gone (they
  lived in the session's scratch directory), but every one of their findings is
  now a test. To see a bug for yourself, revert the fix in `src/compile.js` and
  run the matching test.
- `src/compile.js` can be used as a library — `compile(format, code)` — with no
  server and no worker. That is how the compiler tests drive it. Anything that
  needs a runaway script stopped should go through `src/runner.js` instead.
- The remote is `https://github.com/nishantmajhi/Kiln.git` and this branch has
  not been pushed. The published version still has all of the above in it.

## The lockfile

Fixed 2026-09-19. Noticed because a session running the archive-wide verification
sweep found this repository dirty afterwards and had to work out whether that was
its own doing.

The 2026-09-17 pass edited `package.json` — it set `license` to `MIT`, which had
been `ISC` and did not match the `LICENSE` file, and added `engines: node >= 18`
to say what the code actually needs. It did not re-run `npm install`, so
`package-lock.json` kept the old root entry: `"license": "ISC"` and no `engines`.

npm reconciles the two on the next install, which means the `npm install` this
project's own README and the archive's verify list both tell you to run left
`package-lock.json` modified **every time**, with no way to tell at a glance
whether that was npm tidying up or a real dependency change. The honest options
are to commit the regenerated lockfile or to revert `package.json`; the lockfile
is right, because `package.json` is what the pass meant.

The regenerated lockfile was checked before being committed, which is the part
worth copying if this happens elsewhere: the two files were parsed and compared
entry by entry rather than eyeballed as a diff. 111 dependency entries before,
111 after, zero added, zero removed, and zero changes to any `version`,
`resolved` or `integrity` field. `lockfileVersion` stayed at 3. The only
difference is the root entry's `license` and `engines`. So this commit cannot
have moved a dependency version, which is the one thing a lockfile commit in a
finished project must not do quietly.

## XML mode, in a real browser

Added 2026-09-24. `npm run test:browser` — eight checks, and the manual check
this project had been carrying since 2026-09-17 is discharged.

That check read: "npm start, open <http://localhost:3000>, pick XML from the
dropdown, press Compile on the example, open the file." It was open because
nothing here could drive a browser. `Web/Frontend/With Libraries/MDview`
established on 2026-09-19 that it can — `--headless=new --dump-dom`, no driver,
no protocol client, no dependency — and this is that technique applied here.

`tests/browser/run.js` copies `public/` to a scratch directory with one script
tag appended to `index.html`, and checks the copy against the real file byte for
byte up to that line so a page that drifted fails loudly.
`tests/browser/driver.js` then does exactly what the manual check describes, for
all three formats: click the format tab, choose XML in the dropdown, press
Compile on whatever example the page put in the box. JSZip and FileSaver load
from their real CDNs.

**`window.saveAs` is the seam that makes it work.** `crepuscule-slog.js` hands it the
finished Blob, so replacing it captures the exact bytes a download would have
written — no downloads directory, and no need to know when Chrome finished
writing one.

**The "open the file" half is done properly, in Node.** The driver passes the
bytes back and the runner reads their central directory with `tests/zip.js`,
asserting that `[Content_Types].xml`, `_rels/.rels` and the part the pasted XML
became are all present, that the part contains the example's own text, and that
the Blob's MIME type and the download's extension match the format. That is the
reason it was worth automating rather than clicking: a person opening the file in
Word sees "it worked" and cannot see that a part is missing until some later
document needs it. The spreadsheet's assertion lists every cell of all three
rows, because the npm-mode bug this project's pass found dropped every row of
every spreadsheet it built and still produced a file that opened — "a zip came
out" and "the rows are in it" are separate questions here.

It also asserts XML mode issued **no request to the server**, which is the one
claim about XML mode that clicking through would not reveal either way. The page
is loaded from `file://` with no server running, so a stray request would fail
silently in the console.

**Results.** All eight pass. Nothing was wrong: XML mode works, and it always
did — `--original` stages the baseline commit's page and gets the same eight
passes and byte-for-byte identical archives (2725, 3645 and 10705 bytes). So
this is a check that nothing regressed, confirmed, rather than a fix. Every one
of the driver's seams — the four element ids, the tab markup and the
`window.saveAs` call — is byte-identical between the baseline and the current
page, so nothing is transcribed or adapted, and the baseline commit is looked up
by subject rather than by SHA so a rewritten history fails loudly.

**And the suite was shown to be capable of failing**, which matters more than
usual here because everything passed on both versions. `zip.folder("word").file`
in `compileDocxXml` was changed to write `<w:document/>` instead of the user's
XML, and the run reported exactly that: *`word/document.xml` is missing "Hello
from Crepuscule Slog (XML mode)" from the example*. That is the shape of the bug that
mattered in this project — a compiler that produces a valid file with nothing in
it — and the suite catches it. The mutation was reverted.

It is a separate script rather than part of `npm test`, deliberately, so Crepuscule Slog's
recorded count stays **79** and the archive's verify sweep does not acquire a
dependency on a browser and a network. With neither, it exits 0 and says SKIP;
what a skip loses is exactly what `npm test` already did not cover.

**One trap if you write another of these**, found the same day in
`Web/Frontend/With Libraries/docxDiff` and fixed here before it could bite.
`--virtual-time-budget` has to exceed the driver's **own** timeouts added
together, not the time a passing run takes. This runner waits up to ten seconds
per format, so a run where all three fail needs thirty; it was budgeted twenty,
which would have made Chrome dump the DOM before the driver published and left
the runner reporting "the page produced no results" instead of the three
failures — on exactly the run where the output matters. It is sixty seconds
now. Virtual time costs nothing, so budget generously.

### The flake — found, named and fixed on 2026-09-24

**It was the four-second timeout in `tests/server.test.js`, and the assertion
was "a script that fails comes back as an error the page can show" expecting
500 and getting 504.** It is fixed. What follows replaces three passes of
guessing, and the guess — contention — was right.

The history: this suite failed once on 2026-09-18 and twice on 2026-09-19, both
times on the run after an `npm install` and both times while the machine was
running other projects' suites, and passed on the immediate re-run with nothing
changed. On 2026-09-19 it was attacked with five consecutive `npm install` then
`npm test` cycles, all 79/0, which ruled out "the first run after an install".
Nobody had captured the output, so the failing assertion had never been named,
and the handover asked the next session that saw it to capture it first.

It failed again on 2026-09-24 during the XML-mode browser work, and this time
the output was kept:

```
actual: 504, expected: 500  —  tests/server.test.js:121
"a script that fails comes back as an error the page can show"
```

**The mechanism.** `tests/server.test.js` built one runner for the whole file
with `timeoutMs: 4000`, and `server.js` answers a `CREPUSCULE_SLOG_TIMEOUT` with 504 rather
than the 500 a throwing script produces. So any compile in that file which took
more than four seconds *to be scheduled* reported a timeout instead of its real
result. Only one test in the file actually wanted a short timeout — the
runaway-script one, which has to sit through it. For every other test the four
seconds was not a deadline being measured, just a backstop against hanging the
suite, and it was nine times tighter than it needed to be: a throwing compile
measured six times on an idle machine took 3, 4, 9, 431, 441 and 461 ms, the
spread being worker startup.

The aggravator is Crepuscule Slog's own suite. `node --test` runs the test files in
parallel, and `tests/runner.test.js` spins `while (true) {}` in a worker for
three seconds, which pins a core — so the two files race, and the machine only
had to be busy enough for the losing side to take four seconds to get going.
That is why it never reproduced in isolation and always reproduced "while other
suites were running".

**Reproduced on demand at last**, by running eight copies of the suite
concurrently on this 8-core machine. Against the original, all eight runs
failed, between one and three assertions each, and the named one was among them
at 4075 ms with `actual: 504`. So was "a pasted script comes back as a download",
which is the same thing happening to a compile that should have returned 200.

**The fix** is in `tests/server.test.js` and touches no project code. The
file-wide runner's timeout is now 30000, and the runaway-script test stands up
its own server and runner at 4000 so that the short deadline applies to the one
test that wants it. `request()` and `compile()` took an optional port argument
to make that possible. After it: nine runs — five sequential and four
concurrent — all 79/0, and the two assertions above no longer fail under the
eight-way stress that reproduced them.

### What the eight-way stress still fails, which is new and is not the flake

At eight concurrent copies the fixed suite still loses one or two assertions per
run, and they are always the same two:

- `tests/runner.test.js` — "a script that loops forever is stopped, and the next
  one still works" (its runner is `timeoutMs: 3000`)
- `tests/server.test.js` — "a script that never finishes gets a 504, and the
  server keeps working" (the impatient 4000 runner added by the fix)

Both are the same shape as the flake and neither is the flake: it is the *second
half* of each test failing, the "and the next one still works" compile, which
has to run on the same short-timeout runner as the runaway compile that precedes
it — that is the whole claim, that killing a hung worker does not leave the
runner broken, so it cannot be moved to a different runner without testing
something else. A longer timeout there means the runaway test sits through it.

This was not left alone out of laziness. It is not reachable at any concurrency
this archive actually uses: at four concurrent copies the fixed suite is 79/0
four times over, and the original was 79/0 four times over too. Eight copies of
one Node suite on eight cores is well past what the verify sweep does. Fixing it
properly wants a per-call timeout on `runner.run`, which is production API added
for a test's convenience, and that is a worse trade than a note. **If you see
either of those two fail, it is load, not a regression — re-run it.**

## Licence, 2026-10-02

The author chose one licence for all of their projects: free for personal,
non-commercial use, modifiable for the user's own use, and not to be
distributed, original or modified. No standard licence says that, so `LICENSE`
is a short custom text, the same in every project apart from the year and what
it excludes. The README's licence section says the same, and `package.json`
declares `SEE LICENSE IN LICENSE`, npm's form for a licence with no SPDX
identifier. Anything above that calls the licence an open question, or calls
the project MIT, is superseded.

## CI, 2026-10-02

`.github/workflows/ci.yml` runs `npm test` and the browser suite as two jobs
on `ubuntu-latest`. The browser job needs the network, for the page's CDN
scripts; a runner has it. Nothing has been pushed, so it has not run on GitHub
yet. Every job was run before committing, from a fresh clone of this branch,
the way a runner would run it: Linux jobs in a throwaway Ubuntu 24.04 (the
image `ubuntu-latest` uses) with the same tool versions. 79, and the browser
suite all passing.

The browser job fails if the runner prints `SKIP`, which it does when it finds
no Chrome; otherwise a missing browser would read as a pass.
