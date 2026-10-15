/*
 * Runs inside the real page, in a real browser.
 *
 * `tests/browser/run.js` copies `public/` into a scratch directory, appends
 * one script tag pointing at this file, and loads the result in headless
 * Chrome. Everything here therefore runs against the real markup, the real
 * `crepuscule-slog.js`, and the real JSZip and FileSaver as fetched from their CDNs.
 *
 * XML mode is the reason this exists. It is the one half of Crepuscule Slog that never
 * touches the server: the zip is assembled by JSZip in the tab and handed to
 * FileSaver, so the 79 Node checks -- which drive the server over a real
 * socket -- make no claim about it at all. The manual check recorded in
 * `review.md` was "npm start, pick XML from the dropdown, press Compile on the
 * example, open the file". This performs that sequence for all three formats
 * and hands the bytes back to the runner, which opens the file properly by
 * reading the zip's central directory.
 *
 * `window.saveAs` is the seam. `crepuscule-slog.js` calls it with the finished Blob, so
 * replacing it captures the exact bytes a download would have written without
 * needing a downloads directory or a way to know when Chrome finished writing
 * one.
 *
 * Results go into a <div id="test-results"> as JSON, which the runner reads
 * out of the dumped DOM.
 */

(function () {
  "use strict";

  const results = [];
  const record = (name, ok, detail) => results.push({ name, ok, ...detail });

  const publish = () => {
    const el = document.createElement("div");
    el.id = "test-results";
    el.textContent = JSON.stringify(results);
    document.body.appendChild(el);
  };

  const $ = (id) => document.getElementById(id);

  /* Chrome's virtual time makes real timers fire instantly, so polling is
   * cheap, but a compile that never calls saveAs must still end the run. */
  function waitFor(predicate, timeoutMs) {
    return new Promise((resolve) => {
      const started = Date.now();
      (function poll() {
        if (predicate()) return resolve(true);
        if (Date.now() - started > timeoutMs) return resolve(false);
        setTimeout(poll, 10);
      })();
    });
  }

  function toBase64(bytes) {
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(
        null,
        bytes.subarray(i, i + chunk),
      );
    }
    return btoa(binary);
  }

  /*
   * npm mode posts to `/compile`. Recording every fetch from here on lets the
   * run assert that the XML path made none -- the page is loaded from
   * file:// with no server running at all, so a stray request would otherwise
   * fail silently in the console and look like nothing.
   */
  window.__crepusculeSlogFetches = [];
  const realFetch = window.fetch;
  window.fetch = function (input) {
    window.__crepusculeSlogFetches.push(String((input && input.url) || input));
    return realFetch.apply(this, arguments);
  };

  let captured = null;
  const realSaveAs = window.saveAs;
  window.saveAs = function (blob, filename) {
    captured = { blob: blob, filename: filename };
  };

  /*
   * One format, by the route a person takes: click the tab, choose XML in the
   * dropdown, press Compile on whatever example the page put in the box.
   *
   * Clicking the tab calls `setFormat(fmt, false)`, which resets the mode to
   * npm and loads the npm example; the `change` on the select is what puts the
   * XML example in the box. Doing it in that order is deliberate -- it is the
   * order the UI forces on a user, and the reverse would silently test the npm
   * example instead.
   */
  async function compileFormat(fmt) {
    const name = fmt + ": XML mode compiles and hands a file to the browser";
    try {
      const tab = Array.prototype.find.call(
        document.querySelectorAll(".tab"),
        (t) => t.dataset.fmt === fmt,
      );
      if (!tab) throw new Error("no tab for " + fmt);
      tab.click();

      const modeSelect = $("mode-select");
      modeSelect.value = "xml";
      modeSelect.dispatchEvent(new Event("change"));
      if (modeSelect.value !== "xml") {
        throw new Error("the dropdown has no xml option for " + fmt);
      }

      const pasted = $("code-input").value;
      if (!/^\s*<\?xml/.test(pasted)) {
        throw new Error(
          "choosing XML did not put XML in the box; got: " +
            JSON.stringify(pasted.slice(0, 60)),
        );
      }

      captured = null;
      $("compile-btn").click();

      const done = await waitFor(
        () => captured !== null || $("result").classList.contains("err"),
        10000,
      );

      const resultText = ($("result-msg") || {}).textContent || "";
      const logText = ($("log") || {}).textContent || "";

      if (!done) {
        throw new Error(
          "compile neither produced a file nor reported an error within the " +
            "budget. log: " + JSON.stringify(logText),
        );
      }
      if (captured === null) {
        throw new Error("compile reported an error: " + resultText);
      }

      const buf = new Uint8Array(await captured.blob.arrayBuffer());
      record(name, true, {
        filename: captured.filename,
        type: captured.blob.type,
        bytes: buf.length,
        zip: toBase64(buf),
        resultText: resultText,
        logText: logText,
      });
    } catch (err) {
      record(name, false, { error: String((err && err.stack) || err) });
    }
  }

  async function main() {
    record("the page's two CDN libraries loaded", typeof JSZip !== "undefined" &&
      typeof realSaveAs === "function", {
      jszip: typeof JSZip,
      fileSaver: typeof realSaveAs,
    });

    for (const fmt of ["docx", "xlsx", "pptx"]) {
      await compileFormat(fmt);
    }

    /*
     * XML mode must not reach the server. Crepuscule Slog's npm mode posts to /api/run;
     * this asserts the browser-only path stayed browser-only, which is the one
     * claim `review.md` makes about XML mode that a person clicking through
     * would not notice either way.
     */
    record(
      "XML mode issued no request to the server",
      window.__crepusculeSlogFetches.length === 0,
      { requests: window.__crepusculeSlogFetches },
    );

    publish();
  }

  main().catch((err) => {
    record("driver", false, { error: String((err && err.stack) || err) });
    publish();
  });
})();
