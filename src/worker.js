"use strict";

// One compilation, in a thread of its own.
//
// This exists so the host process can give up on a script. A pasted script is
// arbitrary JavaScript running with the full privileges of this process, and
// `while (true) {}` is a single keystroke away; in a single-threaded server
// that means the whole thing stops answering, forever. Here the parent can
// terminate the thread and carry on.
//
// A thread per compilation also means one compilation can never see another's
// state — no library patched underneath it, no leftover timers, no shared
// heap to exhaust.

const { parentPort, workerData } = require("worker_threads");
const { compile } = require("./compile");

if (!parentPort) {
  throw new Error("src/worker.js is meant to be run as a worker thread");
}

// The parent sends the job once the thread is up, so that the cost of loading
// docx, exceljs and pptxgenjs can be paid before the request arrives.
parentPort.on("message", async (job) => {
  try {
    const buffer = await compile(job.format, job.code, {
      drainMs: job.drainMs,
    });
    parentPort.postMessage({ ok: true, bytes: new Uint8Array(buffer) });
  } catch (err) {
    parentPort.postMessage({
      ok: false,
      message: err && err.message ? err.message : String(err),
      stack: err && err.stack ? err.stack : "",
    });
  }
});

// Tell the parent the libraries are loaded and the thread is ready for work.
parentPort.postMessage({ ready: true, pid: workerData && workerData.tag });
