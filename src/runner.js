"use strict";

// The host side of the worker arrangement: hand a job to a thread, wait no
// longer than the deadline, and kill the thread if it overruns.
//
// Every compilation gets its own thread, which is the only way to be sure a
// script cannot affect the next one. Loading the three Office libraries costs
// most of a second, so one spare thread is kept warm and unreferenced; a
// compilation takes the warm one and starts the next spare behind it.

const path = require("path");
const { Worker } = require("worker_threads");

const WORKER_FILE = path.join(__dirname, "worker.js");

// Long enough for a large deck, short enough that a stuck script is noticed.
const DEFAULT_TIMEOUT_MS = 60000;

function createRunner(options = {}) {
  const timeoutMs = options.timeoutMs || DEFAULT_TIMEOUT_MS;
  const drainMs = options.drainMs;
  let spare = null;
  let stopped = false;

  function spawn() {
    const worker = new Worker(WORKER_FILE, { workerData: {} });
    const state = { worker, ready: false, waiters: [] };
    worker.on("message", (msg) => {
      if (msg && msg.ready) {
        state.ready = true;
        for (const resolve of state.waiters.splice(0)) resolve();
      }
    });
    // The spare must not hold the process open on its own account.
    worker.unref();
    return state;
  }

  function takeWorker() {
    const state = spare && !stopped ? spare : spawn();
    spare = stopped ? null : spawn();
    state.worker.ref();
    return state;
  }

  // Resolves with the compiled bytes, or rejects with the script's own error,
  // a timeout, or a thread that died under it.
  function run(format, code) {
    const state = takeWorker();
    const { worker } = state;

    return new Promise((resolve, reject) => {
      let settled = false;
      const finish = (fn, value) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        worker.terminate();
        fn(value);
      };

      const timer = setTimeout(() => {
        finish(
          reject,
          Object.assign(
            new Error(
              `Compilation did not finish within ${Math.round(timeoutMs / 1000)} seconds and was stopped. ` +
                "A script that loops forever, or waits on something that never arrives, looks like this.",
            ),
            { code: "CREPUSCULE_SLOG_TIMEOUT" },
          ),
        );
      }, timeoutMs);

      worker.on("message", (msg) => {
        if (!msg || msg.ready) return;
        if (msg.ok) finish(resolve, Buffer.from(msg.bytes));
        else {
          const err = new Error(msg.message);
          err.scriptStack = msg.stack;
          finish(reject, err);
        }
      });

      worker.on("error", (err) => finish(reject, err));

      worker.on("exit", (exitCode) => {
        if (settled) return;
        finish(
          reject,
          new Error(
            `The compiler thread stopped unexpectedly (exit code ${exitCode}). ` +
              "A script that exhausts memory ends this way.",
          ),
        );
      });

      const send = () => worker.postMessage({ format, code, drainMs });
      if (state.ready) send();
      else state.waiters.push(send);
    });
  }

  // Let go of the warm spare, so a process using the runner can exit.
  async function shutdown() {
    stopped = true;
    const held = spare;
    spare = null;
    if (held) await held.worker.terminate();
  }

  // Start the first spare now, so the first compilation is as quick as the
  // rest.
  spare = spawn();

  return { run, shutdown };
}

module.exports = { createRunner, DEFAULT_TIMEOUT_MS };
