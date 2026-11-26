// Chord's Office add-in: the task pane. It talks only to Chord on this computer (127.0.0.1):
// pairing once with the code Chord shows, then asking questions with a token Chord gave it.
// Answers are shown as plain text (never as HTML).
/* global Office, ChordOffice */
"use strict";

(function () {
  /** The ports Chord listens on, in order (src/main/office/officeAddress.ts: OFFICE_PORTS). */
  const PORTS = [47820, 47821, 47822];
  const ASK_TIMEOUT_MS = 16 * 60_000;
  const STATUS_TIMEOUT_MS = 4_000;
  const HOST_WORDS = {
    word: { selection: "The selected text", document: "The whole document" },
    excel: { selection: "The selected cells", document: "The whole sheet" },
    powerpoint: { selection: "The selected slide's text", document: "The whole presentation's text" },
  };
  const AI_NAMES = { claude: "Claude", codex: "ChatGPT", gemini: "Gemini" };

  let host;
  let answerText = "";

  const $ = (id) => document.getElementById(id);

  // ---------- the token (kept by this add-in only; Chord keeps only its hash) ----------

  function storageKey() {
    const partition = typeof Office !== "undefined" && Office.context && Office.context.partitionKey;
    return (partition ? partition + ":" : "") + "chord.office.v1";
  }
  function loadPairing() {
    try {
      const saved = JSON.parse(localStorage.getItem(storageKey()) || "null");
      if (saved && typeof saved.token === "string" && PORTS.includes(saved.port)) return saved;
    } catch (err) {
      /* nothing saved */
    }
    return undefined;
  }
  function savePairing(pairing) {
    try {
      if (pairing) localStorage.setItem(storageKey(), JSON.stringify(pairing));
      else localStorage.removeItem(storageKey());
    } catch (err) {
      /* storage off: the add-in pairs again next time */
    }
  }

  // ---------- talking to Chord ----------

  async function call(port, path, options, timeoutMs) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const res = await fetch("http://127.0.0.1:" + port + path, { ...options, signal: controller.signal, cache: "no-store", credentials: "omit", referrerPolicy: "no-referrer" });
      let body = {};
      try {
        body = await res.json();
      } catch (err) {
        body = {};
      }
      return { status: res.status, body };
    } finally {
      clearTimeout(timer);
    }
  }

  function show(id) {
    for (const screen of ["connecting", "offline", "pairing", "asking"]) $(screen).hidden = screen !== id;
  }
  function say(id, text, problem) {
    const el = $(id);
    el.textContent = text || "";
    el.classList.toggle("problem", Boolean(problem));
  }

  /** Paired and Chord answers: ask. Paired but the token is gone: pair. Chord not there: say so. */
  async function start() {
    show("connecting");
    const pairing = loadPairing();
    if (!pairing) return show("pairing");
    try {
      const { status } = await call(pairing.port, "/status", { method: "GET", headers: { Authorization: "Bearer " + pairing.token } }, STATUS_TIMEOUT_MS);
      if (status === 200) return show("asking");
      if (status === 401) {
        savePairing(undefined);
        return show("pairing");
      }
      show("offline");
    } catch (err) {
      show("offline");
    }
  }

  async function pair(event) {
    event.preventDefault();
    const code = $("code").value.replace(/[\s-]/g, "").toUpperCase();
    if (!/^[A-Z0-9]{8}$/.test(code)) return say("pair-status", "Type the 8 letters and numbers Chord shows.", true);
    $("pair-button").disabled = true;
    say("pair-status", "Pairing…");
    let reached = false;
    for (const port of PORTS) {
      try {
        const { status, body } = await call(port, "/pair", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ code }) }, STATUS_TIMEOUT_MS);
        reached = true;
        if (status === 200 && typeof body.token === "string") {
          savePairing({ token: body.token, port });
          $("code").value = "";
          say("pair-status", "");
          $("pair-button").disabled = false;
          return show("asking");
        }
        say("pair-status", body.message || "That didn't work. Try again.", true);
        break;
      } catch (err) {
        /* not on this port: try the next */
      }
    }
    if (!reached) say("pair-status", "Chord isn't answering. Open Chord on this computer and try again.", true);
    $("pair-button").disabled = false;
  }

  async function askChord(event) {
    event.preventDefault();
    const pairing = loadPairing();
    if (!pairing) return show("pairing");
    const question = $("question").value.trim();
    if (!question) return say("ask-status", "Type a question first.", true);
    const kind = document.querySelector('input[name="context"]:checked').value;
    $("ask-button").disabled = true;
    $("answer-box").hidden = true;
    say("put-status", "");
    try {
      const context = kind === "none" || !host ? undefined : { kind, text: await ChordOffice.readText(host, kind) };
      if (context && !context.text.trim()) say("ask-status", kind === "selection" ? "Nothing is selected, so Chord gets the question only. Asking…" : "The document is empty. Asking…");
      else say("ask-status", "Asking Chord… The answer also appears in Chord's Office chat.");
      const body = { question, target: $("target").value, host: host || "word", ...(context && context.text.trim() ? { context } : {}) };
      const { status, body: reply } = await call(pairing.port, "/ask", { method: "POST", headers: { "Content-Type": "application/json", Authorization: "Bearer " + pairing.token }, body: JSON.stringify(body) }, ASK_TIMEOUT_MS);
      if (status === 401) {
        savePairing(undefined);
        say("ask-status", "");
        return show("pairing");
      }
      if (status !== 200 || typeof reply.answer !== "string") return say("ask-status", reply.message || "Chord couldn't answer. Try again.", true);
      answerText = reply.answer;
      $("answer").textContent = answerText;
      $("answer-by").textContent = AI_NAMES[reply.by] ? "Answered by " + AI_NAMES[reply.by] : "";
      $("answer-box").hidden = false;
      say("ask-status", "");
    } catch (err) {
      say("ask-status", err && err.name === "AbortError" ? "That took too long. See the Office chat in Chord." : "Chord isn't answering. Open Chord on this computer and try again.", true);
    } finally {
      $("ask-button").disabled = false;
    }
  }

  async function put(mode) {
    if (!answerText) return;
    say("put-status", "");
    try {
      await ChordOffice.putAnswer(host, answerText, mode);
      say("put-status", mode === "replace" ? "Replaced the selection." : "Inserted.");
    } catch (err) {
      say("put-status", (err && err.message) || "That didn't work.", true);
    }
  }

  function setUp(info) {
    host = ChordOffice.hostOf(info);
    const words = HOST_WORDS[host || "word"];
    $("selection-label").textContent = words.selection;
    $("document-label").textContent = words.document;
    $("pair-form").addEventListener("submit", pair);
    $("ask-form").addEventListener("submit", askChord);
    $("retry").addEventListener("click", start);
    $("insert").addEventListener("click", () => put("insert"));
    $("replace").addEventListener("click", () => put("replace"));
    $("unpair").addEventListener("click", () => {
      savePairing(undefined);
      show("pairing");
    });
    start();
  }

  if (typeof Office !== "undefined" && Office.onReady) Office.onReady(setUp);
  else document.addEventListener("DOMContentLoaded", () => setUp(undefined));
})();
