// Chord's Office add-in: reading the selection or the whole document's text, and putting an answer
// back in, in Word, Excel and PowerPoint (Office.js). No build step: plain script, one global.
/* global Office, Word, Excel, PowerPoint */
"use strict";

(function () {
  /** "word", "excel" or "powerpoint" (or undefined outside those). */
  function hostOf(info) {
    const host = info && info.host;
    if (!host || typeof Office === "undefined" || !Office.HostType) return undefined;
    if (host === Office.HostType.Word) return "word";
    if (host === Office.HostType.Excel) return "excel";
    if (host === Office.HostType.PowerPoint) return "powerpoint";
    return undefined;
  }

  /** Cells as text: one row per line, cells separated by tabs. */
  function cellsText(values, address) {
    const rows = (values || []).map((row) => row.map((v) => (v === null || v === undefined ? "" : String(v))).join("\t"));
    return (address ? "Cells " + address + ":\n" : "") + rows.join("\n");
  }

  /** A Markdown table (| a | b |) or tab-separated lines in the answer, as rows of cells; else undefined. */
  function tableFrom(text) {
    const lines = String(text || "")
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => l !== "");
    if (lines.length === 0) return undefined;
    const piped = lines.filter((l) => l.startsWith("|") && l.endsWith("|") && l.length > 1);
    if (piped.length >= 2 && piped.length === lines.length) {
      const rows = piped.filter((l) => !/^\|[\s:|-]+\|$/.test(l)).map((l) => l.slice(1, -1).split("|").map((c) => c.trim()));
      return rows.length ? square(rows) : undefined;
    }
    if (lines.every((l) => l.includes("\t"))) return square(lines.map((l) => l.split("\t")));
    return undefined;
  }

  function square(rows) {
    const width = Math.max(...rows.map((r) => r.length));
    return rows.map((r) => r.concat(Array(width - r.length).fill("")));
  }

  // ---------- reading ----------

  async function wordText(kind) {
    return Word.run(async (ctx) => {
      const range = kind === "document" ? ctx.document.body : ctx.document.getSelection();
      range.load("text");
      await ctx.sync();
      return range.text || "";
    });
  }

  async function excelText(kind) {
    return Excel.run(async (ctx) => {
      const range = kind === "document" ? ctx.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject(true) : ctx.workbook.getSelectedRange();
      range.load("values,address");
      await ctx.sync();
      if (range.isNullObject) return "";
      return cellsText(range.values, range.address);
    });
  }

  const TEXT_SHAPES = ["GeometricShape", "TextBox", "Placeholder"];

  /** The text of each slide's text shapes ("Slide 2:" before each slide's text). */
  async function slidesText(ctx, slides) {
    slides.load("items");
    await ctx.sync();
    const all = slides.items.map((slide) => {
      slide.shapes.load("items/type");
      return slide;
    });
    await ctx.sync();
    const ranges = all.map((slide) => slide.shapes.items.filter((s) => TEXT_SHAPES.includes(s.type)).map((s) => s.textFrame.textRange.load("text")));
    try {
      await ctx.sync();
    } catch (err) {
      return ""; // a shape without text: nothing read (rare)
    }
    return ranges
      .map((list, i) => {
        const text = list.map((r) => r.text).filter((t) => t && t.trim()).join("\n");
        return text ? "Slide " + (i + 1) + ":\n" + text : "";
      })
      .filter(Boolean)
      .join("\n\n");
  }

  async function powerPointText(kind) {
    return PowerPoint.run(async (ctx) => (kind === "document" ? slidesText(ctx, ctx.presentation.slides) : slidesText(ctx, ctx.presentation.getSelectedSlides())));
  }

  /** The selected text, cells or slide's text, or the whole document's, sheet's or presentation's. */
  function readText(host, kind) {
    if (host === "word") return wordText(kind);
    if (host === "excel") return excelText(kind);
    if (host === "powerpoint") return powerPointText(kind);
    return Promise.resolve("");
  }

  // ---------- writing ----------

  async function wordPut(text, mode) {
    return Word.run(async (ctx) => {
      ctx.document.getSelection().insertText(text, mode === "replace" ? "Replace" : "After");
      await ctx.sync();
    });
  }

  /** Excel: a table goes into cells (from the selection's first cell, or the cell below it); other text into one cell. */
  async function excelPut(text, mode) {
    return Excel.run(async (ctx) => {
      const selected = ctx.workbook.getSelectedRange();
      const start = mode === "replace" ? selected.getCell(0, 0) : selected.getLastRow().getOffsetRange(1, 0).getCell(0, 0);
      const rows = tableFrom(text) || [[text]];
      start.getResizedRange(rows.length - 1, rows[0].length - 1).values = rows;
      await ctx.sync();
    });
  }

  function setSelected(text) {
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result.error ? result.error.message : "Couldn't change the selection."));
      });
    });
  }

  /** PowerPoint: Replace changes the selected text; Insert adds a text box to the selected slide. */
  async function powerPointPut(text, mode) {
    if (mode === "replace") return setSelected(text);
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.getSelectedSlides();
      slides.load("items");
      await ctx.sync();
      const slide = slides.items[0];
      if (!slide) throw new Error("Select a slide first.");
      slide.shapes.addTextBox(text, { left: 40, top: 60, width: 640, height: 360 });
      await ctx.sync();
    });
  }

  function putAnswer(host, text, mode) {
    if (host === "word") return wordPut(text, mode);
    if (host === "excel") return excelPut(text, mode);
    if (host === "powerpoint") return powerPointPut(text, mode);
    return Promise.reject(new Error("Open this in Word, Excel or PowerPoint."));
  }

  window.ChordOffice = { hostOf, readText, putAnswer, tableFrom, cellsText };
})();
