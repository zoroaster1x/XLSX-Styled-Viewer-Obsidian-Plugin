/*
 * XLSX Styled Viewer, an Obsidian plugin that renders .xlsx workbooks with
 * their real styling, frozen panes, autofilters and sheet tabs.
 *
 * Copyright (C) 2026 Zoroaster1x
 *
 * This program is free software: you can redistribute it and/or modify it under
 * the terms of the GNU General Public License as published by the Free Software
 * Foundation, either version 3 of the License, or (at your option) any later
 * version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
 * details.
 *
 * You should have received a copy of the GNU General Public License along with
 * this program. If not, see <https://www.gnu.org/licenses/>.
 */

// Local test runner. Uses bun plus linkedom so the parser and the renderer can
// run outside Obsidian. Run with: bun test/run.mjs

import { createRequire } from "node:module";
import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser, parseHTML } from "linkedom";

globalThis.DOMParser = DOMParser;

const require = createRequire(import.meta.url);
const { readWorkbook } = require("../src/read.js");
const { createRenderer } = require("../src/render.js");
const { formatValue, dateParts } = require("../src/numfmt.js");
const { applyTint } = require("../src/color.js");

// Workbook tests need two real .xlsx files. Put copies in test/fixtures/ or
// point XY_PATH and TT_PATH at them. The fixtures are not committed, since
// real timetables and class lists contain personal data.
const FIXTURES = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const XY_PATH = process.env.XY_PATH || join(FIXTURES, "xy-list.xlsx");
const TT_PATH = process.env.TT_PATH || join(FIXTURES, "timetable.xlsx");
const haveFixtures = existsSync(XY_PATH) && existsSync(TT_PATH);

let pass = 0;
let fail = 0;
function check(name, cond, extra) {
  if (cond) {
    pass++;
    console.log("ok   " + name);
  } else {
    fail++;
    console.log("FAIL " + name + (extra !== undefined ? "  -> " + extra : ""));
  }
}

// ---------- number formats ----------

check("fmt general int", formatValue(1234, "General").text === "1234");
check("fmt general float", formatValue(1234.5, "General").text === "1234.5");
check("fmt 0.00", formatValue(1.5, "0.00").text === "1.50");
check("fmt thousands", formatValue(1234567, "#,##0").text === "1,234,567");
check("fmt thousands decimals", formatValue(1234.5, "#,##0.00").text === "1,234.50");
check("fmt percent", formatValue(0.1234, "0.0%").text === "12.3%");
check("fmt percent plain", formatValue(0.5, "0%").text === "50%");
check("fmt negative parens", formatValue(-1234, "#,##0;(#,##0)").text === "(1,234)");
check("fmt negative minus", formatValue(-1234, "#,##0").text === "-1,234");
check("fmt zero section", formatValue(0, "#,##0;(#,##0);\"-\"").text === "-");
check("fmt red negative", formatValue(-5, "0;[Red]0").color === "#FF0000" && formatValue(-5, "0;[Red]0").text === "5");
check("fmt hash zero", formatValue(0, "#").text === "");
check("fmt text", formatValue(4.2, "0.0\" kg\"").text === "4.2 kg");
const d45000 = dateParts(45000, false);
check("fmt date parts", d45000.y === 2023 && d45000.mo === 3 && d45000.day === 15,
  JSON.stringify(d45000));
check("fmt date", formatValue(45000, "dd/mm/yyyy").text === "15/03/2023", formatValue(45000, "dd/mm/yyyy").text);
check("fmt time", formatValue(0.5, "h:mm AM/PM").text === "12:00 PM", formatValue(0.5, "h:mm AM/PM").text);
check("fmt elapsed", formatValue(1.25, "[h]:mm").text === "30:00", formatValue(1.25, "[h]:mm").text);
check("fmt builtin 14", formatValue(45000, "m/d/yy").text === "3/15/23", formatValue(45000, "m/d/yy").text);

// ---------- colour ----------

const tinted = applyTint("0E2841", 0.9);
check("tint lightens to blue grey", /^#[0-9A-F]{6}$/i.test(tinted) && parseInt(tinted.slice(1, 3), 16) < parseInt(tinted.slice(5, 7), 16), tinted);

if (!haveFixtures) {
  console.log('');
  console.log('No workbook fixtures found, skipping the workbook and renderer tests.');
  console.log('Set XY_PATH and TT_PATH, or put copies in test/fixtures/.');
} else {
// ---------- workbook tests ----------

const xy = readWorkbook(readFileSync(XY_PATH));
check("xy sheets", xy.sheets.length === 1 && xy.sheets[0].name === "Sheet1");
const xySheet = xy.loadSheet(0);
check("xy dims", xySheet.dims.r2 === 147 && xySheet.dims.c2 === 11,
  JSON.stringify(xySheet.dims));
check("xy no merges", xySheet.merges.length === 0);
check("xy first name cell", xySheet.cells.get("4:2") && typeof xySheet.cells.get("4:2").v === "string"
  && xySheet.cells.get("4:2").v.length > 0);
check("xy bottom section heading", xySheet.cells.get("81:2")
  && String(xySheet.cells.get("81:2").v).indexOf("No longer on list") === 0);
check("xy bottom section rows", (() => {
  let count = 0;
  for (let r = 83; r <= 113; r++) {
    const f = xySheet.cells.get(r + ":2");
    if (f && typeof f.v === "string" && f.v.length > 0) count++;
  }
  return count === 31;
})(), "expected 31 name rows");
const xyHeaderStyle = xy.styles.resolveXf(xySheet.cells.get("3:2").s);
const xyHeaderFill = xy.styles.resolveFill(xyHeaderStyle.fill);
check("xy header fill resolved", typeof xyHeaderFill === "string" && xyHeaderFill[0] === "#", xyHeaderFill);
check("xy header bold", xyHeaderStyle.font.bold === true);
check("xy header font", xyHeaderStyle.font.name === "Aptos Narrow");
const xyCellStyle = xy.styles.resolveXf(xySheet.cells.get("4:2").s);
check("xy cell thin border", xyCellStyle.borders.bottom && xyCellStyle.borders.bottom.width === 1
  && xyCellStyle.borders.left && xyCellStyle.borders.left.width === 1);

// ---------- timetable workbook ----------

const tt = readWorkbook(readFileSync(TT_PATH));
check("tt sheets", tt.sheets.length === 2 && tt.sheets[1].name === "Staff initials");
const ttSheet = tt.loadSheet(0);
check("tt dims", ttSheet.dims.r2 === 250 && ttSheet.dims.c2 === 24, JSON.stringify(ttSheet.dims));
check("tt merges", ttSheet.merges.length === 324, ttSheet.merges.length);
check("tt freeze", ttSheet.freeze && ttSheet.freeze.rows === 10 && ttSheet.freeze.cols === 0,
  JSON.stringify(ttSheet.freeze));
check("tt autofilter range", ttSheet.autoFilter && ttSheet.autoFilter.range.r1 === 10
  && ttSheet.autoFilter.range.c1 === 1 && ttSheet.autoFilter.range.r2 === 250 && ttSheet.autoFilter.range.c2 === 13,
  JSON.stringify(ttSheet.autoFilter && ttSheet.autoFilter.range));
check("tt filter hides button on col D", ttSheet.autoFilter.columns.get(3)
  && ttSheet.autoFilter.columns.get(3).showButton === false);
check("tt group cell value", ttSheet.cells.get("11:3") && typeof ttSheet.cells.get("11:3").v === "string"
  && ttSheet.cells.get("11:3").v.length > 0);
check("tt merged lecture text", ttSheet.cells.get("11:5") && String(ttSheet.cells.get("11:5").v).indexOf("LECTURE") === 0,
  JSON.stringify(ttSheet.cells.get("11:5") && ttSheet.cells.get("11:5").v));
check("tt internal hyperlink", (() => {
  const l = ttSheet.hyperlinks.get("4:11");
  return l && typeof l.location === "string" && l.location.indexOf("!") !== -1;
})(), JSON.stringify(ttSheet.hyperlinks.get("4:11")));
check("tt row 11 height", ttSheet.rows.get(11) && ttSheet.rows.get(11).ht === 46.5);
// The right half of merged D19:E19 has no value but carries borders and fill.
check("tt style-only cell kept", (() => {
  const e19 = ttSheet.cells.get("19:5");
  return e19 && e19.s === 45 && e19.v === null;
})(), JSON.stringify(ttSheet.cells.get("19:5")));
check("tt style-only cell borders resolve", (() => {
  const st = tt.styles.resolveXf(ttSheet.cells.get("19:5").s);
  return st.borders.right && st.borders.right.width === 2 && st.fill != null;
})());
check("tt medium border", (() => {
  const st = tt.styles.resolveXf(ttSheet.cells.get("11:5").s);
  return st.borders.bottom && st.borders.bottom.width === 2;
})());
check("tt wrap and centre", (() => {
  const st = tt.styles.resolveXf(ttSheet.cells.get("11:5").s);
  return st.align.wrap === true && st.align.h === "center";
})());
const staff = tt.loadSheet(1);
check("staff dims", staff.dims.r2 === 25 && staff.dims.c2 === 3);
check("staff mailto hyperlink", (() => {
  const l = staff.hyperlinks.get("2:3");
  return l && l.mode === "External" && String(l.target).indexOf("mailto:") === 0;
})(), JSON.stringify(staff.hyperlinks.get("2:3")));

// ---------- renderer ----------

const { document } = parseHTML("<html><body></body></html>");
const container = document.createElement("div");
document.body.appendChild(container);

const renderer = createRenderer({
  container,
  sheet: ttSheet,
  styles: tt.styles,
  theme: tt.theme,
  date1904: tt.date1904,
  settings: { sheetBackground: "white", showGridlines: true, showHeaders: true, maxRows: 2000, scale: 1 },
  filterState: new Map(),
  outlineCollapsed: new Map(),
});
let state = renderer.getState();
check("render visible rows", state.visibleRows === 250, state.visibleRows);
check("render cell count", container.querySelectorAll(".xlsx-cell").length > 4000,
  container.querySelectorAll(".xlsx-cell").length);
check("render filter buttons visible on 5 columns",
  container.querySelectorAll(".xlsx-filter-btn").length === 5,
  container.querySelectorAll(".xlsx-filter-btn").length);
check("render frozen row sticky top", (() => {
  const norm = (s) => String(s || "").replace(/\s+/g, "");
  const cells = container.querySelectorAll(".xlsx-cell");
  for (const el of Array.from(cells)) {
    const style = norm(el.getAttribute("style"));
    if (style.indexOf("grid-row:2/span1") !== -1) {
      return style.indexOf("position:sticky") !== -1 && style.indexOf("top:20px") !== -1;
    }
  }
  return false;
})());
check("render merged region exists", (() => {
  const merged = container.querySelectorAll(".xlsx-merged");
  return merged.length >= 300;
})(), container.querySelectorAll(".xlsx-merged").length);
check("render first row header sticky left", (() => {
  const norm = (s) => String(s || "").replace(/\s+/g, "");
  const head = container.querySelector(".xlsx-row-head");
  const style = norm(head && head.getAttribute("style"));
  return style.indexOf("position:sticky") !== -1 && style.indexOf("left:0") !== -1;
})());
check("render header background white mode", (() => {
  const sheet = container.querySelector(".xlsx-sheet");
  return sheet && sheet.getAttribute("class").indexOf("xlsx-white") !== -1;
})());
check("render: merged right border reaches the next cell", (() => {
  const cells = Array.from(container.querySelectorAll(".xlsx-cell"));
  const f19 = cells.find((el) => {
    const s = el.getAttribute("style") || "";
    return s.indexOf("grid-row:20 / span 1") !== -1 && s.indexOf("grid-column:7 / span 1") !== -1;
  });
  const style = f19 ? f19.getAttribute("style") : "";
  return style.indexOf("border-left:2px solid #000000") !== -1;
})(), (() => {
  const cells = Array.from(container.querySelectorAll(".xlsx-cell"));
  const f19 = cells.find((el) => {
    const s = el.getAttribute("style") || "";
    return s.indexOf("grid-row:20 / span 1") !== -1 && s.indexOf("grid-column:7 / span 1") !== -1;
  });
  return f19 ? f19.getAttribute("style") : "f19 not found";
})());
check("render overflow cell is visible", (() => {
  const cells = Array.from(container.querySelectorAll(".xlsx-cell"));
  const c4 = cells.find((el) => {
    const s = el.getAttribute("style") || "";
    return s.indexOf("grid-row:6 / span 1") !== -1 && s.indexOf("grid-column:4 / span 1") !== -1;
  });
  return c4 && c4.style.overflow === "visible";
})());
check("render frozen boundary keeps its line", (() => {
  const cells = Array.from(container.querySelectorAll(".xlsx-cell"));
  const lastFrozen = cells.find((el) => {
    const s = el.getAttribute("style") || "";
    return s.indexOf("grid-row:11 / span 1") !== -1 && s.indexOf("grid-column:4 / span 1") !== -1;
  });
  const firstAfter = cells.find((el) => {
    const s = el.getAttribute("style") || "";
    return s.indexOf("grid-row:12 / span 1") !== -1 && s.indexOf("grid-column:4 / span 1") !== -1;
  });
  const a = lastFrozen ? (lastFrozen.getAttribute("style") || "") : "";
  const b = firstAfter ? (firstAfter.getAttribute("style") || "") : "";
  return /border-bottom:\s*[12]px/.test(a) && b.indexOf("border-top:") === -1;
})());

// filter one value of the shortest filter column through the filter state
const filterColumns = renderer.getState().filters;
const candidates = filterColumns.filter((f) => f.values.length >= 2)
  .sort((a, b) => a.values.length - b.values.length);
const groupFilter = candidates[0] || filterColumns[0];
const chosen = groupFilter.values[0].value;
renderer.setFilter(groupFilter.colId, new Set([chosen]));
state = renderer.getState();
check("filter shrinks row list", state.visibleRows > 0 && state.visibleRows < 250, state.visibleRows);
check("filter keeps only the chosen value", (() => {
  const texts = Array.from(container.querySelectorAll(".xlsx-cell-text")).map((el) => el.textContent);
  const others = groupFilter.values.map((v) => v.value).filter((v) => v !== chosen);
  return texts.indexOf(chosen) !== -1 && others.every((v) => texts.indexOf(v) === -1);
})());
renderer.setFilter(groupFilter.colId, null);
check("clear filter restores rows", renderer.getState().visibleRows === 250);

// search
const hits = renderer.search(String(chosen));
check("search finds hits", hits > 0, hits);
const next = renderer.searchNext(1);
check("search next", next && next.index === 1 && next.count === hits, JSON.stringify(next));
renderer.search("");

// XY sheet renders too
const xyContainer = document.createElement("div");
document.body.appendChild(xyContainer);
const xyRenderer = createRenderer({
  container: xyContainer,
  sheet: xySheet,
  styles: xy.styles,
  theme: xy.theme,
  date1904: xy.date1904,
  settings: { sheetBackground: "white", showGridlines: true, showHeaders: true, maxRows: 2000, scale: 1 },
  filterState: new Map(),
  outlineCollapsed: new Map(),
});
check("xy render cells", xyContainer.querySelectorAll(".xlsx-cell").length > 1000,
  xyContainer.querySelectorAll(".xlsx-cell").length);
check("xy render no filter buttons", xyContainer.querySelectorAll(".xlsx-filter-btn").length === 0);
check("xy render bottom section", (() => {
  const texts = Array.from(xyContainer.querySelectorAll(".xlsx-cell-text")).map((el) => el.textContent);
  return texts.some((t) => t.indexOf("No longer on list") === 0) && texts.length > 600;
})());
xyRenderer.destroy();
renderer.destroy();

// Synthetic sheet with an autofilter on its first row: the filter button cell
// must be positioned so the arrow anchors to it.
const synthSheet = {
  name: "Synth",
  dims: { r1: 1, c1: 1, r2: 3, c2: 3 },
  defaultRowHeightPt: 15,
  defaultRowHeightPx: 20,
  defaultColWidthChars: null,
  cols: [],
  rows: new Map(),
  cells: new Map(),
  merges: [],
  freeze: null,
  autoFilter: { range: { r1: 1, c1: 1, r2: 3, c2: 3 }, columns: new Map() },
  hyperlinks: new Map(),
  showGridLines: true,
  outlinePr: { summaryBelow: true, summaryRight: true },
  conditional: [],
  tabColor: null,
};
synthSheet.cells.set("1:1", { r: 1, c: 1, s: 0, t: "s", v: "Head" });
synthSheet.cells.set("2:1", { r: 2, c: 1, s: 0, t: "s", v: "row" });
const synthHost = document.createElement("div");
document.body.appendChild(synthHost);
const synthRenderer = createRenderer({
  container: synthHost,
  sheet: synthSheet,
  styles: tt.styles,
  theme: tt.theme,
  date1904: false,
  settings: { sheetBackground: "white", showGridlines: true, showHeaders: true, maxRows: 0, scale: 1 },
  filterState: new Map(),
  outlineCollapsed: new Map(),
});
check("synth filter buttons", synthHost.querySelectorAll(".xlsx-filter-btn").length === 3,
  synthHost.querySelectorAll(".xlsx-filter-btn").length);
check("synth filter button cell positioned", (() => {
  const btn = synthHost.querySelector(".xlsx-filter-btn");
  return btn && btn.parentElement.getAttribute("style").indexOf("position:relative") !== -1;
})(), (() => {
  const btn = synthHost.querySelector(".xlsx-filter-btn");
  return btn ? btn.parentElement.getAttribute("style") : "no button";
})());
check("synth plain cell stays static", (() => {
  const cells = Array.from(synthHost.querySelectorAll(".xlsx-cell"));
  const row3 = cells.find((el) => (el.getAttribute("style") || "").indexOf("grid-row:4 / span 1") !== -1);
  return row3 && (row3.getAttribute("style") || "").indexOf("position:") === -1;
})());
synthRenderer.destroy();
}

console.log("");
console.log(pass + " passed, " + fail + " failed");
process.exit(fail === 0 ? 0 : 1);
