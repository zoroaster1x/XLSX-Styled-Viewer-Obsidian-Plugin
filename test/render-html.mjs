// Renders the two real workbooks to standalone HTML files for inspection.
// Run with: bun test/render-html.mjs

import { createRequire } from "node:module";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser, parseHTML } from "linkedom";

globalThis.DOMParser = DOMParser;

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const require = createRequire(import.meta.url);
const { readWorkbook } = require("../src/read.js");
const { createRenderer } = require("../src/render.js");

const FIXTURES = join(here, "fixtures");
const XY_PATH = process.env.XY_PATH || join(FIXTURES, "xy-list.xlsx");
const TT_PATH = process.env.TT_PATH || join(FIXTURES, "timetable.xlsx");
if (!existsSync(XY_PATH) || !existsSync(TT_PATH)) {
  console.log("No workbook fixtures found. Set XY_PATH and TT_PATH, or put copies in test/fixtures/.");
  process.exit(0);
}

const css = readFileSync(join(root, "styles.css"), "utf8");
const outDir = join(here, "out");
mkdirSync(outDir, { recursive: true });

function renderFile(name, path, opts) {
  const { document } = parseHTML("<html><body><div id=\"host\"></div></body></html>");
  const host = document.getElementById("host");
  const book = readWorkbook(readFileSync(path));
  const sheet = book.loadSheet(opts.sheetIndex || 0);
  const renderer = createRenderer({
    container: host,
    sheet,
    styles: book.styles,
    theme: book.theme,
    date1904: book.date1904,
    settings: Object.assign({
      sheetBackground: "white",
      showGridlines: true,
      showHeaders: true,
      maxRows: 2000,
      scale: 1,
    }, opts.settings || {}),
    filterState: new Map(opts.filterState || []),
    outlineCollapsed: new Map(),
  });
  if (opts.filter) renderer.setFilter(opts.filter[0], new Set(opts.filter[1]));

  const body = host.innerHTML;
  const html = "<!doctype html><html><head><meta charset=\"utf-8\"><title>"
    + name + "</title><style>body{margin:0;background:#888;padding:10px}"
    + " .xlsx-grid-host{height:800px}" + css + "</style></head><body>" + body + "</body></html>";
  writeFileSync(join(outDir, name + ".html"), html);
  const state = renderer.getState();
  console.log(name + ": " + state.visibleRows + "/" + state.totalRows + " rows, "
    + state.totalCols + " cols, " + state.filters.length + " filter columns, "
    + html.length + " bytes");
  renderer.destroy();
}

renderFile("timetable", TT_PATH, {});
renderFile("timetable-x1x2", TT_PATH, { filter: [2, ["X1/X2"]] });
renderFile("xy-list", XY_PATH, {});
