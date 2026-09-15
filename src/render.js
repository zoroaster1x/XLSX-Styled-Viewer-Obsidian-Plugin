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

"use strict";

// Renders one parsed sheet into a DOM grid: styled cells, merged ranges,
// frozen panes, gridlines, wrapped text, filter dropdowns and outline groups.

const { colToLetter, key, naturalCompare, parseRange, splitSheetRef } = require("./util");
const { formatValue } = require("./numfmt");
const { strongerBorder } = require("./styles");
const { isAutoColor, resolveColor } = require("./color");

const ROW_HEADER_WIDTH = 46;
const COL_HEADER_HEIGHT = 20;
const GUTTER_WIDTH = 14;
const PAD_X = 3;
const PAD_Y = 2;

// Stacking layers. Headers must sit above frozen cells, frozen cells above
// scrolling content, and frozen row numbers above the scrolling numbers.
const Z = {
  cell: 1,
  overflow: 2,
  scrollHead: 3,
  frozen: 4,
  overflowFrozen: 5,
  frozenHead: 7,
  colHead: 7,
  corner: 8,
};

function createRenderer(opts) {
  const container = opts.container;
  const doc = container.ownerDocument || document;
  const sheet = opts.sheet;
  const styles = opts.styles;
  const theme = opts.theme;
  const callbacks = {
    onNavigate: opts.onNavigate || function () {},
    onOpenExternal: opts.onOpenExternal || function (url) { window.open(url, "_blank"); },
    onFilterChange: opts.onFilterChange || function () {},
    onSelect: opts.onSelect || function () {},
    showContextMenu: opts.showContextMenu || null,
  };
  const settings = Object.assign({
    sheetBackground: "white",
    showGridlines: true,
    showHeaders: true,
    maxRows: 2000,
    scale: 1,
  }, opts.settings || {});
  const filterState = opts.filterState || new Map();
  const outlineCollapsed = opts.outlineCollapsed || new Map();
  const date1904 = Boolean(opts.date1904);

  let rootEl = null;
  let scrollEl = null;
  let gridEl = null;
  let popupEl = null;
  let destroyed = false;
  let itemEls = [];
  let resolveItemEls = new Map();
  let selectedEl = null;
  const selection = { ref: null, r: null, c: null, text: "" };
  let G = null;
  let styleCache = new Map();
  let displayCache = new Map();
  let measureCanvas = null;
  let measureCtx = null;
  const measureCache = new Map();

  // ---------- static per-run geometry ----------

  function styleIndexFor(r, c) {
    const cell = sheet.cells.get(key(r, c));
    if (cell && cell.s != null) return cell.s;
    const row = sheet.rows.get(r);
    if (row && row.customFormat && row.s != null) return row.s;
    const colDef = colDefFor(c);
    if (colDef && colDef.style) return colDef.style;
    return 0;
  }

  let colDefMap = null;
  function colDefFor(c) {
    if (!colDefMap) {
      colDefMap = new Array((G ? G.maxCol : sheet.dims.c2) + 2).fill(null);
      for (const def of sheet.cols) {
        for (let c2 = def.min; c2 <= def.max; c2++) {
          colDefMap[c2] = def;
          if (c2 >= colDefMap.length) break;
        }
      }
    }
    return colDefMap[c] || null;
  }

  function resolvedStyle(r, c) {
    const idx = styleIndexFor(r, c);
    const cacheKey = r + ":" + c;
    if (styleCache.has(cacheKey)) return styleCache.get(cacheKey);
    const st = styles.resolveXf(idx);
    styleCache.set(cacheKey, st);
    return st;
  }

  function cellValue(r, c) {
    const cell = sheet.cells.get(key(r, c));
    return cell ? cell.v : null;
  }

  function cellValueType(r, c) {
    const cell = sheet.cells.get(key(r, c));
    return cell ? cell.t : "n";
  }

  // Returns { text, color, extra } for the cell, applying number format and
  // conditional formatting on top of the raw value.
  function displayOf(r, c) {
    const cacheKey = r + ":" + c;
    if (displayCache.has(cacheKey)) return displayCache.get(cacheKey);
    const cell = sheet.cells.get(key(r, c));
    let out = { text: "", color: null, extra: null };
    if (cell) {
      const st = styles.resolveXf(styleIndexFor(r, c));
      const t = cell.t;
      if (t === "s" || t === "str" || t === "is" || t === "inlineStr" || t === "e") {
        out = { text: cell.v == null ? "" : String(cell.v), color: null, extra: null };
      } else {
        const formatted = formatValue(cell.v, st.fmt, { date1904 });
        out = { text: formatted.text, color: formatted.color || null, extra: null };
      }
      if (typeof cell.v === "number") {
        const cf = conditionalOverride(r, c, cell.v, out.text);
        if (cf) {
          out = Object.assign({}, out, { extra: cf });
          if (cf.fill) out.extraFill = cf.fill;
          if (cf.fontColor) out.color = cf.fontColor;
        }
      }
    }
    displayCache.set(cacheKey, out);
    return out;
  }

  function conditionalOverride(r, c, numValue, textValue) {
    if (!sheet.conditional.length) return null;
    for (const block of sheet.conditional) {
      let inside = false;
      for (const rg of block.ranges) {
        if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) { inside = true; break; }
      }
      if (!inside) continue;
      for (const rule of block.rules) {
        const hit = evalRule(rule, numValue, textValue);
        if (!hit) continue;
        if (rule.type === "colorScale" && rule.colorScale) {
          const color = colorScaleColor(rule.colorScale, numValue);
          if (color) return { fill: color };
        } else if (rule.type === "dataBar" && rule.dataBar) {
          const t = cfValuePosition(rule.dataBar.values, numValue);
          const color = resolveColor(rule.dataBar.color, theme, "#638EC6");
          return { bar: { t: t == null ? 0 : t, color } };
        } else if (rule.dxfId != null) {
          const dxf = styles.dxfStyle(rule.dxfId);
          if (dxf) {
            return {
              fill: dxf.fill,
              fontColor: dxf.font && dxf.font.color ? dxf.font.color : null,
              bold: dxf.font ? dxf.font.bold : false,
              italic: dxf.font ? dxf.font.italic : false,
            };
          }
        }
        if (rule.stopIfTrue) return null;
      }
    }
    return null;
  }

  function evalRule(rule, numValue, textValue) {
    switch (rule.type) {
      case "cellIs": {
        const v = parseFloat(rule.formulas[0]);
        if (isNaN(v)) return false;
        switch (rule.operator) {
          case "greaterThan": return numValue > v;
          case "greaterThanOrEqual": return numValue >= v;
          case "lessThan": return numValue < v;
          case "lessThanOrEqual": return numValue <= v;
          case "equal": return numValue === v;
          case "notEqual": return numValue !== v;
          case "between": {
            const v2 = parseFloat(rule.formulas[1]);
            return numValue >= Math.min(v, v2) && numValue <= Math.max(v, v2);
          }
          case "notBetween": {
            const v2 = parseFloat(rule.formulas[1]);
            return numValue < Math.min(v, v2) || numValue > Math.max(v, v2);
          }
          default: return false;
        }
      }
      case "containsText":
        return rule.text != null && String(textValue).toLowerCase().indexOf(String(rule.text).toLowerCase()) !== -1;
      case "notContainsText":
        return rule.text != null && String(textValue).toLowerCase().indexOf(String(rule.text).toLowerCase()) === -1;
      case "beginsWith":
        return rule.text != null && String(textValue).toLowerCase().startsWith(String(rule.text).toLowerCase());
      case "endsWith":
        return rule.text != null && String(textValue).toLowerCase().endsWith(String(rule.text).toLowerCase());
      case "containsBlanks":
        return String(textValue) === "";
      case "notContainsBlanks":
        return String(textValue) !== "";
      case "colorScale":
      case "dataBar":
        return true;
      default:
        return false;
    }
  }

  function cfValuePosition(cfvoList, value) {
    if (!cfvoList || cfvoList.length < 2) return null;
    const lo = cfvoNumber(cfvoList[0], value);
    const hi = cfvoNumber(cfvoList[cfvoList.length - 1], value);
    if (lo == null || hi == null || hi === lo) return 0;
    return Math.max(0, Math.min(1, (value - lo) / (hi - lo)));
  }

  function cfvoNumber(cfvo, fallbackValue) {
    if (cfvo.type === "num") return parseFloat(cfvo.val);
    if (cfvo.type === "min") return fallbackValue;
    if (cfvo.type === "max") return fallbackValue;
    if (cfvo.type === "percent") return parseFloat(cfvo.val);
    return parseFloat(cfvo.val);
  }

  function colorScaleColor(scale, value) {
    if (!scale.colors.length) return null;
    const positions = scale.values.map((v) => cfValuePosition(scale.values, value));
    void positions;
    const lo = scale.values[0];
    const hi = scale.values[scale.values.length - 1];
    const loN = cfvoValue(lo, value);
    const hiN = cfvoValue(hi, value);
    let t = 0;
    if (loN != null && hiN != null && hiN !== loN) t = Math.max(0, Math.min(1, (value - loN) / (hiN - loN)));
    if (scale.colors.length === 1) return resolveColor(scale.colors[0], theme, null);
    if (scale.colors.length === 2 || value <= loN) {
      return resolveColor(scale.colors[0], theme, null);
    }
    if (value >= hiN) return resolveColor(scale.colors[scale.colors.length - 1], theme, null);
    const mid = scale.values.length === 3 ? cfvoValue(scale.values[1], value) : null;
    if (scale.colors.length === 3 && mid != null && mid !== loN && value <= mid) {
      const tt = (value - loN) / (mid - loN);
      return mixColors(resolveColor(scale.colors[0], theme, "#fff"), resolveColor(scale.colors[1], theme, "#fff"), tt);
    }
    if (scale.colors.length === 3 && mid != null && mid !== hiN) {
      const tt = (value - mid) / (hiN - mid);
      return mixColors(resolveColor(scale.colors[1], theme, "#fff"), resolveColor(scale.colors[2], theme, "#fff"), tt);
    }
    return mixColors(resolveColor(scale.colors[0], theme, "#fff"), resolveColor(scale.colors[1], theme, "#fff"), t);
  }

  function cfvoValue(cfvo, value) {
    if (cfvo.type === "num") return parseFloat(cfvo.val);
    if (cfvo.type === "percent") return parseFloat(cfvo.val);
    if (cfvo.type === "min") return value;
    if (cfvo.type === "max") return value;
    return parseFloat(cfvo.val);
  }

  function mixColors(a, b, t) {
    const pa = parseCss(a);
    const pb = parseCss(b);
    if (!pa || !pb) return a;
    const mix = (x, y) => Math.round(x + (y - x) * t);
    return "rgb(" + mix(pa[0], pb[0]) + ", " + mix(pa[1], pb[1]) + ", " + mix(pa[2], pb[2]) + ")";
  }

  function parseCss(color) {
    if (!color) return null;
    if (color[0] === "#") {
      const hex = color.length === 4
        ? color[1] + color[1] + color[2] + color[2] + color[3] + color[3]
        : color.slice(1);
      return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
    }
    const m = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/.exec(color);
    if (m) return [Number(m[1]), Number(m[2]), Number(m[3])];
    return null;
  }

  // Automatic text colour on a coloured fill: black on light fills, white on
  // dark ones, so multicolour sheets stay readable in either theme.
  function autoTextColorForFill(fillCss) {
    const rgb = parseCss(fillCss);
    if (!rgb) return null;
    const lum = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255;
    return lum > 0.55 ? "#000000" : "#FFFFFF";
  }

  // ---------- visibility ----------

  function prepare() {
    styleCache = new Map();
    displayCache = new Map();
    colDefMap = null;

    const dims = sheet.dims;
    const maxRow = Math.max(1, dims.r2);
    const maxCol = Math.max(1, dims.c2);
    // Excel always shows the grid from row 1 and column A, even when the used
    // range starts further in.
    const usedR1 = 1;
    const usedC1 = 1;

    const hiddenRows0 = new Set();
    const hiddenCols0 = new Set();
    for (const [r, row] of sheet.rows) if (row.hidden) hiddenRows0.add(r);
    for (const def of sheet.cols) {
      if (!def.hidden) continue;
      for (let c = def.min; c <= def.max; c++) hiddenCols0.add(c);
    }
    // Also treat zero width columns as hidden.
    for (const def of sheet.cols) {
      if (def.width != null && def.width <= 0) {
        for (let c = def.min; c <= def.max; c++) hiddenCols0.add(c);
      }
    }

    const hiddenByFilter = new Set();
    const filterRange = sheet.autoFilter ? sheet.autoFilter.range : null;
    if (filterRange) {
      for (const [colId, sel] of filterState) {
        if (sel == null) continue;
        const col = filterRange.c1 + colId;
        const custom = sel && sel.custom ? sel.custom : null;
        const set = sel instanceof Set ? sel : null;
        for (let r = filterRange.r1 + 1; r <= filterRange.r2 && r <= maxRow; r++) {
          const text = displayTextRaw(r, col);
          let matched;
          if (custom) matched = matchesCustom(custom, text, r, col);
          else if (set) matched = set.has(text);
          else matched = true;
          if (!matched) hiddenByFilter.add(r);
        }
      }
    }

    const rowGroups = buildOutlineGroups("row");
    const colGroups = buildOutlineGroups("col");
    const hiddenByOutline = new Set();
    const hiddenColsByOutline = new Set();
    for (const group of rowGroups) {
      if (!outlineCollapsed.get(group.key)) continue;
      for (let r = group.start; r <= group.end; r++) hiddenByOutline.add(r);
    }
    for (const group of colGroups) {
      if (!outlineCollapsed.get(group.key)) continue;
      for (let c = group.start; c <= group.end; c++) hiddenColsByOutline.add(c);
    }

    const hiddenRows = new Set([...hiddenRows0, ...hiddenByFilter, ...hiddenByOutline]);
    const hiddenCols = new Set([...hiddenCols0, ...hiddenColsByOutline]);

    const maxRowsSetting = settings.maxRows > 0 ? settings.maxRows : 0;
    const renderRowEnd = maxRowsSetting ? Math.min(maxRow, maxRowsSetting) : maxRow;

    const visRows = [];
    const visRowIndex = new Map();
    for (let r = usedR1; r <= renderRowEnd; r++) {
      if (hiddenRows.has(r)) continue;
      visRowIndex.set(r, visRows.length);
      visRows.push(r);
    }
    const visCols = [];
    const visColIndex = new Map();
    for (let c = usedC1; c <= maxCol; c++) {
      if (hiddenCols.has(c)) continue;
      visColIndex.set(c, visCols.length);
      visCols.push(c);
    }

    const colWidths = new Array(maxCol + 2).fill(0);
    const defaultColWidth = sheet.defaultColWidthChars != null
      ? Math.round(sheet.defaultColWidthChars * 7) + 5
      : 64;
    for (let c = 1; c <= maxCol; c++) {
      const def = colDefFor(c);
      colWidths[c] = def && def.width != null ? Math.round(def.width * 7) + 5 : defaultColWidth;
    }

    const rowHeights = new Array(maxRow + 2).fill(0);
    for (let r = 1; r <= maxRow; r++) {
      const row = sheet.rows.get(r);
      rowHeights[r] = row && row.ht != null && row.ht > 0
        ? Math.round((row.ht * 96) / 72 * 100) / 100
        : sheet.defaultRowHeightPx;
    }

    const colLeftData = new Array(visCols.length + 1).fill(0);
    for (let j = 0; j < visCols.length; j++) {
      colLeftData[j + 1] = colLeftData[j] + colWidths[visCols[j]];
    }
    const rowTopData = new Array(visRows.length + 1).fill(0);
    for (let i = 0; i < visRows.length; i++) {
      rowTopData[i + 1] = rowTopData[i] + rowHeights[visRows[i]];
    }

    const coverMap = new Map();
    for (const merge of sheet.merges) {
      for (let r = merge.r1; r <= Math.min(merge.r2, maxRow); r++) {
        for (let c = merge.c1; c <= Math.min(merge.c2, maxCol); c++) {
          coverMap.set(key(r, c), merge);
        }
      }
    }

    G = {
      maxRow, maxCol, usedR1, usedC1, usedR2: maxRow, usedC2: maxCol,
      hiddenRows, hiddenCols, hiddenRows0, hiddenByFilter, hiddenByOutline,
      hiddenCols0, hiddenColsByOutline,
      visRows, visRowIndex, visCols, visColIndex,
      colWidths, rowHeights, colLeftData, rowTopData,
      coverMap, filterRange, rowGroups, colGroups,
      renderRowEnd, rowCapped: renderRowEnd < maxRow,
      hasRowOutlines: rowGroups.length > 0,
      hasColOutlines: colGroups.length > 0,
    };
  }

  function buildOutlineGroups(kind) {
    const groups = [];
    if (kind === "row") {
      let maxLevel = 0;
      for (const row of sheet.rows.values()) if (row.outline > maxLevel) maxLevel = row.outline;
      for (let level = 1; level <= maxLevel; level++) {
        let start = null;
        for (let r = 1; r <= G_maxRowHint(); r++) {
          const row = sheet.rows.get(r);
          const lvl = row ? row.outline : 0;
          if (lvl >= level) {
            if (start == null) start = r;
          } else if (start != null) {
            groups.push({ kind, level, start, end: r - 1, key: "row:" + level + ":" + start });
            start = null;
          }
        }
        if (start != null) {
          groups.push({ kind, level, start, end: G_maxRowHint(), key: "row:" + level + ":" + start });
        }
      }
    } else {
      const levels = new Map();
      for (const def of sheet.cols) {
        if (def.outline > 0) levels.set(def.min, def.outline);
      }
      let maxLevel = 0;
      for (const lvl of levels.values()) if (lvl > maxLevel) maxLevel = lvl;
      for (let level = 1; level <= maxLevel; level++) {
        let start = null;
        const maxCol = G_maxColHint();
        for (let c = 1; c <= maxCol; c++) {
          const def = colDefFor(c);
          const lvl = def ? def.outline : 0;
          if (lvl >= level) {
            if (start == null) start = c;
          } else if (start != null) {
            groups.push({ kind, level, start, end: c - 1, key: "col:" + level + ":" + start });
            start = null;
          }
        }
        if (start != null) groups.push({ kind, level, start, end: maxCol, key: "col:" + level + ":" + start });
      }
    }
    return groups;
  }

  // Small indirection so buildOutlineGroups can run before G exists.
  function G_maxRowHint() {
    const dims = sheet.dims;
    let max = dims.r2 || 1;
    for (const r of sheet.rows.keys()) if (r > max) max = r;
    return max;
  }
  function G_maxColHint() {
    const dims = sheet.dims;
    let max = dims.c2 || 1;
    for (const merge of sheet.merges) if (merge.c2 > max) max = merge.c2;
    return max;
  }

  function displayTextRaw(r, c) {
    return displayOf(r, c).text;
  }

  function matchesCustom(custom, text, r, c) {
    const num = Number(text);
    let result = null;
    for (const filter of custom.filters) {
      let hit;
      const val = filter.val != null ? filter.val : "";
      const numVal = Number(val);
      switch (filter.operator) {
        case "greaterThan": hit = !isNaN(num) && num > numVal; break;
        case "greaterThanOrEqual": hit = !isNaN(num) && num >= numVal; break;
        case "lessThan": hit = !isNaN(num) && num < numVal; break;
        case "lessThanOrEqual": hit = !isNaN(num) && num <= numVal; break;
        case "equal": hit = text === val; break;
        case "notEqual": hit = text !== val; break;
        default: hit = true;
      }
      if (result == null) result = hit;
      else result = custom.and ? (result && hit) : (result || hit);
    }
    if (result == null) return true;
    if (custom.and) return result;
    return result;
  }

  // ---------- text measuring ----------

  function fontString(style) {
    return (style.font.italic ? "italic " : "") + (style.font.bold ? "bold " : "")
      + style.font.sizePx + "px " + style.font.family;
  }

  function measureText(text, style) {
    const font = fontString(style);
    const cacheKey = font + "\u0000" + text;
    if (measureCache.has(cacheKey)) return measureCache.get(cacheKey);
    try {
      if (!measureCtx) {
        measureCanvas = doc.createElement("canvas");
        measureCtx = measureCanvas.getContext ? measureCanvas.getContext("2d") : null;
        if (!measureCtx) return null;
      }
      measureCtx.font = font;
      const w = measureCtx.measureText(text).width;
      measureCache.set(cacheKey, w);
      return w;
    } catch (err) {
      return null;
    }
  }

  // ---------- borders ----------

  function rawBorder(r, c, side) {
    const st = styles.resolveXf(styleIndexFor(r, c));
    return st.borders[side];
  }

  function cellEdgeBorder(r, c, side) {
    const merge = G.coverMap.get(key(r, c));
    if (!merge) return rawBorder(r, c, side);
    let best = null;
    if (side === "left" && c === merge.c1) {
      for (let rr = merge.r1; rr <= merge.r2; rr++) best = strongerBorder(best, rawBorder(rr, merge.c1, "left"));
      return best;
    }
    if (side === "right" && c === merge.c2) {
      for (let rr = merge.r1; rr <= merge.r2; rr++) best = strongerBorder(best, rawBorder(rr, merge.c2, "right"));
      return best;
    }
    if (side === "top" && r === merge.r1) {
      for (let cc = merge.c1; cc <= merge.c2; cc++) best = strongerBorder(best, rawBorder(merge.r1, cc, "top"));
      return best;
    }
    if (side === "bottom" && r === merge.r2) {
      for (let cc = merge.c1; cc <= merge.c2; cc++) best = strongerBorder(best, rawBorder(merge.r2, cc, "bottom"));
      return best;
    }
    return null;
  }

  function gridlineSpec() {
    return { width: 1, css: "solid", color: "var(--xlsx-gridline)", weight: 0, style: "thin" };
  }

  function pickEdge(r1, c1, r2, c2, side) {
    let best = null;
    if (side === "left" || side === "right") {
      for (let r = r1; r <= r2; r++) {
        best = strongerBorder(best, cellEdgeBorder(r, c1, side));
      }
      if (side === "left" && c1 > G.usedC1) {
        for (let r = r1; r <= r2; r++) best = strongerBorder(best, cellEdgeBorder(r, c1 - 1, "right"));
      }
      if (side === "right" && c2 < G.usedC2) {
        for (let r = r1; r <= r2; r++) best = strongerBorder(best, cellEdgeBorder(r, c2 + 1, "left"));
      }
    } else {
      for (let c = c1; c <= c2; c++) {
        best = strongerBorder(best, cellEdgeBorder(r1, c, side));
      }
      if (side === "top" && r1 > G.usedR1) {
        for (let c = c1; c <= c2; c++) best = strongerBorder(best, cellEdgeBorder(r1 - 1, c, "bottom"));
      }
      if (side === "bottom" && r2 < G.usedR2) {
        for (let c = c1; c <= c2; c++) best = strongerBorder(best, cellEdgeBorder(r2 + 1, c, "top"));
      }
    }
    if (!best && sheet.showGridLines !== false && settings.showGridlines) best = gridlineSpec();
    return best;
  }

  function applyBorder(el, which, spec, force) {
    if (!spec) return;
    if (!force && which !== "left" && which !== "top") return;
    el.style["border" + which.charAt(0).toUpperCase() + which.slice(1)] = spec.width + "px " + spec.css + " " + spec.color;
  }

  // ---------- build ----------

  function render(preserveScroll) {
    if (destroyed) return;
    const prev = preserveScroll && scrollEl ? { top: scrollEl.scrollTop, left: scrollEl.scrollLeft } : null;
    if (!preserveScroll) closePopup();
    prepare();

    const showHeaders = settings.showHeaders;
    const gutter = G.hasRowOutlines || G.hasColOutlines ? GUTTER_WIDTH : 0;
    const rowHeaderW = showHeaders ? ROW_HEADER_WIDTH : 0;
    const colHeaderH = showHeaders ? COL_HEADER_HEIGHT : 0;
    const colBase = (gutter ? 1 : 0) + (showHeaders ? 1 : 0);

    const templateCols = [];
    if (gutter) templateCols.push(GUTTER_WIDTH + "px");
    if (showHeaders) templateCols.push(ROW_HEADER_WIDTH + "px");
    for (const c of G.visCols) templateCols.push(G.colWidths[c] + "px");
    templateCols.push("minmax(40px, 1fr)");

    const templateRows = [];
    if (showHeaders) templateRows.push(COL_HEADER_HEIGHT + "px");
    for (const r of G.visRows) templateRows.push(G.rowHeights[r] + "px");

    const frag = doc.createDocumentFragment();
    itemEls = [];
    resolveItemEls = new Map();

    const freezeRows = sheet.freeze ? sheet.freeze.rows : 0;
    const freezeCols = sheet.freeze ? sheet.freeze.cols : 0;

    const rowTopOf = (r) => {
      const i = G.visRowIndex.get(r);
      return i == null ? null : colHeaderH + G.rowTopData[i];
    };
    const colLeftOf = (c) => {
      const j = G.visColIndex.get(c);
      return j == null ? null : rowHeaderW + gutter + G.colLeftData[j];
    };

    // Column headers and row headers.
    if (showHeaders) {
      const corner = doc.createElement("div");
      corner.className = "xlsx-corner";
      corner.style.gridColumn = "1 / span " + Math.max(colBase, 1);
      corner.style.gridRow = "1";
      corner.style.position = "sticky";
      corner.style.top = "0";
      corner.style.left = "0";
      corner.style.zIndex = String(Z.corner);
      frag.appendChild(corner);

      for (let j = 0; j < G.visCols.length; j++) {
        const c = G.visCols[j];
        const head = doc.createElement("div");
        head.className = "xlsx-col-head";
        head.textContent = colToLetter(c);
        head.style.gridColumn = String(colBase + j + 1);
        head.style.gridRow = "1";
        head.style.position = "sticky";
        head.style.top = "0";
        head.style.zIndex = String(Z.colHead);
        if (c <= freezeCols) head.style.left = (rowHeaderW + gutter + G.colLeftData[j]) + "px";
        frag.appendChild(head);
      }
    }

    for (let i = 0; i < G.visRows.length; i++) {
      const r = G.visRows[i];
      const rowHead = doc.createElement("div");
      if (showHeaders) {
        rowHead.className = "xlsx-row-head";
        rowHead.textContent = String(r);
        rowHead.style.gridColumn = String(gutter + 1);
        rowHead.style.gridRow = String(i + 2);
        rowHead.style.position = "sticky";
        rowHead.style.left = "0";
        rowHead.style.zIndex = String(r <= freezeRows ? Z.frozenHead : Z.scrollHead);
        if (r <= freezeRows) rowHead.style.top = (colHeaderH + G.rowTopData[i]) + "px";
        frag.appendChild(rowHead);
      }
    }

    // Outline gutter buttons.
    if (G.hasRowOutlines) {
      for (const group of G.rowGroups) {
        const controlRow = sheet.outlinePr.summaryBelow
          ? Math.min(group.end + 1, G.maxRow)
          : Math.max(group.start - 1, 1);
        let target = controlRow;
        if (!G.visRowIndex.has(target)) {
          target = null;
          for (let r = group.start; r <= group.end; r++) {
            if (G.visRowIndex.has(r)) { target = r; break; }
          }
        }
        if (target == null) continue;
        const btn = doc.createElement("div");
        btn.className = "xlsx-outline-btn";
        btn.textContent = outlineCollapsed.get(group.key) ? "+" : "-";
        btn.style.gridColumn = "1";
        btn.style.gridRow = String(G.visRowIndex.get(target) + 2);
        btn.style.position = "sticky";
        btn.style.left = "0";
        btn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          toggleOutline(group.key);
        });
        frag.appendChild(btn);
      }
    }
    if (G.hasColOutlines) {
      for (const group of G.colGroups) {
        const controlCol = sheet.outlinePr.summaryRight ? Math.min(group.end + 1, G.maxCol) : Math.max(group.start - 1, 1);
        let target = controlCol;
        if (!G.visColIndex.has(target)) {
          target = null;
          for (let c = group.start; c <= group.end; c++) {
            if (G.visColIndex.has(c)) { target = c; break; }
          }
        }
        if (target == null) continue;
        const btn = doc.createElement("div");
        btn.className = "xlsx-outline-btn";
        btn.textContent = outlineCollapsed.get(group.key) ? "+" : "-";
        btn.style.gridColumn = String(colBase + G.visColIndex.get(target) + 1);
        btn.style.gridRow = "1";
        btn.style.position = "sticky";
        btn.style.top = "0";
        btn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          toggleOutline(group.key);
        });
        frag.appendChild(btn);
      }
    }

    // Data cells.
    for (let i = 0; i < G.visRows.length; i++) {
      const r = G.visRows[i];
      for (let j = 0; j < G.visCols.length; j++) {
        const c = G.visCols[j];
        const merge = G.coverMap.get(key(r, c));
        if (merge && (merge.r1 !== r || merge.c1 !== c)) continue;
        const r2 = merge ? Math.min(merge.r2, G.maxRow) : r;
        const c2 = merge ? Math.min(merge.c2, G.maxCol) : c;
        buildItem(frag, r, c, r2, c2, Boolean(merge), i, j, colBase, gutter, rowHeaderW, colHeaderH, freezeRows, freezeCols, rowTopOf, colLeftOf);
      }
    }

    // Rebuild root.
    rootEl = doc.createElement("div");
    rootEl.className = "xlsx-sheet" + (settings.sheetBackground === "white" ? " xlsx-white" : "");
    scrollEl = doc.createElement("div");
    scrollEl.className = "xlsx-sheet-scroll";
    gridEl = doc.createElement("div");
    gridEl.className = "xlsx-sheet-grid";
    gridEl.style.position = "relative";
    gridEl.style.gridTemplateColumns = templateCols.join(" ");
    gridEl.style.gridTemplateRows = templateRows.join(" ");
    if (settings.scale && settings.scale !== 1) gridEl.style.zoom = String(settings.scale);
    gridEl.appendChild(frag);
    scrollEl.appendChild(gridEl);
    rootEl.appendChild(scrollEl);

    if (G.rowCapped) {
      const note = doc.createElement("div");
      note.className = "xlsx-cap-note";
      note.textContent = "Showing rows 1 to " + G.renderRowEnd + " of " + G.maxRow
        + ". Change the row limit in the plugin settings to see more.";
      rootEl.appendChild(note);
    }

    container.textContent = "";
    container.appendChild(rootEl);

    scrollEl.addEventListener("click", onGridClick);
    gridEl.addEventListener("contextmenu", onGridContext);
    restoreSelection();

    if (prev) {
      scrollEl.scrollTop = prev.top;
      scrollEl.scrollLeft = prev.left;
    }
  }

  function buildItem(frag, r, c, r2, c2, isMerge, i, j, colBase, gutter, rowHeaderW, colHeaderH, freezeRows, freezeCols, rowTopOf, colLeftOf) {
    const el = doc.createElement("div");
    el.className = "xlsx-cell";
    let spanRows = 0;
    for (let rr = r; rr <= r2; rr++) if (G.visRowIndex.has(rr)) spanRows++;
    let spanCols = 0;
    for (let cc = c; cc <= c2; cc++) if (G.visColIndex.has(cc)) spanCols++;
    if (!spanRows || !spanCols) return;
    el.style.gridRow = (i + 2) + " / span " + spanRows;
    el.style.gridColumn = (colBase + j + 1) + " / span " + spanCols;
    if (isMerge) el.classList.add("xlsx-merged");

    const style = resolvedStyle(r, c);
    const disp = displayOf(r, c);
    const extra = disp.extra || null;
    const value = cellValue(r, c);
    const hasValue = value != null && disp.text !== "";

    let fillCss = styles.resolveFill(style.fill);
    if (extra && extra.fill) fillCss = extra.fill;
    const isGradient = style.fill && style.fill.type === "gradient" && !(extra && extra.fill);

    const stickyTop = r <= freezeRows ? rowTopOf(r) : null;
    const stickyLeft = c <= freezeCols ? colLeftOf(c) : null;
    const isSticky = stickyTop != null || stickyLeft != null;
    if (stickyTop != null) el.style.top = stickyTop + "px";
    if (stickyLeft != null) el.style.left = stickyLeft + "px";
    // Plain cells stay static and unstacked on purpose: thousands of stacking
    // contexts make scrolling stutter.

    if (fillCss) el.style.backgroundColor = fillCss;
    else if (isGradient) el.style.backgroundImage = styles.gradientCss(style.fill);
    else if (isSticky) el.style.backgroundColor = "var(--xlsx-sheet-bg)";

    // Text layout.
    el.style.display = "flex";
    el.style.flexDirection = "column";
    const vAlign = style.align.v === "center" ? "center" : style.align.v === "top" ? "flex-start" : "flex-end";
    el.style.justifyContent = vAlign;
    let hAlign = style.align.h;
    if (!hAlign) {
      const t = cellValueType(r, c);
      hAlign = (typeof value === "number") ? "right" : (typeof value === "boolean" || t === "e") ? "center" : "left";
    }
    el.style.alignItems = hAlign === "center" ? "center" : hAlign === "right" ? "flex-end" : "flex-start";
    const indent = style.align.indent || 0;
    el.style.padding = PAD_Y + "px " + (PAD_X + indent * 9) + "px " + PAD_Y + "px " + (hAlign === "right" && indent ? (PAD_X + indent * 9) : PAD_X) + "px";

    // Borders. Left and top are drawn by every item; right and bottom only at
    // the edge of the used range, so shared lines are drawn once. The frozen
    // boundary keeps its line while the sheet scrolls.
    const freezeBottom = freezeRows > 0 && r2 === freezeRows;
    const freezeRight = freezeCols > 0 && c2 === freezeCols;
    const afterFreezeRow = freezeRows > 0 && r === freezeRows + 1;
    const afterFreezeCol = freezeCols > 0 && c === freezeCols + 1;
    if (!afterFreezeCol) applyBorder(el, "left", pickEdge(r, c, r2, c2, "left"), true);
    if (!afterFreezeRow) applyBorder(el, "top", pickEdge(r, c, r2, c2, "top"), true);
    if (freezeRight) applyBorder(el, "right", pickEdge(r, c, r2, c2, "right"), true);
    if (freezeBottom) applyBorder(el, "bottom", pickEdge(r, c, r2, c2, "bottom"), true);
    if (!freezeRight && c2 >= G.usedC2) applyBorder(el, "right", pickEdge(r, c, r2, c2, "right"), true);
    if (!freezeBottom && r2 >= G.usedR2) applyBorder(el, "bottom", pickEdge(r, c, r2, c2, "bottom"), true);

    // Conditional data bar.
    if (extra && extra.bar) {
      const bar = doc.createElement("div");
      bar.className = "xlsx-databar";
      bar.style.width = Math.round(extra.bar.t * 100) + "%";
      bar.style.backgroundColor = extra.bar.color;
      el.appendChild(bar);
    }

    let overflows = false;
    let hasFilterButton = false;
    if (hasValue) {
      const span = doc.createElement("div");
      span.className = "xlsx-cell-text";
      span.textContent = disp.text;
      span.style.fontFamily = style.font.family;
      span.style.fontSize = style.font.sizePx + "px";
      if (style.font.bold || (extra && extra.bold)) span.style.fontWeight = "700";
      if (style.font.italic || (extra && extra.italic)) span.style.fontStyle = "italic";
      if (style.font.underline) {
        span.style.textDecoration = "underline";
        if (style.font.underline === "double") span.style.textDecorationStyle = "double";
      }
      if (style.font.strike) span.style.textDecoration = "line-through";

      let color = disp.color || styles.resolveFontColor(style.font.color, null);
      const autoFont = !disp.color && isAutoColor(style.font.color) && !extra;
      if (autoFont && fillCss) color = autoTextColorForFill(fillCss);
      else if (autoFont && settings.sheetBackground === "theme") color = null;
      if (color) span.style.color = color;

      const wrap = style.align.wrap;
      if (wrap || isMerge) {
        span.classList.add("xlsx-wrap");
        span.style.whiteSpace = "pre-wrap";
        span.style.overflowWrap = "break-word";
        span.style.maxWidth = "100%";
      } else {
        span.style.whiteSpace = "pre";
        // Text overflows into empty neighbours the way Excel does, and is
        // clipped at the first non empty cell. max-content width is needed
        // because flex items would otherwise shrink to the cell width.
        el.style.overflow = "visible";
        overflows = true;
        span.style.width = "max-content";
        const limit = findBlocker(r, c);
        if (limit != null) {
          const own = G.colLeftData[G.visColIndex.get(c)];
          span.style.maxWidth = Math.max(0, limit - own - PAD_X) + "px";
          span.style.overflow = "hidden";
        } else {
          span.style.overflow = "visible";
        }
      }

      // Numbers that do not fit show hashes, like Excel.
      if (typeof value === "number" && !wrap && !isMerge) {
        const avail = G.colWidths[c] - PAD_X * 2 - indent * 9;
        const est = measureText(disp.text, style);
        const hashW = measureText("#", style);
        if (est != null && hashW != null && hashW > 0 && est > avail) {
          span.textContent = "#".repeat(Math.max(1, Math.floor(avail / hashW)));
        }
      }

      el.appendChild(span);
    }
    // Only content cells need clipping; empty cells skip the extra paint pass.
    if (!overflows && hasValue) el.style.overflow = "hidden";

    // Positioning is only applied where it is needed (sticky freeze, text
    // overflow, data bars, filter buttons). Everything else stays static so
    // scrolling does not maintain thousands of stacking contexts.
    if (isSticky) {
      el.style.position = "sticky";
      el.style.zIndex = String(overflows ? Z.overflowFrozen : Z.frozen);
    } else if (overflows) {
      el.style.position = "relative";
      el.style.zIndex = String(Z.overflow);
    } else if ((extra && extra.bar) || hasFilterButton) {
      el.style.position = "relative";
    }

    // Hyperlinks.
    const link = sheet.hyperlinks.get(key(r, c));
    if (link) {
      el.classList.add("xlsx-link");
      el.dataset.linkKey = key(r, c);
      // The theme link colour is only safe on cells without a fill.
      if (!fillCss && !disp.color) el.classList.add("xlsx-link-default");
    }

    // Filter buttons on the header row of the autofilter range.
    if (G.filterRange && r === G.filterRange.r1 && c >= G.filterRange.c1 && c <= G.filterRange.c2) {
      const colId = c - G.filterRange.c1;
      const def = sheet.autoFilter.columns.get(colId);
      if (!def || def.showButton !== false) {
        hasFilterButton = true;
        const btn = doc.createElement("div");
        btn.className = "xlsx-filter-btn";
        btn.title = "Filter column " + colToLetter(c);
        btn.textContent = "v";
        btn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          openFilterPopup(colId, btn);
        });
        el.appendChild(btn);
        if (filterState.has(colId) && filterState.get(colId) != null) el.classList.add("xlsx-filtered");
      }
    }

    frag.appendChild(el);
    itemEls.push({ el, r, c, text: disp.text });
    resolveItemEls.set(key(r, c), el);
  }

  function findBlocker(r, c) {
    for (let cc = c + 1; cc <= G.usedC2; cc++) {
      if (!G.visColIndex.has(cc)) continue;
      const merge = G.coverMap.get(key(r, cc));
      if (merge) return G.colLeftData[G.visColIndex.get(cc)];
      if (displayOf(r, cc).text !== "") return G.colLeftData[G.visColIndex.get(cc)];
    }
    return null;
  }

  function toggleOutline(groupKey) {
    outlineCollapsed.set(groupKey, !outlineCollapsed.get(groupKey));
    render(true);
  }

  // ---------- interactions ----------

  function onGridClick(ev) {
    const target = ev.target;
    const linkEl = closestWithClass(target, "xlsx-link");
    if (linkEl) {
      const entry = sheet.hyperlinks.get(linkEl.dataset.linkKey);
      if (entry) {
        ev.preventDefault();
        if (entry.location) {
          const info = splitSheetRef(entry.location);
          callbacks.onNavigate(info.sheet, info.ref);
        } else if (entry.target) {
          callbacks.onOpenExternal(entry.target);
        }
      }
    }
    const cellEl = closestWithClass(target, "xlsx-cell");
    if (cellEl) selectElement(cellEl);
  }

  function selectElement(el) {
    if (selectedEl === el) return;
    if (selectedEl) selectedEl.classList.remove("xlsx-selected");
    selectedEl = el;
    el.classList.add("xlsx-selected");
    const item = itemForElement(el);
    if (item) {
      selection.ref = colToLetter(item.c) + item.r;
      selection.r = item.r;
      selection.c = item.c;
      selection.text = item.text;
    }
    callbacks.onSelect({ ref: selection.ref, r: selection.r, c: selection.c, text: selection.text });
  }

  function getSelection() {
    return { ref: selection.ref, r: selection.r, c: selection.c, text: selection.text };
  }

  function restoreSelection() {
    if (!selection.ref || !resolveItemEls) return;
    const el = resolveItemEls.get(key(selection.r, selection.c));
    if (el) {
      selectedEl = el;
      el.classList.add("xlsx-selected");
    }
  }

  function onGridContext(ev) {
    const target = ev.target;
    const cellEl = closestWithClass(target, "xlsx-cell");
    if (!cellEl || !callbacks.showContextMenu) return;
    ev.preventDefault();
    const item = itemForElement(cellEl);
    if (!item) return;
    callbacks.showContextMenu(ev, [
      { title: "Copy cell value", icon: "copy", action: () => navigator.clipboard.writeText(item.text) },
      { title: "Copy cell reference", icon: "hash", action: () => navigator.clipboard.writeText(colToLetter(item.c) + item.r) },
    ]);
  }

  function itemForElement(el) {
    for (const item of itemEls) if (item.el === el) return item;
    return null;
  }

  function closestWithClass(node, className) {
    let cur = node;
    while (cur && cur.classList) {
      if (cur.classList.contains(className)) return cur;
      cur = cur.parentNode;
    }
    return null;
  }

  function scrollToCell(ref) {
    const range = parseRange(ref || "A1");
    if (!range) return;
    const el = resolveItemEls.get(key(range.r1, range.c1));
    if (!el) return;
    doScroll(el.offsetTop - 60, el.offsetLeft - 60);
    el.classList.add("xlsx-flash");
    setTimeout(() => el.classList.remove("xlsx-flash"), 900);
  }

  // scrollTo is missing in some DOM shims (and older webviews), so fall back
  // to setting scrollTop and scrollLeft.
  function doScroll(top, left) {
    if (!scrollEl) return;
    const t = Math.max(0, top);
    const l = Math.max(0, left);
    if (typeof scrollEl.scrollTo === "function") {
      scrollEl.scrollTo({ top: t, left: l });
    } else {
      scrollEl.scrollTop = t;
      scrollEl.scrollLeft = l;
    }
  }

  // ---------- filters ----------

  function filterColumns() {
    if (!G || !G.filterRange) return [];
    const out = [];
    const range = G.filterRange;
    for (let c = range.c1; c <= range.c2; c++) {
      const colId = c - range.c1;
      const def = sheet.autoFilter.columns.get(colId);
      if (def && def.showButton === false) continue;
      const values = new Map();
      let blanks = false;
      for (let r = range.r1 + 1; r <= range.r2; r++) {
        const text = displayTextRaw(r, c);
        if (text === "") blanks = true;
        values.set(text, (values.get(text) || 0) + 1);
      }
      const headerText = displayTextRaw(range.r1, c);
      out.push({
        colId,
        col: c,
        letter: colToLetter(c),
        headerText,
        values: Array.from(values.entries())
          .map(([value, count]) => ({ value, count }))
          .sort((a, b) => naturalCompare(a.value, b.value)),
        blanks,
        showButton: !(def && def.showButton === false),
      });
    }
    return out;
  }

  function openFilterPopup(colId, anchor) {
    closePopup();
    const range = G.filterRange;
    if (!range) return;
    const col = range.c1 + colId;
    const values = [];
    let blanks = false;
    for (let r = range.r1 + 1; r <= range.r2; r++) {
      const text = displayTextRaw(r, col);
      if (text === "") { blanks = true; continue; }
      if (values.indexOf(text) === -1) values.push(text);
    }
    values.sort(naturalCompare);
    const allValues = blanks ? values.concat([""]) : values;

    let current = filterState.get(colId);
    if (!(current instanceof Set)) current = current && current.custom ? null : current;

    const popup = doc.createElement("div");
    popup.className = "xlsx-filter-popup";

    const head = doc.createElement("div");
    head.className = "xlsx-filter-popup-head";
    const headerText = displayTextRaw(range.r1, col);
    head.textContent = "Column " + colToLetter(col) + (headerText ? " (" + headerText + ")" : "");
    popup.appendChild(head);

    const search = doc.createElement("input");
    search.type = "text";
    search.placeholder = "Search values";
    search.className = "xlsx-filter-search";
    popup.appendChild(search);

    const allRow = doc.createElement("label");
    allRow.className = "xlsx-filter-item xlsx-filter-all";
    const allBox = doc.createElement("input");
    allBox.type = "checkbox";
    allBox.checked = current == null;
    allRow.appendChild(allBox);
    allRow.appendChild(doc.createTextNode("Select all"));
    popup.appendChild(allRow);

    const list = doc.createElement("div");
    list.className = "xlsx-filter-list";
    popup.appendChild(list);

    const rows = new Map();
    const makeItem = (value, label) => {
      const row = doc.createElement("label");
      row.className = "xlsx-filter-item";
      const box = doc.createElement("input");
      box.type = "checkbox";
      box.checked = current == null || current.has(value);
      box.addEventListener("change", () => {
        const state = filterState.get(colId);
        let set;
        if (state instanceof Set) set = new Set(state);
        else if (state && state.custom) set = new Set(allValues);
        else set = new Set(allValues);
        if (box.checked) set.add(value);
        else set.delete(value);
        let next = set;
        if (set.size === allValues.length) next = null;
        filterState.set(colId, next);
        allBox.checked = next == null;
        callbacks.onFilterChange();
        render(true);
      });
      row.appendChild(box);
      row.appendChild(doc.createTextNode(label));
      rows.set(value, { row, box, label });
      list.appendChild(row);
      return row;
    };
    for (const value of values) makeItem(value, value);
    if (blanks) makeItem("", "(Blanks)");

    allBox.addEventListener("change", () => {
      if (allBox.checked) filterState.set(colId, null);
      else filterState.set(colId, new Set());
      for (const entry of rows.values()) entry.box.checked = allBox.checked;
      callbacks.onFilterChange();
      render(true);
    });

    search.addEventListener("input", () => {
      const q = search.value.toLowerCase();
      for (const entry of rows.values()) {
        const show = entry.label.toLowerCase().indexOf(q) !== -1;
        entry.row.style.display = show ? "" : "none";
      }
    });

    const footer = doc.createElement("div");
    footer.className = "xlsx-filter-popup-foot";

    const clearBtn = doc.createElement("button");
    clearBtn.className = "xlsx-filter-btn-plain";
    clearBtn.textContent = "Clear filter";
    clearBtn.addEventListener("click", () => {
      filterState.delete(colId);
      callbacks.onFilterChange();
      closePopup();
      render(true);
    });
    footer.appendChild(clearBtn);

    const closeBtn = doc.createElement("button");
    closeBtn.className = "xlsx-filter-btn-plain xlsx-filter-close";
    closeBtn.textContent = "Done";
    closeBtn.addEventListener("click", () => closePopup());
    footer.appendChild(closeBtn);
    popup.appendChild(footer);

    doc.body.appendChild(popup);
    const rect = anchor.getBoundingClientRect();
    const width = 230;
    let left = rect.right - width;
    if (left < 8) left = 8;
    if (left + width > doc.documentElement.clientWidth - 8) left = doc.documentElement.clientWidth - width - 8;
    popup.style.width = width + "px";
    popup.style.left = left + "px";
    const below = rect.bottom + 4;
    const maxHeight = Math.min(320, doc.documentElement.clientHeight - below - 16);
    popup.style.top = below + "px";
    popup.style.maxHeight = maxHeight + "px";
    popupEl = popup;

    const onDocDown = (ev) => {
      if (popup.contains(ev.target)) return;
      closePopup();
    };
    const onKey = (ev) => {
      if (ev.key === "Escape") closePopup();
    };
    popup._onDocDown = onDocDown;
    popup._onKey = onKey;
    doc.addEventListener("mousedown", onDocDown, true);
    doc.addEventListener("keydown", onKey, true);
  }

  function closePopup() {
    if (!popupEl) return;
    if (popupEl._onDocDown) doc.removeEventListener("mousedown", popupEl._onDocDown, true);
    if (popupEl._onKey) doc.removeEventListener("keydown", popupEl._onKey, true);
    if (popupEl.parentNode) popupEl.parentNode.removeChild(popupEl);
    popupEl = null;
  }

  // ---------- public API ----------

  function getState() {
    const filters = [];
    if (G && G.filterRange) {
      for (const colInfo of filterColumns()) {
        const sel = filterState.get(colInfo.colId);
        let selected = null;
        if (sel instanceof Set) selected = Array.from(sel);
        else if (sel && sel.custom) selected = "(custom)";
        filters.push({
          colId: colInfo.colId,
          letter: colInfo.letter,
          headerText: colInfo.headerText,
          values: colInfo.values,
          blanks: colInfo.blanks,
          totalCount: colInfo.values.length + (colInfo.blanks ? 1 : 0),
          selected,
          all: sel == null,
        });
      }
    }
    return {
      sheetName: sheet.name,
      totalRows: G ? G.maxRow : 0,
      visibleRows: G ? G.visRows.length : 0,
      totalCols: G ? G.maxCol : 0,
      visibleCols: G ? G.visCols.length : 0,
      rowCapped: G ? G.rowCapped : false,
      frozenRows: sheet.freeze ? sheet.freeze.rows : 0,
      frozenCols: sheet.freeze ? sheet.freeze.cols : 0,
      filters,
    };
  }

  function search(query) {
    const q = String(query || "").toLowerCase();
    let count = 0;
    for (const item of itemEls) {
      const hit = q !== "" && item.text.toLowerCase().indexOf(q) !== -1;
      if (hit) {
        item.el.classList.add("xlsx-search-hit");
        count++;
      } else {
        item.el.classList.remove("xlsx-search-hit");
        item.el.classList.remove("xlsx-search-current");
      }
    }
    searchIndex = -1;
    searchCount = count;
    return count;
  }

  let searchIndex = -1;
  let searchCount = 0;
  function searchNext(dir) {
    const hits = itemEls.filter((item) => item.el.classList.contains("xlsx-search-hit"));
    if (!hits.length) return null;
    searchIndex = (searchIndex + dir + hits.length * 2) % hits.length;
    for (const hit of hits) hit.el.classList.remove("xlsx-search-current");
    const item = hits[searchIndex];
    item.el.classList.add("xlsx-search-current");
    doScroll(item.el.offsetTop - 80, item.el.offsetLeft - 80);
    return { index: searchIndex + 1, count: hits.length };
  }

  function getScrollPosition() {
    return scrollEl ? { top: scrollEl.scrollTop, left: scrollEl.scrollLeft } : { top: 0, left: 0 };
  }

  function setScrollPosition(pos) {
    if (!scrollEl || !pos) return;
    scrollEl.scrollTop = pos.top || 0;
    scrollEl.scrollLeft = pos.left || 0;
  }

  function setSettings(patch) {
    Object.assign(settings, patch);
  }

  function destroy() {
    destroyed = true;
    closePopup();
    if (selectedEl) selectedEl.classList.remove("xlsx-selected");
    selectedEl = null;
    if (container) container.textContent = "";
    rootEl = scrollEl = gridEl = null;
  }

  render(false);

  return {
    render,
    destroy,
    getState,
    search,
    searchNext,
    scrollToCell,
    setSettings,
    getScrollPosition,
    setScrollPosition,
    getSelection,
    getFilterValues: filterColumns,
    setFilter(colId, selection) {
      filterState.set(colId, selection);
      render(true);
    },
    clearFilters() {
      filterState.clear();
      render(true);
    },
  };
}

module.exports = { createRenderer };
