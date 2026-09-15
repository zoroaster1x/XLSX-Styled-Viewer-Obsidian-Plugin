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

// Small helpers shared by the xlsx parser and the renderer.

function colToLetter(col) {
  let s = "";
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function letterToCol(letters) {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    const code = letters.charCodeAt(i);
    if (code < 65 || code > 90) continue;
    col = col * 26 + (code - 64);
  }
  return col;
}

function parseCellRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(ref).trim());
  if (!m) return null;
  return { c: letterToCol(m[1].toUpperCase()), r: parseInt(m[2], 10) };
}

function parseRange(ref) {
  const parts = String(ref).split(":");
  const a = parseCellRef(parts[0]);
  const b = parts.length > 1 ? parseCellRef(parts[1]) : a;
  if (!a || !b) return null;
  return {
    r1: Math.min(a.r, b.r),
    c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r),
    c2: Math.max(a.c, b.c),
  };
}

// Splits "Sheet name!A1:B2" into { sheet, range }. Handles quoted names.
function splitSheetRef(ref) {
  const bang = String(ref).lastIndexOf("!");
  if (bang === -1) return { sheet: null, ref: String(ref) };
  let sheet = String(ref).slice(0, bang).trim();
  if (sheet.startsWith("'") && sheet.endsWith("'")) {
    sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }
  return { sheet, ref: String(ref).slice(bang + 1) };
}

// Excel column width (in characters) to pixels, using the usual 7px digit width.
function pxFromCharWidth(chars) {
  return Math.round(chars * 7) + 5;
}

// Excel row height and font sizes are stored in points.
function pxFromPt(pt) {
  return (pt * 96) / 72;
}

function clamp(v, lo, hi) {
  return v < lo ? lo : v > hi ? hi : v;
}

function key(r, c) {
  return r + ":" + c;
}

// Numbers first, then text, both in a human order.
function naturalCompare(a, b) {
  const sa = String(a == null ? "" : a);
  const sb = String(b == null ? "" : b);
  const na = sa.trim() === "" ? NaN : Number(sa);
  const nb = sb.trim() === "" ? NaN : Number(sb);
  const aNum = !isNaN(na);
  const bNum = !isNaN(nb);
  if (aNum && bNum) return na - nb;
  if (aNum) return -1;
  if (bNum) return 1;
  return sa.localeCompare(sb, undefined, { numeric: true, sensitivity: "base" });
}

// ---------- XML helpers ----------

function parseXml(text) {
  return new DOMParser().parseFromString(text, "application/xml");
}

function tagName(el) {
  const t = el.tagName || el.nodeName || "";
  const i = t.indexOf(":");
  return i === -1 ? t : t.slice(i + 1);
}

function childrenOf(el, tag) {
  const out = [];
  const list = el.children || [];
  for (let i = 0; i < list.length; i++) {
    if (tagName(list[i]) === tag) out.push(list[i]);
  }
  return out;
}

function firstOf(el, tag) {
  const list = el.children || [];
  for (let i = 0; i < list.length; i++) {
    if (tagName(list[i]) === tag) return list[i];
  }
  return null;
}

function attr(el, name) {
  const v = el.getAttribute(name);
  return v === null ? null : v;
}

function attrInt(el, name, dflt) {
  const v = el.getAttribute(name);
  if (v === null || v === "") return dflt === undefined ? 0 : dflt;
  const n = parseInt(v, 10);
  return isNaN(n) ? (dflt === undefined ? 0 : dflt) : n;
}

function attrBool(el, name, dflt) {
  const v = el.getAttribute(name);
  if (v === null) return dflt === undefined ? false : dflt;
  return v === "1" || v === "true";
}

// Joins text runs of a shared string or inline string. Skips phonetic runs.
function textOfRuns(el) {
  let out = "";
  const walk = (node) => {
    const list = node.children || [];
    for (let i = 0; i < list.length; i++) {
      const child = list[i];
      const tag = tagName(child);
      if (tag === "rPh" || tag === "phoneticPr") continue;
      if (tag === "t") {
        out += child.textContent || "";
      } else {
        walk(child);
      }
    }
  };
  walk(el);
  return out;
}

module.exports = {
  colToLetter,
  letterToCol,
  parseCellRef,
  parseRange,
  splitSheetRef,
  pxFromCharWidth,
  pxFromPt,
  clamp,
  key,
  naturalCompare,
  parseXml,
  tagName,
  childrenOf,
  firstOf,
  attr,
  attrInt,
  attrBool,
  textOfRuns,
};
