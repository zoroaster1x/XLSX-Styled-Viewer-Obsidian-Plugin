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

// Colour handling: rgb/argb, theme colours with tint, legacy indexed palette.

// The classic 64 colour indexed palette used by older files.
const INDEXED_COLORS = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

function parseHex(hex) {
  let h = String(hex).replace("#", "").trim();
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  if (h.length === 6) h = "FF" + h;
  if (h.length !== 8) return null;
  const n = parseInt(h, 16);
  if (isNaN(n)) return null;
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: (n >> 24) & 255 };
}

function rgbToHex(r, g, b) {
  const to = (v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0");
  return "#" + to(r) + to(g) + to(b);
}

function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return { h: h / 6, s, l };
}

function hueToRgb(p, q, t) {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
}

function hslToRgb(h, s, l) {
  if (s === 0) {
    const v = l * 255;
    return { r: v, g: v, b: v };
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return {
    r: hueToRgb(p, q, h + 1 / 3) * 255,
    g: hueToRgb(p, q, h) * 255,
    b: hueToRgb(p, q, h - 1 / 3) * 255,
  };
}

// Excel tint: positive tint moves the colour toward white, negative toward black.
function applyTint(hex, tint) {
  if (!tint) return hex;
  const rgb = parseHex(hex);
  if (!rgb) return hex;
  const hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);
  let l = hsl.l;
  l = tint < 0 ? l * (1 + tint) : l + (1 - l) * tint;
  const out = hslToRgb(hsl.h, hsl.s, Math.max(0, Math.min(1, l)));
  return rgbToHex(out.r, out.g, out.b);
}

// colour is { rgb, theme, tint, indexed, auto }. Returns css colour or fallback.
function resolveColor(color, theme, fallback) {
  const fb = fallback === undefined ? null : fallback;
  if (!color) return fb;
  if (color.auto) return fb;
  if (color.rgb) {
    const c = parseHex(color.rgb);
    if (!c) return fb;
    if (c.a >= 255) return rgbToHex(c.r, c.g, c.b);
    return "rgba(" + c.r + ", " + c.g + ", " + c.b + ", " + (c.a / 255).toFixed(3) + ")";
  }
  if (color.theme != null && theme && theme.colors && theme.colors[color.theme]) {
    let hex = theme.colors[color.theme];
    if (hex && hex[0] !== "#") hex = "#" + hex;
    if (color.tint) hex = applyTint(hex, color.tint);
    return hex;
  }
  if (color.indexed != null) {
    if (color.indexed === 64) return "#000000";
    if (color.indexed === 65) return "#FFFFFF";
    const hex = INDEXED_COLORS[color.indexed];
    if (!hex) return fb;
    let out = hex;
    if (color.tint) out = applyTint(out, color.tint);
    return rgbToHex(parseHex(out).r, parseHex(out).g, parseHex(out).b);
  }
  return fb;
}

// True when the colour means "automatic" (no explicit colour given).
function isAutoColor(color) {
  if (!color) return true;
  if (color.auto) return true;
  if (color.theme === 1 && !color.tint) return true;
  return false;
}

module.exports = {
  INDEXED_COLORS,
  parseHex,
  rgbToHex,
  applyTint,
  resolveColor,
  isAutoColor,
};
