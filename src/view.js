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

const { FileView, Menu, setIcon } = require("obsidian");
const { readWorkbook } = require("./read");
const { createRenderer } = require("./render");

const VIEW_TYPE = "xlsx-styled-viewer";
const DEFAULT_CHIP_LIMIT = 10;

function serializeFilters(map) {
  const out = {};
  for (const [colId, sel] of map) {
    if (sel instanceof Set) out[colId] = { values: Array.from(sel) };
    else if (sel && sel.custom) out[colId] = { custom: sel.custom };
    else out[colId] = { all: true };
  }
  return out;
}

function deserializeFilters(data) {
  const map = new Map();
  if (!data) return map;
  for (const key of Object.keys(data)) {
    const colId = Number(key);
    const entry = data[key];
    if (!entry || entry.all) map.set(colId, null);
    else if (entry.values) map.set(colId, new Set(entry.values));
    else if (entry.custom) map.set(colId, { custom: entry.custom });
  }
  return map;
}

class XlsxView extends FileView {
  constructor(leaf, plugin) {
    super(leaf);
    this.plugin = plugin;
    this.file = null;
    this.book = null;
    this.sheetIndex = 0;
    this.filterStates = new Map();
    this.outlineStates = new Map();
    this.scrollPositions = new Map();
    this.renderer = null;
    this.error = null;
    this.built = false;
    this.showFilterBar = true;
    this.savedFilters = null;
    this.selection = null;
  }

  getViewType() {
    return VIEW_TYPE;
  }

  getDisplayText() {
    return this.file ? this.file.basename : "XLSX";
  }

  getIcon() {
    return "table-2";
  }

  async onOpen() {
    this.buildChrome();
  }

  async onLoadFile(file) {
    this.buildChrome();
    await this.renderFile(file);
  }

  async onUnloadFile() {
    this.teardownRenderer();
    this.book = null;
    this.filterStates = new Map();
    this.outlineStates = new Map();
    this.scrollPositions = new Map();
    if (this.gridHostEl) {
      this.gridHostEl.textContent = "";
      this.gridHostEl.createDiv("xlsx-loading", (el) => el.setText("No file loaded."));
    }
    if (this.statusEl) this.statusEl.setText("");
  }

  async onClose() {
    this.teardownRenderer();
  }

  buildChrome() {
    if (this.built) return;
    this.built = true;
    const content = this.contentEl;
    content.empty();
    content.addClass("xlsx-view");

    this.toolbarEl = content.createDiv("xlsx-toolbar");
    this.titleEl = this.toolbarEl.createDiv("xlsx-title");
    this.titleEl.setText("XLSX viewer");
    const tools = this.toolbarEl.createDiv("xlsx-tools");

    const makeButton = (icon, title, onClick) => {
      const btn = tools.createEl("button", { cls: "xlsx-tool-btn", attr: { "aria-label": title, title } });
      setIcon(btn, icon);
      btn.addEventListener("click", onClick);
      return btn;
    };

    makeButton("chevron-up", "Previous match", () => this.stepSearch(-1));
    makeButton("chevron-down", "Next match", () => this.stepSearch(1));
    this.searchCountEl = tools.createSpan("xlsx-search-count");
    this.searchInput = tools.createEl("input", {
      cls: "xlsx-search-input",
      attr: { type: "text", placeholder: "Search", spellcheck: "false" },
    });
    this.searchInput.addEventListener("input", () => this.runSearch(this.searchInput.value));
    this.searchInput.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault();
        this.stepSearch(ev.shiftKey ? -1 : 1);
      }
    });

    this.zoomLabel = tools.createSpan("xlsx-zoom-label");
    makeButton("zoom-out", "Zoom out", () => this.setZoom(this.currentZoom() - 10));
    makeButton("zoom-in", "Zoom in", () => this.setZoom(this.currentZoom() + 10));

    this.bgBtn = makeButton("contrast", "Toggle white or theme background", async () => {
      const next = this.plugin.settings.sheetBackground === "white" ? "theme" : "white";
      await this.plugin.updateSetting("sheetBackground", next);
      this.applySettings();
    });
    this.filterBtn = makeButton("filter", "Show or hide the filter chips", () => {
      this.showFilterBar = !this.showFilterBar;
      this.updateFilterBarVisibility();
      this.saveFileState();
    });
    makeButton("list-x", "Clear all filters", () => {
      if (this.renderer) {
        this.renderer.clearFilters();
        this.refreshFilterUi();
        this.saveFileState();
      }
    });
    makeButton("refresh-cw", "Reload the file", () => this.reload());

    this.tabsEl = content.createDiv("xlsx-tabs");
    this.gridHostEl = content.createDiv("xlsx-grid-host");
    this.filterBarEl = content.createDiv("xlsx-filterbar");
    this.statusEl = content.createDiv("xlsx-status");
  }

  currentZoom() {
    const z = this.plugin.settings.zoom;
    return typeof z === "number" && z >= 40 && z <= 400 ? z : 100;
  }

  async setZoom(value) {
    const zoom = Math.max(40, Math.min(400, Math.round(value / 10) * 10));
    await this.plugin.updateSetting("zoom", zoom);
    this.applySettings();
  }

  applySettings() {
    if (!this.renderer) return;
    this.renderer.setSettings(this.rendererSettings());
    this.renderer.render(true);
    this.updateStatus();
    this.bgBtn.toggleClass("is-active", this.plugin.settings.sheetBackground === "white");
    this.zoomLabel.setText(this.currentZoom() + "%");
  }

  rendererSettings() {
    return {
      sheetBackground: this.plugin.settings.sheetBackground === "theme" ? "theme" : "white",
      showGridlines: this.plugin.settings.showGridlines !== false,
      showHeaders: this.plugin.settings.showHeaders !== false,
      maxRows: Number(this.plugin.settings.maxRows) > 0 ? Number(this.plugin.settings.maxRows) : 0,
      scale: this.currentZoom() / 100,
    };
  }

  // Called from onLoadFile and from the reload button. Never throws: file
  // problems are reported in the view so FileView does not close the leaf.
  async renderFile(file) {
    this.buildChrome();
    this.error = null;
    this.titleEl.setText(file.name);
    this.renderer && this.renderer.destroy();
    this.renderer = null;
    this.gridHostEl.empty();
    this.gridHostEl.createDiv("xlsx-loading", (el) => el.setText("Reading " + file.name + " ..."));
    this.statusEl.setText("");
    try {
      const buffer = await this.app.vault.readBinary(file);
      this.book = readWorkbook(buffer);
      this.filterStates = new Map();
      this.outlineStates = new Map();
      this.scrollPositions = new Map();
      const saved = (this.plugin.settings.fileState || {})[file.path];
      this.showFilterBar = saved ? saved.showFilterBar !== false : true;
      this.savedFilters = saved && saved.filters ? saved.filters : null;
      this.sheetIndex = this.firstVisibleSheet();
      if (saved && Number.isInteger(saved.sheetIndex) && this.book.sheets[saved.sheetIndex]
        && this.book.sheets[saved.sheetIndex].state === "visible") {
        this.sheetIndex = saved.sheetIndex;
      }
      this.buildTabs();
      this.renderSheet();
    } catch (err) {
      this.book = null;
      this.showError(err);
    }
  }

  saveFileState() {
    if (!this.file || !this.book) return;
    const path = this.file.path;
    const sheetName = this.book.sheets[this.sheetIndex] ? this.book.sheets[this.sheetIndex].name : String(this.sheetIndex);
    const existing = ((this.plugin.settings.fileState || {})[path]) || {};
    const filters = Object.assign({}, existing.filters || {});
    filters[sheetName] = serializeFilters(this.filterStateFor(this.sheetIndex));
    this.plugin.saveFileState(path, {
      showFilterBar: this.showFilterBar !== false,
      filters,
      sheetIndex: this.sheetIndex,
    });
  }

  firstVisibleSheet() {
    if (!this.book) return 0;
    for (let i = 0; i < this.book.sheets.length; i++) {
      if (this.book.sheets[i].state === "visible") return i;
    }
    return 0;
  }

  visibleSheets() {
    if (!this.book) return [];
    const out = [];
    for (let i = 0; i < this.book.sheets.length; i++) {
      if (this.book.sheets[i].state === "visible") out.push(this.book.sheets[i]);
    }
    return out;
  }

  buildTabs() {
    this.tabsEl.empty();
    const sheets = this.visibleSheets();
    this.tabsEl.toggleClass("is-hidden", sheets.length <= 1);
    for (const sheet of sheets) {
      const tab = this.tabsEl.createEl("button", { cls: "xlsx-tab" });
      tab.setText(sheet.name);
      tab.toggleClass("is-active", sheet.index === this.sheetIndex);
      tab.addEventListener("click", () => this.switchSheet(sheet.index));
    }
  }

  updateTabs() {
    const tabs = this.tabsEl.querySelectorAll(".xlsx-tab");
    const sheets = this.visibleSheets();
    for (let i = 0; i < tabs.length && i < sheets.length; i++) {
      tabs[i].toggleClass("is-active", sheets[i].index === this.sheetIndex);
    }
  }

  switchSheet(index, ref) {
    if (!this.book || index === this.sheetIndex) {
      if (ref && this.renderer) this.renderer.scrollToCell(ref);
      return;
    }
    if (this.renderer) this.scrollPositions.set(this.sheetIndex, this.renderer.getScrollPosition());
    this.saveFileState();
    this.sheetIndex = index;
    this.renderSheet();
    if (ref && this.renderer) this.renderer.scrollToCell(ref);
    const saved = this.scrollPositions.get(index);
    if (!ref && saved && this.renderer) this.renderer.setScrollPosition(saved);
    this.saveFileState();
  }

  filterStateFor(index) {
    if (!this.filterStates.has(index)) this.filterStates.set(index, new Map());
    return this.filterStates.get(index);
  }

  outlineStateFor(index) {
    if (!this.outlineStates.has(index)) this.outlineStates.set(index, new Map());
    return this.outlineStates.get(index);
  }

  renderSheet() {
    this.teardownRenderer();
    const model = this.book.loadSheet(this.sheetIndex);
    const filterState = this.filterStateFor(this.sheetIndex);
    const savedForSheet = this.savedFilters ? this.savedFilters[model.name] : null;
    if (savedForSheet) {
      filterState.clear();
      for (const [colId, sel] of deserializeFilters(savedForSheet)) filterState.set(colId, sel);
    } else if (filterState.size === 0 && model.autoFilter) {
      for (const [colId, def] of model.autoFilter.columns) {
        if (def.values) {
          const set = new Set(def.values);
          if (def.blank) set.add("");
          filterState.set(colId, set.size ? set : null);
        } else if (def.custom) {
          filterState.set(colId, { custom: def.custom });
        }
      }
    }
    this.renderer = createRenderer({
      container: this.gridHostEl,
      sheet: model,
      styles: this.book.styles,
      theme: this.book.theme,
      date1904: this.book.date1904,
      settings: this.rendererSettings(),
      filterState,
      outlineCollapsed: this.outlineStateFor(this.sheetIndex),
      onNavigate: (sheetName, ref) => {
        const target = this.book.sheetIndexByName(sheetName);
        if (target === -1) return;
        this.switchSheet(target, ref);
      },
      onOpenExternal: (url) => this.openExternal(url),
      showContextMenu: (ev, items) => this.showCellMenu(ev, items),
      onSelect: (sel) => {
        this.selection = sel;
        this.updateStatus();
      },
      onFilterChange: () => {
        this.refreshFilterUi();
        this.updateStatus();
        this.saveFileState();
      },
    });
    this.refreshFilterUi();
    this.updateTabs();
    this.updateStatus();
  }

  updateFilterBarVisibility() {
    this.filterBarEl.toggleClass("is-hidden", this.showFilterBar === false || this.filterBarEl.childElementCount === 0);
    if (this.filterBtn) this.filterBtn.toggleClass("is-active", this.showFilterBar !== false);
  }

  openExternal(url) {
    try {
      const electron = require("electron");
      if (electron && electron.shell && electron.shell.openExternal) {
        electron.shell.openExternal(url);
        return;
      }
    } catch (err) {
      // Not on desktop; fall back to window.open below.
    }
    window.open(url, "_blank");
  }

  showCellMenu(ev, items) {
    const menu = new Menu();
    for (const item of items) {
      menu.addItem((mi) => mi.setTitle(item.title).setIcon(item.icon).onClick(item.action));
    }
    menu.showAtMouseEvent(ev);
  }

  refreshFilterUi() {
    this.filterBarEl.empty();
    if (!this.renderer) return;
    const state = this.renderer.getState();
    const limit = Number(this.plugin.settings.chipLimit) > 0 ? Number(this.plugin.settings.chipLimit) : DEFAULT_CHIP_LIMIT;
    let groups = 0;
    for (const filter of state.filters) {
      const valueCount = filter.values.length + (filter.blanks ? 1 : 0);
      if (valueCount > limit) continue;
      // Columns with nothing but blanks have nothing worth filtering.
      if (filter.values.length < 2) continue;
      groups++;
      const group = this.filterBarEl.createDiv("xlsx-chip-group");
      const label = group.createSpan("xlsx-chip-label");
      label.setText(filter.letter + (filter.headerText ? ": " + filter.headerText : ""));

      const selected = filter.selected instanceof Array ? new Set(filter.selected) : null;
      const isAll = filter.all;
      const allValues = filter.values.map((v) => v.value);
      if (filter.blanks) allValues.push("");
      const toggleValue = (value) => {
        let set;
        if (selected) set = new Set(selected);
        else set = new Set(allValues);
        if (set.has(value)) set.delete(value);
        else set.add(value);
        const next = set.size === allValues.length ? null : set;
        this.renderer.setFilter(filter.colId, next);
        this.refreshFilterUi();
        this.updateStatus();
        this.saveFileState();
      };
      for (const value of filter.values) {
        const chip = group.createEl("button", { cls: "xlsx-chip" });
        chip.setText(value.value);
        chip.toggleClass("is-off", !(isAll || (selected && selected.has(value.value))));
        chip.addEventListener("click", () => toggleValue(value.value));
      }
      if (filter.blanks) {
        const chip = group.createEl("button", { cls: "xlsx-chip" });
        chip.setText("(Blanks)");
        chip.toggleClass("is-off", !(isAll || (selected && selected.has(""))));
        chip.addEventListener("click", () => toggleValue(""));
      }
      const allBtn = group.createEl("button", { cls: "xlsx-chip xlsx-chip-all" });
      allBtn.setText("all");
      allBtn.addEventListener("click", () => {
        this.renderer.setFilter(filter.colId, null);
        this.refreshFilterUi();
        this.updateStatus();
        this.saveFileState();
      });
      const noneBtn = group.createEl("button", { cls: "xlsx-chip xlsx-chip-all" });
      noneBtn.setText("none");
      noneBtn.addEventListener("click", () => {
        this.renderer.setFilter(filter.colId, new Set());
        this.refreshFilterUi();
        this.updateStatus();
        this.saveFileState();
      });
    }
    this.updateFilterBarVisibility();
  }

  updateStatus() {
    if (!this.renderer) return;
    const state = this.renderer.getState();
    const parts = [];
    parts.push("Sheet: " + state.sheetName);
    parts.push(state.totalRows + " rows x " + state.totalCols + " cols");
    if (state.visibleRows !== state.totalRows || state.visibleCols !== state.totalCols) {
      parts.push(state.visibleRows + " x " + state.visibleCols + " visible");
    }
    const active = state.filters.filter((f) => !f.all);
    if (active.length) parts.push(active.length + " filter" + (active.length === 1 ? "" : "s") + " active");
    if (this.selection && this.selection.ref) {
      parts.push("Selected " + this.selection.ref + " (" + this.selection.text.replace(/\s+/g, " ").slice(0, 60) + ")");
    }
    this.statusEl.setText(parts.join("  |  "));
    this.zoomLabel.setText(this.currentZoom() + "%");
  }

  runSearch(query) {
    if (!this.renderer) return;
    const count = this.renderer.search(query);
    this.searchCountEl.setText(query ? count + (count === 1 ? " match" : " matches") : "");
  }

  stepSearch(dir) {
    if (!this.renderer) return;
    const result = this.renderer.searchNext(dir);
    if (result) this.searchCountEl.setText(result.index + " of " + result.count);
  }

  showError(err) {
    this.gridHostEl.empty();
    const box = this.gridHostEl.createDiv("xlsx-error");
    box.createEl("h3", { text: "Could not read this workbook" });
    box.createEl("p", { text: err && err.message ? err.message : String(err) });
    box.createEl("p", {
      cls: "xlsx-error-hint",
      text: "The viewer reads .xlsx and .xlsm files. Old .xls files are a different format and are not supported.",
    });
    this.statusEl.setText("Error");
  }

  teardownRenderer() {
    if (this.renderer) {
      this.renderer.destroy();
      this.renderer = null;
    }
  }

  async reload() {
    if (this.file) await this.renderFile(this.file);
  }
}

module.exports = { XlsxView, VIEW_TYPE };
