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

const { Plugin, PluginSettingTab, Setting } = require("obsidian");
const { XlsxView, VIEW_TYPE } = require("./view");

const DEFAULT_SETTINGS = {
  sheetBackground: "white",
  showGridlines: true,
  showHeaders: true,
  zoom: 100,
  maxRows: 2000,
  chipLimit: 10,
  fileState: {},
};

class XlsxStyledViewerPlugin extends Plugin {
  async onload() {
    await this.loadSettings();

    this.registerView(VIEW_TYPE, (leaf) => new XlsxView(leaf, this));
    this.registerExtensions(["xlsx", "xlsm"], VIEW_TYPE);
    this.addSettingTab(new XlsxSettingTab(this.app, this));

    this.addCommand({
      id: "reload-workbook",
      name: "Reload the current workbook",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(XlsxView);
        if (!view) return false;
        if (!checking) view.reload();
        return true;
      },
    });
  }

  async loadSettings() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }

  async updateSetting(key, value) {
    this.settings[key] = value;
    await this.saveSettings();
    for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE)) {
      const view = leaf.view;
      if (view instanceof XlsxView && typeof view.applySettings === "function") {
        view.applySettings();
      }
    }
  }

  // Per file view state (filters, filter bar, active sheet). Saved quietly so
  // other open views are not re-rendered on every chip click.
  async saveFileState(path, state) {
    if (!this.settings.fileState) this.settings.fileState = {};
    this.settings.fileState[path] = state;
    await this.saveSettings();
  }
}

class XlsxSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", { text: "XLSX Styled Viewer" });

    new Setting(containerEl)
      .setName("Sheet background")
      .setDesc("White keeps the sheet looking like Excel even in a dark theme. Theme follows the Obsidian colours for cells that have no fill of their own.")
      .addDropdown((dropdown) =>
        dropdown
          .addOption("white", "Always white (recommended)")
          .addOption("theme", "Follow Obsidian theme")
          .setValue(this.plugin.settings.sheetBackground)
          .onChange((value) => this.plugin.updateSetting("sheetBackground", value))
      );

    new Setting(containerEl)
      .setName("Show gridlines")
      .setDesc("Draw light gridlines in cells that have no border of their own.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showGridlines !== false)
          .onChange((value) => this.plugin.updateSetting("showGridlines", value))
      );

    new Setting(containerEl)
      .setName("Show row and column headers")
      .setDesc("Show the A B C column letters and the row numbers around the sheet.")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.showHeaders !== false)
          .onChange((value) => this.plugin.updateSetting("showHeaders", value))
      );

    new Setting(containerEl)
      .setName("Default zoom")
      .setDesc("Starting zoom for new views. You can also change it from the toolbar.")
      .addSlider((slider) =>
        slider
          .setLimits(40, 300, 10)
          .setValue(Number(this.plugin.settings.zoom) || 100)
          .setDynamicTooltip()
          .onChange((value) => this.plugin.updateSetting("zoom", value))
      );

    new Setting(containerEl)
      .setName("Row render limit")
      .setDesc("Maximum number of rows to draw at once, to keep very large sheets responsive. Set 0 for no limit.")
      .addText((text) =>
        text
          .setPlaceholder("2000")
          .setValue(String(this.plugin.settings.maxRows))
          .onChange((value) => {
            const n = parseInt(value, 10);
            this.plugin.updateSetting("maxRows", isNaN(n) || n < 0 ? 0 : n);
          })
      );

    new Setting(containerEl)
      .setName("Filter chips limit")
      .setDesc("Show filter chips above the grid for columns with at most this many values. The full filter dropdown is still available on every filterable column.")
      .addSlider((slider) =>
        slider
          .setLimits(4, 40, 2)
          .setValue(Number(this.plugin.settings.chipLimit) || 10)
          .setDynamicTooltip()
          .onChange((value) => this.plugin.updateSetting("chipLimit", value))
      );

    containerEl.createEl("p", {
      cls: "setting-item-description",
      text: "Tip: for sheets with an autofilter, the filter chips above the grid let you hide whole groups (for example group A or group B) with one click.",
    });
  }
}

module.exports = XlsxStyledViewerPlugin;
