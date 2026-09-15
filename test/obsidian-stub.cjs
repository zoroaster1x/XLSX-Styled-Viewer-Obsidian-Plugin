"use strict";

// Minimal Obsidian API stand-in used by test/smoke.mjs. Enough to build the
// plugin, register the view and construct the view chrome without a DOM.

function makeEl(tag) {
  const el = {
    tagName: (tag || "div").toUpperCase(),
    children: [],
    classes: new Set(),
    attrs: {},
    style: {},
    value: "",
    textContent: "",
    empty() { this.children.length = 0; this.textContent = ""; return this; },
    addClass(cls) { if (cls) this.classes.add(cls); return this; },
    removeClass(cls) { this.classes.delete(cls); return this; },
    toggleClass(cls, on) { if (on === undefined) on = !this.classes.has(cls); if (on) this.classes.add(cls); else this.classes.delete(cls); return this; },
    setText(text) { this.textContent = String(text); return this; },
    setAttribute(name, value) { this.attrs[name] = value; return this; },
    getAttribute(name) { return Object.prototype.hasOwnProperty.call(this.attrs, name) ? this.attrs[name] : null; },
    createEl(t, info, cb) {
      const child = makeEl(t);
      if (info && info.cls) child.classes.add(info.cls);
      if (info && info.attr) Object.assign(child.attrs, info.attr);
      this.children.push(child);
      if (cb) cb(child);
      return child;
    },
    createDiv(info, cb) { return this.createEl("div", typeof info === "string" ? { cls: info } : info, typeof info === "function" ? info : cb); },
    createSpan(info, cb) { return this.createEl("span", typeof info === "string" ? { cls: info } : info, typeof info === "function" ? info : cb); },
    insertBefore(child) { this.children.unshift(child); return child; },
    appendChild(child) { this.children.push(child); return child; },
    removeChild(child) { const i = this.children.indexOf(child); if (i !== -1) this.children.splice(i, 1); return child; },
    remove() {},
    querySelectorAll() { return []; },
    querySelector() { return null; },
    addEventListener() {},
    removeEventListener() {},
    getBoundingClientRect() { return { top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }; },
    focus() {},
    show() {},
    hide() {},
    setAttr(name, value) { return this.setAttribute(name, value); },
  };
  return el;
}

class Plugin {
  constructor(app, manifest) {
    this.app = app;
    this.manifest = manifest || {};
    this.registered = { views: [], extensions: [], commands: [], settingTabs: [] };
  }
  async loadData() { return null; }
  async saveData() {}
  async saveFileState() {}
  registerView(type, factory) { this.registered.views.push({ type, factory }); }
  registerExtensions(exts, type) { this.registered.extensions.push({ exts, type }); }
  addSettingTab(tab) { this.registered.settingTabs.push(tab); }
  addCommand(cmd) { this.registered.commands.push(cmd); }
  addRibbonIcon() { return makeEl("div"); }
  async onload() {}
}

class ItemView {
  constructor(leaf) {
    this.leaf = leaf;
    this.app = leaf && leaf.app ? leaf.app : {};
    this.containerEl = makeEl("div");
    this.contentEl = makeEl("div");
  }
  getViewType() { return "x"; }
  getDisplayText() { return "x"; }
  async onOpen() {}
  async onClose() {}
}

// FileView is what Obsidian uses for file-backed views. It owns this.file and
// calls onLoadFile and onUnloadFile when the leaf opens a file.
class FileView extends ItemView {
  constructor(leaf) {
    super(leaf);
    this.file = null;
    this.allowNoFile = false;
  }
  async onLoadFile() {}
  async onUnloadFile() {}
  async onRename() {}
  getState() { return {}; }
  setState() {}
}

class Setting {
  constructor(container) { this.containerEl = container; }
  setName() { return this; }
  setDesc() { return this; }
  setClass() { return this; }
  addDropdown() { return this; }
  addToggle() { return this; }
  addSlider() { return this; }
  addText() { return this; }
  addButton() { return this; }
  addExtraButton() { return this; }
}

class PluginSettingTab {
  constructor(app, plugin) {
    this.app = app;
    this.plugin = plugin;
    this.containerEl = makeEl("div");
  }
  display() {}
}

class Menu {
  addItem(cb) { cb({ setTitle() { return this; }, setIcon() { return this; }, onClick() { return this; } }); return this; }
  showAtMouseEvent() {}
}

class Notice {
  constructor(message) { this.message = message; }
}

module.exports = {
  Plugin,
  ItemView,
  FileView,
  Setting,
  PluginSettingTab,
  Menu,
  Notice,
  setIcon: () => {},
  normalizePath: (p) => p,
};
