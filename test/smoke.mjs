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

// Smoke test for the built bundle: loads main.js with a stubbed Obsidian API,
// runs onload, and constructs the view chrome. Run with: bun test/smoke.mjs

import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const require = createRequire(import.meta.url);

let pass = 0;
let fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log("ok   " + name); }
  else { fail++; console.log("FAIL " + name + (extra !== undefined ? "  -> " + extra : "")); }
}

const bundle = readFileSync(join(root, "main.js"), "utf8");
const stubPath = join(here, "obsidian-stub.cjs");
const patched = bundle.split('require("obsidian")').join("require(" + JSON.stringify(stubPath) + ")");
const tmp = mkdtempSync(join(tmpdir(), "xlsx-smoke-"));
const bundlePath = join(tmp, "main.cjs");
writeFileSync(bundlePath, patched);

const PluginClass = require(bundlePath);
check("bundle exports a plugin class", typeof PluginClass === "function");

const app = {
  vault: { readBinary: async () => new Uint8Array() },
  workspace: {
    getLeavesOfType: () => [],
    getActiveViewOfType: () => null,
    on: () => {},
  },
};

const plugin = new PluginClass(app, { id: "xlsx-styled-viewer" });
await plugin.onload();

check("plugin registered the view type", plugin.registered.views.length === 1
  && plugin.registered.views[0].type === "xlsx-styled-viewer",
  JSON.stringify(plugin.registered.views.map((v) => v.type)));
check("plugin registered xlsx and xlsm", plugin.registered.extensions.length === 1
  && plugin.registered.extensions[0].exts.indexOf("xlsx") !== -1
  && plugin.registered.extensions[0].exts.indexOf("xlsm") !== -1,
  JSON.stringify(plugin.registered.extensions));
check("plugin added a settings tab", plugin.registered.settingTabs.length === 1);
check("plugin added a reload command", plugin.registered.commands.length === 1
  && plugin.registered.commands[0].id === "reload-workbook");

const leaf = { app };
const view = plugin.registered.views[0].factory(leaf);
check("view reports the right type", view.getViewType() === "xlsx-styled-viewer");
check("view has a display name", typeof view.getDisplayText() === "string");
await view.onOpen();
check("view chrome built", view.toolbarEl && view.tabsEl && view.filterBarEl && view.gridHostEl && view.statusEl);
const tabEl = view.buildTabsForTest ? null : null;
void tabEl;
view.applySettings();
check("applySettings is safe without a renderer", true);
await view.onClose();
check("view closes cleanly", true);

console.log("");
console.log(pass + " passed, " + fail + " failed");
process.exit(fail === 0 ? 0 : 1);
