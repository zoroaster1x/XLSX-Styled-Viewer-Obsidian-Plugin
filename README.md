>Replaced by https://github.com/zoroaster1x/Styled-MS-Office-Viewer-for-Obsidian

# XLSX-Styled-Viewer-Obsidian-Plugin

An Obsidian plugin that opens .xlsx workbooks as styled spreadsheets instead of plain value tables. Fills, fonts, borders, merged cells, frozen panes, wrapped text, row and column sizes and autofilters all render the way they do in Excel.

## Why this plugin exists

I keep spreadsheets in my vault. Plain value tables turn a sheet with pastel session blocks, merged cells and a group column into an unreadable wall of text. I wanted something closer to opening the file in Excel or LibreOffice: same colours, same cell sizes, same wrapped text, plus a quick way to hide the groups I am not part of.

## Features

- Full cell styling: solid and gradient fills, themed colours with tints, font name, size, bold, italic, underline, strike, font colour, all border styles and weights.
- Excel geometry: column widths and row heights in real units, hidden rows and columns, frozen rows and columns (sticky while you scroll).
- Merged cells, including borders that follow the merged region instead of the individual cells.
- Text wrapping, vertical and horizontal alignment, indentation and text overflow into empty neighbours, the way Excel draws it.
- Number formats: General, decimals, thousands separators, percent, scientific, text, colours and conditions, plus date and time codes including elapsed time.
- Autofilter dropdowns on the header row, with value search, Select all, Clear filter and per value checkboxes.
- Filter chips live at the bottom of the view and can be hidden with the filter button in the toolbar.
- Every file remembers its filters, hidden filter bar and last active sheet, so reopening a workbook shows the same groups you had turned off.
- Click any cell to select it; the ring shows which cell the status bar is describing.
- Filter chips above the grid for columns with a short list of values, so you can hide whole groups (for example group A, group B) with one click.
- Sheet tabs at the bottom of the toolbar area, with mailto links and links between sheets working.
- Search with match count and next and previous navigation.
- Row and column outline groups (the +/- groups from Excel) with working collapse buttons.
- Basic conditional formatting: colour scales, data bars, cell comparison rules and text rules.
- Text overflow into empty neighbours, clipped at the first non empty cell, the way Excel draws it.
- Automatic font colour on filled cells picks black or white from the fill brightness, so a dark theme cannot turn black text white on a yellow cell.
- Cell context menu with Copy cell value and Copy cell reference.
- Sheet background stays white even in a dark theme by default, so spreadsheets look like spreadsheets. A toolbar button and a setting switch it to follow the Obsidian theme.

## Installation

### Option A: install script

```bash
./install.sh /path/to/YourVault
```

The script builds the plugin (bun or npm) and copies `main.js`, `manifest.json` and `styles.css` into `<Vault>/.obsidian/plugins/xlsx-styled-viewer/`.

### Option B: manual

1. Build with `bun install && bun run build` (or `npm install && npm run build`).
2. Copy `main.js`, `manifest.json` and `styles.css` into `<Vault>/.obsidian/plugins/xlsx-styled-viewer/`.
3. In Obsidian, turn off Restricted mode, reload plugins and enable **XLSX Styled Viewer**.

## Group filtering with an autofilter

When a sheet has an autofilter with a column of group values in it, there are two ways to hide the groups you are not in:

1. Filter chips: every filter column with 30 or fewer values shows a row of chips under the toolbar. Click a chip to turn that value off. "all" and "none" reset the column.
2. Filter arrow: click the small arrow in a header cell to open the full dropdown with search and checkboxes.

Filtering hides whole rows, so sessions for other groups disappear and the sheet shrinks to what you selected.

## Settings

- Sheet background: Always white (default) or Follow Obsidian theme.
- Show gridlines: draw light gridlines in cells that have no border of their own.
- Show row and column headers: the A B C letters and the row numbers.
- Default zoom: starting zoom, also adjustable from the toolbar.
- Filter chips limit: show chips for filter columns with at most this many values (default 10). Larger columns keep the header dropdown only.
- Show gridlines: still available in settings; the toolbar toggle was removed to reduce clutter.
- Row render limit: how many rows to draw at once for very large sheets. 0 means no limit.

## Command

- Reload the current workbook: re-reads the file from disk, useful while a spreadsheet is still being edited elsewhere.

## What the parser handles

Workbook level: multiple sheets, hidden sheets, shared strings, inline strings, rich text runs (flattened to plain text), theme colours with tints, indexed colours, custom number formats, 1900 and 1904 date systems.

Worksheet level: cell styles with row and column fallbacks, merged ranges, frozen panes, autofilter definitions including preset value and custom filters, hyperlinks (web, mailto and internal sheet links), row and column outline levels, hidden rows and columns, sheet gridline setting.

## Limits

- .xls files are a different binary format and are not supported. Only .xlsx and .xlsm.
- Formulas are shown from their cached values. The plugin does not calculate formulas.
- Images, charts, comments, sparklines and pivot tables are not drawn.
- Pattern fills other than solid are approximated with a translucent fill.
- Text rotation is ignored, the text renders horizontal.
- Conditional formatting supports the common rule types listed above; formula based rules that need a full formula engine are skipped.

## Development

```bash
bun install          # or npm install
bun run build        # bundles src into main.js with esbuild
bun test/run.mjs     # format, parser and renderer tests
bun test/smoke.mjs   # loads the built bundle with a stubbed Obsidian API
```

The workbook tests read two sample files from `test/fixtures/` (not committed, real class lists contain personal data). Put a copy there named `xy-list.xlsx` and `timetable.xlsx`, or set `XY_PATH` and `TT_PATH`. Without fixtures the pure format and colour tests still run.

`bun test/render-html.mjs` writes `test/out/*.html` so you can inspect the rendered markup in a browser.

Source layout:

- `src/main.js`: plugin entry, settings tab, commands.
- `src/view.js`: the Obsidian item view, toolbar, tabs, chips, status bar.
- `src/read.js`: zip and workbook level parsing.
- `src/sheet.js`: worksheet parsing.
- `src/styles.js`: style table parsing and resolution.
- `src/render.js`: the grid renderer.
- `src/numfmt.js`: number format engine.
- `src/color.js`: colour and tint maths.
- `src/util.js`: cell reference and XML helpers.

## Licence

GPL-3.0-or-later. Copyright (C) 2026 Zoroaster1x.

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version. The full licence text is in the LICENSE file.
