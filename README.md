# SheetDesk

SheetDesk is a clean, offline-first desktop spreadsheet app. It ships a real
Excel-like formula engine — dependency graph with minimal recalculation,
dynamic arrays with spill, `LET`/`LAMBDA`, and 96 built-in functions — behind
an Excel-familiar grid UI. No account, no cloud, no telemetry: your workbooks
stay on your machine.

The UI is vanilla JS with no framework and no build step. It runs inside the
Electron shell and also works when `app/index.html` is opened directly in a
browser via `file://` (Electron APIs are feature-detected).

## Features

- **Full 200 × 52 grid** (columns A–AZ) with sticky headers, click /
  shift-click / drag selection, and full keyboard navigation.
- **Formula bar** with `fx` label, bound to the active cell (Enter commits,
  Esc cancels); in-cell editing via F2, double-click, or just typing.
- **Dynamic arrays & spill**: array formulas spill down/right from their
  anchor; spilled cells are read-only mirrors with a dotted outline, and the
  formula bar shows the anchor's formula when one is selected.
- **Formatting**: bold, italic, left/center/right alignment, number formats
  (General, Number, Currency, Percent, Date), font sizes.
- **Sheet tabs**: add (`+`), double-click to rename, `×` to delete, click to
  switch. Cross-sheet formulas work (`=SUM(Sheet2!A1:A10)`).
- **Status bar** with Excel-style `Sum / Average / Count` for the numeric
  cells in the selection, plus a transient message area.
- **Undo/redo**: multi-level (up to 60 snapshots) covering cell edits, paste,
  CSV import, sheet add/rename/delete, and format changes.
- **Clipboard**: Ctrl+C / Ctrl+X / Ctrl+V copies raw values (formulas
  preserved) as TSV; works with other apps via the system clipboard.
- **File support**: native `.sheetdesk.json` workbooks; CSV import (with a
  proper quoted-field parser) and CSV/XLSX export of the used range.
- **Offline**: zero network calls, zero CDNs — even the XLSX library is
  vendored locally.

## Install / run

Prerequisites: [Node.js](https://nodejs.org/) (LTS recommended).

```bash
npm install   # install Electron + electron-builder (dev dependencies)
npm start     # launch the desktop app
```

No `npm install` needed to just look at the UI: open `app/index.html`
directly in a browser and everything works except the native file dialogs
(save/open fall back to browser download/upload).

## Packaging

```bash
npm run dist
```

This runs `electron-builder`, producing installers per the `build` section of
`package.json` (NSIS on Windows, DMG/ZIP on macOS arm64, AppImage on Linux).

### macOS Gatekeeper note

Distributable builds are **unsigned**. On first launch macOS will refuse to
open the app ("cannot be opened because the developer cannot be verified").
To run it: **right-click (or Control-click) the app → Open → Open** in the
dialog. You only need to do this once; afterwards it launches normally.

## Keyboard shortcuts

| Action | Shortcut |
|---|---|
| New workbook | Ctrl/Cmd+N |
| Open workbook | Ctrl/Cmd+O |
| Save | Ctrl/Cmd+S |
| Undo / Redo | Ctrl/Cmd+Z · Ctrl/Cmd+Shift+Z (or Ctrl+Y) |
| Cut / Copy / Paste | Ctrl/Cmd+X · C · V |
| Bold / Italic | Ctrl/Cmd+B · Ctrl/Cmd+I |
| Clear selected cells | Delete or Backspace |
| Move selection | Arrow keys |
| Extend selection | Shift+Arrows (or Shift+Click / drag) |
| Move down / up | Enter / Shift+Enter |
| Move right / left | Tab / Shift+Tab |
| Edit cell in place | F2 or double-click (or just start typing) |
| Commit / cancel edit | Enter / Esc |

## File format

Workbooks are saved as `.sheetdesk.json` — plain JSON:

```json
{
  "app": "sheetdesk",
  "version": 1,
  "sheets": [
    { "name": "Sheet1", "cells": { "A1": { "raw": "=SUM(B1:B3)", "fmt": { "bold": true } } } }
  ]
}
```

- `raw` is exactly what was typed (formulas keep their leading `=`).
- `fmt` holds display formatting: `bold`, `italic`, `align`
  (`left`/`center`/`right`), `numFmt`
  (`general`/`number`/`currency`/`percent`/`date`), `fontSize`.
- Spill ranges are **not** stored — they recompute automatically on load.

## Function reference

96 functions, generated from the engine registry (`F.FUNCTIONS`).

<!-- FUNCTABLE -->
| Function | Description |
|---|---|
| `SUM` | Sum of numbers; text in ranges ignored. |
| `AVERAGE` | Arithmetic mean; #DIV/0! when no numbers. |
| `MIN` | Smallest number; 0 when none. |
| `MAX` | Largest number; 0 when none. |
| `COUNT` | Count of numeric values. |
| `COUNTA` | Count of non-blank values. |
| `PRODUCT` | Product of numbers. |
| `ABS` | Absolute value. |
| `ROUND` | Round half away from zero. |
| `ROUNDUP` | Round away from zero. |
| `ROUNDDOWN` | Round toward zero. |
| `INT` | Round down to integer. |
| `TRUNC` | Truncate toward zero. |
| `MOD` | Remainder with divisor sign (Excel MOD). |
| `POWER` | x raised to y. |
| `SQRT` | Square root; #NUM! for negatives. |
| `PI` | The constant pi. |
| `RAND` | Random [0,1); recalculates always. |
| `RANDBETWEEN` | Random integer in [bottom, top]. |
| `SUMIF` | Sum of sum_range where range matches criteria. |
| `COUNTIF` | Count of cells matching criteria. |
| `AVERAGEIF` | Mean of avg_range where range matches criteria. |
| `IF` | Branch on condition (lifts over arrays). |
| `AND` | TRUE if all args true (short-circuits). |
| `OR` | TRUE if any arg true (short-circuits). |
| `NOT` | Logical negation. |
| `XOR` | TRUE when an odd number of args are true. |
| `IFERROR` | Fallback value when first arg is an error. |
| `CONCAT` | Concatenate values and ranges. |
| `TEXTJOIN` | Join text with delimiter; optionally skip empties. |
| `LEFT` | First n characters. |
| `RIGHT` | Last n characters. |
| `MID` | Characters from 1-based start. |
| `LEN` | Character count. |
| `TRIM` | Strip leading/trailing spaces; collapse inner runs. |
| `UPPER` | Uppercase. |
| `LOWER` | Lowercase. |
| `SUBSTITUTE` | Replace old with new; optional instance number. |
| `TEXT` | Format a value as text ("0.00", "yyyy-mm-dd", ...). |
| `VALUE` | Parse text as a number. |
| `TODAY` | Today's serial date. |
| `NOW` | Now as serial date + time fraction. |
| `DATE` | Serial for y/m/d (month overflow allowed). |
| `YEAR` | Year of a serial date. |
| `MONTH` | Month of a serial date. |
| `DAY` | Day of a serial date. |
| `EDATE` | Serial n months before/after. |
| `EOMONTH` | Last day of the month n months away. |
| `WEEKDAY` | Day of week (type 1: Sun=1..Sat=7). |
| `VLOOKUP` | Vertical lookup; exact unless 4th arg TRUE. |
| `HLOOKUP` | Horizontal lookup; exact unless 4th arg TRUE. |
| `MATCH` | 1-based position of lookup in a vector. |
| `INDEX` | Value at 1-based row/col of an array. |
| `XLOOKUP` | Exact match in lookup vector; optional if-not-found. |
| `CHOOSE` | Return the nth value (only it is evaluated). |
| `ROW` | Row number of a reference (or this cell). |
| `COLUMN` | Column number of a reference (or this cell). |
| `ROWS` | Row count of a reference. |
| `COLUMNS` | Column count of a reference. |
| `ISNUMBER` | TRUE for numbers. |
| `ISTEXT` | TRUE for text. |
| `ISBLANK` | TRUE for blank cells. |
| `ISERROR` | TRUE for any error value. |
| `PROFITMARGIN` | Example custom function: (revenue - cost) / revenue. |
| `TEXTBEFORE` | Text before the nth delimiter (negative n = from the end). |
| `TEXTAFTER` | Text after the nth delimiter (negative n = from the end). |
| `FILTER` | Rows (or columns) of array where include is true; #CALC! when empty. |
| `SORT` | Sort rows by a column (by_col=1 sorts columns by a row). |
| `SORTBY` | Sort rows by one or more key arrays (multi-key). |
| `UNIQUE` | Distinct rows (by_col=1: columns; exactly_once=1: singletons). |
| `CHOOSECOLS` | Pick columns by number (negative = from the end). |
| `CHOOSEROWS` | Pick rows by number (negative = from the end). |
| `DROP` | Drop leading rows/cols (negative = from the end); #CALC! when empty. |
| `TAKE` | Take leading rows/cols (negative = from the end). |
| `HSTACK` | Append arrays side-by-side; shorter ones padded with #N/A. |
| `VSTACK` | Stack arrays top-to-bottom; narrower ones padded with #N/A. |
| `LET` | Bind name/value pairs in order, then evaluate calc. Bindings are eager and sequential. |
| `LAMBDA` | Define a function value. Call via LAMBDA(...)(args) or a LET-bound name. Recursion allowed (depth cap 1000). |
| `IFNA` | Fallback only for #N/A; all other errors propagate. |
| `ISERR` | TRUE for any error except #N/A. |
| `ISNA` | TRUE for #N/A. |
| `ERROR.TYPE` | 1-7 for #NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A; else #N/A. |
| `MEDIAN` | Median; #NUM! when no numbers. |
| `MODE` | Most frequent value; #N/A when nothing repeats. |
| `MODE.SNGL` | Alias of MODE. |
| `LARGE` | k-th largest value. |
| `SMALL` | k-th smallest value. |
| `STDEV.S` | Sample standard deviation. |
| `STDEV.P` | Population standard deviation. |
| `VAR.S` | Sample variance. |
| `VAR.P` | Population variance. |
| `PERCENTILE.INC` | k-th percentile, k in [0,1]. |
| `PERCENTILE.EXC` | k-th percentile, k in (0,1). |
| `QUARTILE.INC` | Quartile 0-4 (inclusive). |
| `QUARTILE.EXC` | Quartile 1-3 (exclusive). |
| `AGGREGATE` | Aggregate with options 0-7 to ignore errors / nested AGGREGATEs. (No hidden rows in v1: options 1,3,5,7 behave like 0,2,4,6.) |
<!-- /FUNCTABLE -->

## Adding a function

All functions live in one registry in `app/formula.js` (`FUNCTIONS`,
name → definition). Lookup is case-insensitive. Copy this pattern — it is the
same shape as the `PROFITMARGIN` example shipped in the engine:

```js
registerFunction('MYFUNC', {
  minArgs: 2,            // required arg count; violations -> #N/A
  maxArgs: 3,            // omit or null for "no upper limit"
  lazy: false,          // true: receive UNEVALUATED thunks (IF / AND / OR / ...)
  volatile: false,      // true: recalculated on EVERY change (RAND, TODAY, NOW)
  propagateErrors: true, // false: your fn receives raw error values
                        //   instead of having them short-circuit first
  passRef: false,       // true: cell/range args arrive as reference
                        //   descriptors {__ref, sheet, c, r, ...}
  astArgs: false,       // true: receive RAW AST nodes (LET names, LAMBDA params)
  desc: 'One-line description shown in docs.',
  fn: function (args, C) {
    // args: evaluated values — number | string | boolean | null (blank),
    //   { __err: '#DIV/0!' }, or { __range: true, values: [[..],[..]] }.
    // C helpers: C.num(v), C.str(v), C.bool(v), C.err(code),
    //   C.scalar(v), C.flat(args).
    // Return a value, or an error via C.err('#CODE!').
    // Do NOT throw for bad input — return an error value instead.
  }
});
```

Conventions for range/array arguments:

- Ranges arrive as `{ __range: true, values: rows }`. Use `C.flat(args)` for a
  simple list of cell values, or read `values[r][c]` directly.
- Follow Excel: aggregate functions (SUM, AVERAGE, …) **ignore** text and
  blank cells inside ranges but **propagate** errors; direct (non-range)
  arguments are coerced (`"5"` → 5, `TRUE` → 1) and non-numeric text is
  `#VALUE!`.
- A missing/omitted optional argument arrives as `undefined`.

After adding a function, add tests in `test/engine.test.js`
(`node test/engine.test.js`) and re-generate the table above with:

```bash
node -e "const F=require('./app/formula.js');for(const n of Object.keys(F.FUNCTIONS)){console.log('| \`'+n+'\` | '+F.FUNCTIONS[n].desc+' |');}"
```

## Dynamic arrays & spill behavior

A formula that returns an array (e.g. `=SORT(A1:A10)`, `=UNIQUE(B1:B20)`,
`=HSTACK(1,2,3)`) **spills** its results into neighboring cells, down and to
the right of the anchor cell holding the formula.

- The **anchor** shows the top-left value of the result.
- **Spilled cells are display mirrors**: they show their spilled value but
  hold no content of their own. They cannot be edited — F2, double-click,
  Delete, formula-bar Enter, and paste onto a spilled cell all show a
  transient status-bar message (`Spilled from A1 — edit the anchor instead`)
  and do nothing. Spilled cells get a subtle dotted outline.
- Selecting a spilled cell shows the **anchor's formula** in the formula bar
  as read-only context (Enter does nothing there).
- If something blocks the spill range, the anchor shows `#SPILL!` and the
  would-be spilled cells stay blank. Clear the blocker (or move the formula)
  and the spill resolves.
- Editing or clearing the anchor automatically updates or removes the spill —
  no extra step needed.
- There is **no `#` spill operator** (`A1#`): referencing it yields `#NAME?`.

## LAMBDA / LET

- `LET(name1, value1, …, calc)` binds names **eagerly and sequentially**, then
  evaluates `calc`. Later bindings can use earlier ones:
  `=LET(x, 5, y, x*2, x+y)` → `15`.
- `LAMBDA(param, …, body)` defines a function value. Call it immediately —
  `=LAMBDA(a, a*3)(7)` → `21` — or bind it with `LET` and call the name:
  `=LET(triple, LAMBDA(a, a*3), triple(7))` → `21`.
- **Honest limits:**
  - `MAP`, `REDUCE`, `SCAN`, `MAKEARRAY`, `BYROW`, `BYCOL` do **not** exist in
    v1 — calling them yields `#NAME?`.
  - There is no `A1#` spill operator (see above).
  - LAMBDA recursion is allowed but capped: nesting deeper than **1000**
    returns `#CALC!`.
  - `IF` with an array condition **lifts elementwise** and spills the result:
    `=IF({1;0;1}, 10, 20)` spills `10, 20, 10`.

## AGGREGATE

`AGGREGATE(function_num, options, ref, [k])` supports function numbers 1–19
and options 0–7:

- Options 2, 3, 6, 7 ignore error values; the others propagate the first
  error encountered.
- Options 0–3 ignore **nested** `AGGREGATE`/`SUBTOTAL` calls — but only when
  they appear as **plain references** (a cell in the referenced range whose
  own formula contains `AGGREGATE`/`SUBTOTAL`). Arrays produced by
  calculations are not screened.
- v1 has **no hidden rows**, so options 1, 3, 5, 7 behave exactly like
  0, 2, 4, 6.
- Note: `SUBTOTAL` itself is not a registered function in v1 (it yields
  `#NAME?`); it is only recognized inside the nested-call screening above.

## Known limitations

- **No fill handle** yet — no drag-to-fill series or copy-down.
- **No print layout** / page setup / print preview.
- Paste inserts raw values/formulas as-is; it does **not** shift relative
  cell references the way Excel does.
- Number formatting is a fixed set (General, Number, Currency, Percent,
  Date); there is no custom format editor, and date display is M/D/YYYY only.
- Charts, conditional formatting, data validation, and frozen panes are not
  implemented.
- The grid is fixed at 200 rows × 52 columns (A–AZ); CSV imports larger than
  that are truncated with a warning.
- Find/replace is not implemented.
- In plain-browser (`file://`) mode there is no native file dialog: saving
  downloads a file and opening uses a file picker, and the app cannot remember
  the save path between Ctrl+S presses (each save downloads again).

## Project layout

```
sheetdesk/
  main.js                  Electron main process (window, menus, file dialogs)
  preload.js               contextBridge: window.api (contextIsolation on)
  package.json             electron ^44 / electron-builder ^26, build config
  app/
    index.html             loads vendor xlsx -> formula.js -> app.js -> styles.css
    app.js                 spreadsheet UI (vanilla JS, no build step)
    formula.js             formula engine (lexer/parser/graph/spill/LAMBDA)
    styles.css             Excel-familiar styling
    vendor/
      xlsx.full.min.js     vendored SheetJS (no CDN)
  test/
    engine.test.js         47 engine tests — run: node test/engine.test.js
```
