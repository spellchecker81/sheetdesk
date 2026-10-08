/* SheetDesk — spreadsheet UI (vanilla JS, no framework, no build step).
 *
 * Works both inside Electron (window.api provided by preload.js) and when
 * app/index.html is opened directly in a browser via file:// (all Electron
 * access is feature-detected via `typeof window.api !== 'undefined'`).
 */
(function () {
'use strict';

var F = window.SheetDeskFormula;
if (!F) {
  document.body.innerHTML = '<p style="padding:2em;font-family:sans-serif">SheetDesk failed to start: formula.js did not load.</p>';
  return;
}

var hasApi = typeof window.api !== 'undefined';
var hasXLSX = typeof window.XLSX !== 'undefined';

var MAX_ROWS = F.MAX_ROWS;
var MAX_COLS = F.MAX_COLS;

/* ------------------------------- state ---------------------------------- */
var wb = new F.Workbook();
var activeSheet = 'Sheet1';
var filePath = null;      // known save path (Electron only)
var dirty = false;

var anchor = { r: 0, c: 0 };          // focused cell: edits land here
var pivot = { r: 0, c: 0 };           // fixed corner for shift-extend
var sel = { r1: 0, c1: 0, r2: 0, c2: 0 }; // normalized selection rectangle

var undoStack = [];
var redoStack = [];
var UNDO_CAP = 60;

var copyBuf = null;   // { text, rows, cols, cut, sheet, r1, c1, r2, c2 }
var dragging = false;
var cellEditor = null; // in-cell edit <input>, when active
var editorCanceled = false;
var msgTimer = null;

/* ------------------------------ DOM refs -------------------------------- */
function $(id) { return document.getElementById(id); }
var gridWrap = $('grid-wrap'), grid = $('grid'), gridBody = $('grid-body');
var formulaInput = $('formula-input'), nameBox = $('namebox');
var statusMsg = $('status-msg'), statusAgg = $('status-agg');
var tabsEl = $('tabs');
var fileOpenInput = $('file-open'), fileCsvInput = $('file-csv');

// cellDivs[r][c] -> inner .v div; lastPaint[r][c] -> {text, cls} cache
var cellDivs = [];
var lastPaint = [];

/* ------------------------------ utilities ------------------------------- */
function clamp(v, lo, hi) { return v < lo ? lo : v > hi ? hi : v; }

function refOf(r, c) { return F.indexToCol(c) + (r + 1); }

function parseRef(ref) {
  var m = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$/.exec(ref);
  if (!m) return null;
  return { c: F.colToIndex(m[2].toUpperCase()), r: parseInt(m[4], 10) - 1 };
}

function normSel(a, b) {
  return {
    r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c),
  };
}

function selRef(s) {
  var a = refOf(s.r1, s.c1);
  var b = refOf(s.r2, s.c2);
  return a === b ? a : a + ':' + b;
}

function baseName(p) {
  var parts = String(p).split(/[/\\]/);
  return parts[parts.length - 1] || 'Untitled';
}

function defaultSaveName() { return 'Untitled.sheetdesk.json'; }

/* --------------------------- spill helpers ------------------------------ */
function spillAtCell(r, c) { return wb.spillAt(activeSheet, c, r); }

function isSpilledMirror(r, c) {
  var s = spillAtCell(r, c);
  return !!s && !s.anchor;
}

function anchorPosOf(spill) {
  // spill.anchorRef is like 'A1' (engine-generated, no $ signs)
  var p = parseRef(spill.anchorRef);
  return p;
}

function spillMessage(spill) {
  showMsg('Spilled from ' + spill.anchorRef + ' — edit the anchor instead');
}

/* --------------------------- value display ------------------------------ */
function fmtThousands(v, dp) {
  return v.toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp });
}

function displayText(v, fmt) {
  if (v === null || v === undefined) return '';
  if (F.isErr(v)) return v.__err;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    var nf = (fmt && fmt.numFmt) || 'general';
    if (nf === 'number') return fmtThousands(v, 2);
    if (nf === 'currency') return '$' + fmtThousands(v, 2);
    if (nf === 'percent') return F.numToString(v * 100) + '%';
    if (nf === 'date') {
      var d = F.serialToDate(Math.round(v));
      return d.m + '/' + d.d + '/' + d.y;
    }
    return F.numToString(v);
  }
  return String(v);
}

/* ------------------------------ status bar ------------------------------ */
function showMsg(text, ms) {
  statusMsg.textContent = text;
  statusAgg.textContent = '';
  if (msgTimer) clearTimeout(msgTimer);
  msgTimer = setTimeout(function () {
    statusMsg.textContent = '';
    updateAggregates();
  }, ms || 2600);
}

function updateAggregates() {
  var sum = 0, n = 0;
  for (var r = sel.r1; r <= sel.r2; r++) {
    for (var c = sel.c1; c <= sel.c2; c++) {
      var v = wb.getValue(activeSheet, c, r);
      if (typeof v === 'number' && isFinite(v)) { sum += v; n++; }
    }
  }
  if (n > 0) {
    statusAgg.textContent =
      'Sum: ' + F.numToString(sum) +
      '   Average: ' + F.numToString(sum / n) +
      '   Count: ' + n;
  } else {
    statusAgg.textContent = '';
  }
}

/* --------------------------- dirty / title ------------------------------ */
function updateTitle() {
  var name = filePath ? baseName(filePath) : 'Untitled';
  document.title = (dirty ? '• ' : '') + name + ' — SheetDesk';
}
function markDirty() { dirty = true; updateTitle(); }
function clearDirty() { dirty = false; updateTitle(); }

/* ------------------------------ undo/redo ------------------------------- */
function pushUndo() {
  undoStack.push(JSON.stringify(wb.toJSON()));
  if (undoStack.length > UNDO_CAP) undoStack.shift();
  redoStack = [];
}

function restoreSnapshot(json) {
  wb.loadJSON(JSON.parse(json));
  if (wb.sheetNames().indexOf(activeSheet) === -1) {
    activeSheet = wb.sheetNames()[0] || 'Sheet1';
  }
  cancelCellEdit();
  renderAll();
  renderTabs();
}

function undo() {
  if (undoStack.length === 0) { showMsg('Nothing to undo'); return; }
  redoStack.push(JSON.stringify(wb.toJSON()));
  restoreSnapshot(undoStack.pop());
  markDirty();
  showMsg('Undone');
}

function redo() {
  if (redoStack.length === 0) { showMsg('Nothing to redo'); return; }
  undoStack.push(JSON.stringify(wb.toJSON()));
  restoreSnapshot(redoStack.pop());
  markDirty();
  showMsg('Redone');
}

/* ------------------------------ grid build ------------------------------ */
function buildGrid() {
  var thead = document.createElement('thead');
  var h = '<tr><th class="corner"></th>';
  for (var c = 0; c < MAX_COLS; c++) h += '<th>' + F.indexToCol(c) + '</th>';
  h += '</tr>';
  thead.innerHTML = h;
  grid.insertBefore(thead, gridBody);

  var parts = [];
  for (var r = 0; r < MAX_ROWS; r++) {
    parts.push('<tr><td class="rowhdr">' + (r + 1) + '</td>');
    for (var c2 = 0; c2 < MAX_COLS; c2++) {
      parts.push('<td class="c" data-r="' + r + '" data-c="' + c2 + '"><div class="v"></div></td>');
    }
    parts.push('</tr>');
  }
  gridBody.innerHTML = parts.join('');

  var rows = gridBody.rows;
  for (var r2 = 0; r2 < MAX_ROWS; r2++) {
    cellDivs.push([]);
    lastPaint.push([]);
    var tds = rows[r2].cells;
    for (var c3 = 0; c3 < MAX_COLS; c3++) {
      cellDivs[r2].push(tds[c3 + 1].firstChild);
      lastPaint[r2].push({ text: null, cls: null, fs: null });
    }
  }
}

function renderCell(r, c) {
  var v = wb.getValue(activeSheet, c, r);
  var fmt = wb.getFormat(activeSheet, c, r);
  var text = displayText(v, fmt);
  var cls = 'v';
  if (r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2) cls += ' sel';
  if (r === anchor.r && c === anchor.c) cls += ' anchor';
  if (isSpilledMirror(r, c)) cls += ' spill';
  if (F.isErr(v)) cls += ' err';
  if (fmt.bold) cls += ' b';
  if (fmt.italic) cls += ' i';
  var al = fmt.align;
  if (!al && typeof v === 'number') al = 'right';
  if (al === 'left') cls += ' al';
  else if (al === 'center') cls += ' ac';
  else if (al === 'right') cls += ' ar';
  var fs = fmt.fontSize ? String(fmt.fontSize) : '';

  var cache = lastPaint[r][c];
  if (cache.text === text && cache.cls === cls && cache.fs === fs) return;
  cache.text = text; cache.cls = cls; cache.fs = fs;

  var el = cellDivs[r][c];
  el.className = cls;
  if (el.textContent !== text) el.textContent = text;
  el.style.fontSize = fs ? fs + 'px' : '';
  var raw = wb.getRaw(activeSheet, c, r);
  if (raw && raw.charAt(0) === '=') el.title = raw;
  else el.title = '';
}

function renderAll() {
  for (var r = 0; r < MAX_ROWS; r++) {
    for (var c = 0; c < MAX_COLS; c++) renderCell(r, c);
  }
  updateFormulaBar();
  updateAggregates();
  updateToolbarState();
}

/* ---------------------------- formula bar ------------------------------- */
function updateFormulaBar() {
  var s = spillAtCell(anchor.r, anchor.c);
  if (s && !s.anchor) {
    // Spilled mirror: show the anchor's formula, bar is read-only context.
    var ap = anchorPosOf(s);
    formulaInput.value = ap ? wb.getRaw(activeSheet, ap.c, ap.r) : '';
    formulaInput.readOnly = true;
  } else {
    formulaInput.value = wb.getRaw(activeSheet, anchor.c, anchor.r);
    formulaInput.readOnly = false;
  }
  nameBox.textContent = selRef(sel);
}

function revertFormulaBar() {
  updateFormulaBar();
  formulaInput.blur();
}

function commitFormulaBar(moveAfter) {
  if (formulaInput.readOnly) return; // spilled mirror: Enter does nothing
  var raw = formulaInput.value;
  var current = wb.getRaw(activeSheet, anchor.c, anchor.r);
  formulaInput.blur();
  if (raw === current) {
    // Nothing changed: just move, no undo snapshot, no dirty flag.
    if (moveAfter === 'down') moveAnchor(1, 0, false);
    else if (moveAfter === 'up') moveAnchor(-1, 0, false);
    return;
  }
  commitCell(anchor.r, anchor.c, raw);
  if (moveAfter === 'down') moveAnchor(1, 0, false);
  else if (moveAfter === 'up') moveAnchor(-1, 0, false);
}

/* ------------------------------ selection ------------------------------- */
function selectCell(r, c, keepInView) {
  anchor = { r: r, c: c };
  pivot = { r: r, c: c };
  sel = { r1: r, c1: c, r2: r, c2: c };
  renderAll();
  if (keepInView !== false) scrollAnchorIntoView();
}

function extendTo(r, c) {
  anchor = { r: r, c: c };
  sel = normSel(pivot, anchor);
  renderAll();
  scrollAnchorIntoView();
}

function moveAnchor(dr, dc, extend) {
  var r = clamp(anchor.r + dr, 0, MAX_ROWS - 1);
  var c = clamp(anchor.c + dc, 0, MAX_COLS - 1);
  if (extend) extendTo(r, c);
  else selectCell(r, c);
}

function scrollAnchorIntoView() {
  var el = cellDivs[anchor.r][anchor.c];
  if (el && el.scrollIntoView) el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

/* ------------------------------- editing -------------------------------- */
function commitCell(r, c, raw) {
  var s = spillAtCell(r, c);
  if (s && !s.anchor) { spillMessage(s); return false; }
  pushUndo();
  wb.setCell(activeSheet, c, r, raw);
  markDirty();
  renderAll();
  return true;
}

function clearSelection() {
  var r, c, s;
  for (r = sel.r1; r <= sel.r2; r++) {
    for (c = sel.c1; c <= sel.c2; c++) {
      s = spillAtCell(r, c);
      if (s && !s.anchor) { spillMessage(s); return; }
    }
  }
  pushUndo();
  for (r = sel.r1; r <= sel.r2; r++) {
    for (c = sel.c1; c <= sel.c2; c++) wb.setCell(activeSheet, c, r, '');
  }
  markDirty();
  renderAll();
}

/* --------------------------- in-cell editor ----------------------------- */
function startCellEdit(initialText) {
  if (cellEditor) return;
  var s = spillAtCell(anchor.r, anchor.c);
  if (s && !s.anchor) { spillMessage(s); return; }
  var rect = cellDivs[anchor.r][anchor.c].getBoundingClientRect();
  var input = document.createElement('input');
  input.id = 'cell-editor';
  input.setAttribute('spellcheck', 'false');
  input.setAttribute('autocomplete', 'off');
  input.style.left = rect.left + 'px';
  input.style.top = rect.top + 'px';
  input.style.width = Math.max(rect.width + 24, 120) + 'px';
  input.style.height = rect.height + 'px';
  input.value = (initialText !== undefined && initialText !== null)
    ? initialText
    : wb.getRaw(activeSheet, anchor.c, anchor.r);
  document.body.appendChild(input);
  cellEditor = input;
  editorCanceled = false;
  input.focus();
  if (initialText) {
    input.setSelectionRange(input.value.length, input.value.length);
  } else {
    input.select();
  }
}

function finishCellEdit(move) {
  if (!cellEditor) return;
  var input = cellEditor;
  cellEditor = null;
  var raw = input.value;
  var wasCanceled = editorCanceled;
  input.remove();
  if (wasCanceled) return;
  var current = wb.getRaw(activeSheet, anchor.c, anchor.r);
  if (raw === current) {
    if (move === 'down') moveAnchor(1, 0, false);
    else if (move === 'up') moveAnchor(-1, 0, false);
    else if (move === 'right') moveAnchor(0, 1, false);
    else if (move === 'left') moveAnchor(0, -1, false);
    return;
  }
  if (commitCell(anchor.r, anchor.c, raw)) {
    if (move === 'down') moveAnchor(1, 0, false);
    else if (move === 'up') moveAnchor(-1, 0, false);
    else if (move === 'right') moveAnchor(0, 1, false);
    else if (move === 'left') moveAnchor(0, -1, false);
  }
}

function cancelCellEdit() {
  if (!cellEditor) return;
  editorCanceled = true;
  var input = cellEditor;
  cellEditor = null;
  input.remove();
}

/* ------------------------------- formats -------------------------------- */
function applyFormat(patch) {
  pushUndo();
  for (var r = sel.r1; r <= sel.r2; r++) {
    for (var c = sel.c1; c <= sel.c2; c++) {
      var s = spillAtCell(r, c);
      if (s && !s.anchor) continue; // never write formats onto spill mirrors
      wb.setFormat(activeSheet, c, r, patch);
    }
  }
  markDirty();
  renderAll();
}

function toggleBold() {
  var fmt = wb.getFormat(activeSheet, anchor.c, anchor.r);
  applyFormat({ bold: !fmt.bold });
}

function toggleItalic() {
  var fmt = wb.getFormat(activeSheet, anchor.c, anchor.r);
  applyFormat({ italic: !fmt.italic });
}

function updateToolbarState() {
  var fmt = wb.getFormat(activeSheet, anchor.c, anchor.r);
  $('btn-bold').classList.toggle('on', !!fmt.bold);
  $('btn-italic').classList.toggle('on', !!fmt.italic);
  $('sel-numfmt').value = fmt.numFmt || 'general';
  $('sel-fontsize').value = fmt.fontSize ? String(fmt.fontSize) : '';
  $('btn-align-left').classList.toggle('on', fmt.align === 'left');
  $('btn-align-center').classList.toggle('on', fmt.align === 'center');
  $('btn-align-right').classList.toggle('on', fmt.align === 'right');
}

/* ------------------------------ clipboard ------------------------------- */
function selectionToTSV() {
  var lines = [];
  for (var r = sel.r1; r <= sel.r2; r++) {
    var fields = [];
    for (var c = sel.c1; c <= sel.c2; c++) {
      fields.push(wb.getRaw(activeSheet, c, r));
    }
    lines.push(fields.join('\t'));
  }
  return lines.join('\n');
}

function copySelection(cut) {
  var text = selectionToTSV();
  copyBuf = {
    text: text,
    rows: sel.r2 - sel.r1 + 1,
    cols: sel.c2 - sel.c1 + 1,
    cut: !!cut,
    sheet: activeSheet,
    r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2,
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).catch(function () { /* internal buffer suffices */ });
  }
  showMsg(cut ? 'Cut' : 'Copied');
}

function parseTSV(text) {
  return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
    .split('\n').map(function (line) { return line.split('\t'); });
}

function pasteClipboard() {
  var useText = function (text) {
    if (!text) { showMsg('Clipboard is empty'); return; }
    doPaste(text);
  };
  if (navigator.clipboard && navigator.clipboard.readText) {
    navigator.clipboard.readText().then(
      function (text) {
        if (text) useText(text);
        else if (copyBuf) useText(copyBuf.text);
        else showMsg('Clipboard is empty');
      },
      function () {
        if (copyBuf) useText(copyBuf.text);
        else showMsg('Clipboard read was blocked — copy again inside SheetDesk');
      }
    );
  } else if (copyBuf) {
    useText(copyBuf.text);
  } else {
    showMsg('Clipboard is empty');
  }
}

function doPaste(text) {
  var s0 = spillAtCell(anchor.r, anchor.c);
  if (s0 && !s0.anchor) { spillMessage(s0); return; }

  var rows = parseTSV(text);
  // Drop a single trailing empty line produced by a final newline.
  if (rows.length > 1 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') {
    rows.pop();
  }
  var nRows = rows.length;
  var nCols = Math.max.apply(null, rows.map(function (x) { return x.length; }));

  var truncated = false;
  var maxR = Math.min(anchor.r + nRows, MAX_ROWS);
  var maxC = Math.min(anchor.c + nCols, MAX_COLS);
  if (anchor.r + nRows > MAX_ROWS || anchor.c + nCols > MAX_COLS) truncated = true;

  // Pre-flight: refuse if any target cell is a spill mirror.
  for (var r = anchor.r; r < maxR; r++) {
    for (var c = anchor.c; c < maxC; c++) {
      var s = spillAtCell(r, c);
      if (s && !s.anchor) { spillMessage(s); return; }
    }
  }

  pushUndo();
  var ri, ci;
  for (ri = 0; ri < maxR - anchor.r; ri++) {
    var line = rows[ri] || [];
    for (ci = 0; ci < maxC - anchor.c; ci++) {
      var val = ri < rows.length && ci < line.length ? line[ci] : '';
      wb.setCell(activeSheet, anchor.c + ci, anchor.r + ri, val);
    }
  }
  // Cut: clear the source range after a successful paste.
  if (copyBuf && copyBuf.cut && copyBuf.sheet === activeSheet) {
    for (var sr = copyBuf.r1; sr <= copyBuf.r2; sr++) {
      for (var sc = copyBuf.c1; sc <= copyBuf.c2; sc++) {
        wb.setCell(activeSheet, sc, sr, '');
      }
    }
    copyBuf = null;
  }
  sel = { r1: anchor.r, c1: anchor.c, r2: maxR - 1, c2: maxC - 1 };
  pivot = { r: anchor.r, c: anchor.c };
  markDirty();
  renderAll();
  showMsg(truncated ? 'Pasted (truncated to grid bounds)' : 'Pasted');
}

/* -------------------------------- sheets -------------------------------- */
function renderTabs() {
  tabsEl.innerHTML = '';
  wb.sheetNames().forEach(function (name) {
    var tab = document.createElement('button');
    tab.className = 'sheet-tab' + (name === activeSheet ? ' active' : '');
    var label = document.createElement('span');
    label.textContent = name;
    label.title = 'Double-click to rename';
    label.addEventListener('dblclick', function (e) {
      e.stopPropagation();
      renameSheet(name);
    });
    var x = document.createElement('span');
    x.className = 'tab-x';
    x.textContent = '×';
    x.title = 'Delete sheet';
    x.addEventListener('click', function (e) {
      e.stopPropagation();
      deleteSheet(name);
    });
    tab.appendChild(label);
    tab.appendChild(x);
    tab.addEventListener('click', function () { switchSheet(name); });
    tabsEl.appendChild(tab);
  });
}

function switchSheet(name) {
  if (wb.sheetNames().indexOf(name) === -1) return;
  cancelCellEdit();
  activeSheet = name;
  selectCell(0, 0, false);
  renderTabs();
  gridWrap.scrollTop = 0;
  gridWrap.scrollLeft = 0;
}

function suggestSheetName() {
  var n = wb.sheetNames().length + 1;
  var name = 'Sheet' + n;
  while (wb.sheetNames().indexOf(name) !== -1) { n++; name = 'Sheet' + n; }
  return name;
}

function addSheet() {
  var name = prompt('New sheet name:', suggestSheetName());
  if (name === null) return;
  name = name.trim();
  if (!name) { showMsg('Sheet name cannot be empty'); return; }
  pushUndo();
  var actual = wb.addSheet(name);
  activeSheet = actual;
  markDirty();
  selectCell(0, 0, false);
  renderTabs();
  showMsg('Sheet "' + actual + '" added');
}

function renameSheet(oldName) {
  var name = prompt('Rename sheet "' + oldName + '" to:', oldName);
  if (name === null) return;
  name = name.trim();
  if (!name) { showMsg('Sheet name cannot be empty'); return; }
  if (name === oldName) return;
  pushUndo();
  if (wb.renameSheet(oldName, name)) {
    activeSheet = name;
    markDirty();
    renderTabs();
  } else {
    undoStack.pop(); // revert the unused snapshot
    showMsg('Could not rename — a sheet with that name may already exist');
  }
}

function deleteSheet(name) {
  if (wb.sheetNames().length <= 1) {
    showMsg('A workbook needs at least one sheet');
    return;
  }
  if (!confirm('Delete sheet "' + name + '"? This cannot be undone except via Undo.')) return;
  pushUndo();
  if (wb.removeSheet(name)) {
    if (activeSheet === name) activeSheet = wb.sheetNames()[0];
    markDirty();
    selectCell(0, 0, false);
    renderTabs();
    showMsg('Sheet "' + name + '" deleted');
  } else {
    undoStack.pop();
    showMsg('Could not delete sheet "' + name + '"');
  }
}

/* ------------------------------- file I/O ------------------------------- */
function downloadBlob(blob, filename) {
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(function () {
    URL.revokeObjectURL(a.href);
    a.remove();
  }, 500);
}

function serializeWorkbook() {
  return JSON.stringify(wb.toJSON());
}

async function saveWorkbook() {
  var data = serializeWorkbook();
  if (hasApi) {
    try {
      if (filePath) {
        var res = await window.api.saveFileTo(filePath, data);
        if (res && res.path) {
          filePath = res.path;
          clearDirty();
          showMsg('Saved');
        }
      } else {
        await saveWorkbookAs();
      }
    } catch (e) {
      showMsg('Save failed: ' + (e && e.message ? e.message : e));
    }
  } else {
    downloadBlob(new Blob([data], { type: 'application/json' }), defaultSaveName());
    clearDirty();
    showMsg('Downloaded ' + defaultSaveName());
  }
}

async function saveWorkbookAs() {
  var data = serializeWorkbook();
  if (hasApi) {
    try {
      var res = await window.api.saveFileAs(defaultSaveName(), data);
      if (res && res.path) {
        filePath = res.path;
        clearDirty();
        showMsg('Saved to ' + baseName(filePath));
      }
    } catch (e) {
      showMsg('Save failed: ' + (e && e.message ? e.message : e));
    }
  } else {
    downloadBlob(new Blob([data], { type: 'application/json' }), defaultSaveName());
    clearDirty();
    showMsg('Downloaded ' + defaultSaveName());
  }
}

function confirmDiscard() {
  return !dirty || confirm('You have unsaved changes. Discard them?');
}

function loadWorkbookData(data, path) {
  var obj;
  try {
    obj = JSON.parse(data);
  } catch (e) {
    showMsg('Could not open file: not valid JSON');
    return false;
  }
  if (!obj || obj.app !== 'sheetdesk' || !Array.isArray(obj.sheets) || obj.sheets.length === 0) {
    showMsg('Could not open file: not a SheetDesk workbook');
    return false;
  }
  wb.loadJSON(obj);
  var names = wb.sheetNames();
  activeSheet = names[0];
  filePath = path || null;
  undoStack = [];
  redoStack = [];
  copyBuf = null;
  clearDirty();
  cancelCellEdit();
  selectCell(0, 0, false);
  renderTabs();
  gridWrap.scrollTop = 0;
  gridWrap.scrollLeft = 0;
  showMsg(path ? 'Opened ' + baseName(path) : 'Workbook loaded');
  return true;
}

function newWorkbook() {
  if (!confirmDiscard()) return;
  cancelCellEdit();
  wb = new F.Workbook();
  activeSheet = wb.sheetNames()[0];
  filePath = null;
  undoStack = [];
  redoStack = [];
  copyBuf = null;
  clearDirty();
  selectCell(0, 0, false);
  renderTabs();
  gridWrap.scrollTop = 0;
  gridWrap.scrollLeft = 0;
  showMsg('New workbook');
}

function openWorkbook() {
  if (!confirmDiscard()) return;
  if (hasApi) {
    window.api.openFile().then(function (res) {
      if (res && res.data) loadWorkbookData(res.data, res.path);
    }).catch(function (e) {
      showMsg('Open failed: ' + (e && e.message ? e.message : e));
    });
  } else {
    fileOpenInput.value = '';
    fileOpenInput.click();
  }
}

/* ------------------------------ CSV import ------------------------------ */
function parseCSV(text) {
  var rows = [], row = [], field = '', inQ = false;
  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);
    if (inQ) {
      if (ch === '"') {
        if (text.charAt(i + 1) === '"') { field += '"'; i++; }
        else inQ = false;
      } else {
        field += ch;
      }
    } else {
      if (ch === '"') inQ = true;
      else if (ch === ',') { row.push(field); field = ''; }
      else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
      else if (ch !== '\r') field += ch;
    }
  }
  row.push(field);
  rows.push(row);
  // Drop trailing all-empty rows (e.g. from a final newline).
  while (rows.length > 0 && rows[rows.length - 1].every(function (f) { return f === ''; })) {
    rows.pop();
  }
  return rows;
}

function importCSVData(text) {
  var rows = parseCSV(text);
  if (rows.length === 0) { showMsg('CSV file is empty'); return; }
  if (wb.usedRange(activeSheet)) {
    if (!confirm('The active sheet is not empty. Importing will overwrite cells starting at A1. Continue?')) return;
  }
  var truncated = false;
  pushUndo();
  var nRows = Math.min(rows.length, MAX_ROWS);
  if (rows.length > MAX_ROWS) truncated = true;
  for (var r = 0; r < nRows; r++) {
    var nCols = Math.min(rows[r].length, MAX_COLS);
    if (rows[r].length > MAX_COLS) truncated = true;
    for (var c = 0; c < nCols; c++) {
      wb.setCell(activeSheet, c, r, rows[r][c]);
    }
  }
  markDirty();
  selectCell(0, 0, false);
  renderAll();
  showMsg(truncated ? 'CSV imported (truncated to grid bounds)' : 'CSV imported into ' + activeSheet);
}

function importCSV() {
  fileCsvInput.value = '';
  fileCsvInput.click();
}

/* ------------------------------ CSV export ------------------------------ */
function exportScalar(v) {
  if (v === null || v === undefined) return '';
  if (F.isErr(v)) return v.__err;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return F.numToString(v);
  return String(v);
}

function csvEscape(field) {
  if (/[",\n\r]/.test(field)) return '"' + field.replace(/"/g, '""') + '"';
  return field;
}

function exportCSV() {
  var ur = wb.usedRange(activeSheet);
  if (!ur) { showMsg('Nothing to export'); return; }
  var lines = [];
  for (var r = ur.r1; r <= ur.r2; r++) {
    var fields = [];
    for (var c = ur.c1; c <= ur.c2; c++) {
      fields.push(csvEscape(exportScalar(wb.getValue(activeSheet, c, r))));
    }
    lines.push(fields.join(','));
  }
  var csv = lines.join('\r\n');
  var name = (filePath ? baseName(filePath).replace(/\.sheetdesk\.json$/i, '') : 'SheetDesk')
    + '-' + activeSheet + '.csv';
  if (hasApi) {
    window.api.saveFileAs(name, csv, {
      title: 'Export CSV',
      filters: [
        { name: 'CSV files', extensions: ['csv'] },
        { name: 'All files', extensions: ['*'] },
      ],
    }).then(function (res) {
      if (res && res.path) showMsg('Exported ' + baseName(res.path));
    }).catch(function (e) {
      showMsg('Export failed: ' + (e && e.message ? e.message : e));
    });
  } else {
    downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8' }), name);
    showMsg('Downloaded ' + name);
  }
}

/* ------------------------------ XLSX export ----------------------------- */
function exportXLSX() {
  if (!hasXLSX) {
    showMsg('XLSX export is unavailable (spreadsheet library failed to load)');
    return;
  }
  var ur = wb.usedRange(activeSheet);
  if (!ur) { showMsg('Nothing to export'); return; }
  var aoa = [];
  for (var r = ur.r1; r <= ur.r2; r++) {
    var row = [];
    for (var c = ur.c1; c <= ur.c2; c++) {
      var v = wb.getValue(activeSheet, c, r);
      if (v === null || v === undefined) row.push('');
      else if (F.isErr(v)) row.push(v.__err);
      else row.push(v);
    }
    aoa.push(row);
  }
  var ws = window.XLSX.utils.aoa_to_sheet(aoa);
  var book = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(book, ws, activeSheet);
  var name = (filePath ? baseName(filePath).replace(/\.sheetdesk\.json$/i, '') : 'SheetDesk')
    + '-' + activeSheet + '.xlsx';
  try {
    // Blob download works in plain browsers and in the Electron renderer
    // (no node integration there), so one path covers both.
    var out = window.XLSX.write(book, { bookType: 'xlsx', type: 'array' });
    downloadBlob(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
    showMsg('Downloaded ' + name);
  } catch (e) {
    showMsg('XLSX export failed: ' + (e && e.message ? e.message : e));
  }
}

/* ------------------------------ menu actions ---------------------------- */
var MENU_ACTIONS = {
  'new': newWorkbook,
  'open': openWorkbook,
  'save': saveWorkbook,
  'save-as': saveWorkbookAs,
  'import-csv': importCSV,
  'export-csv': exportCSV,
  'export-xlsx': exportXLSX,
  'undo': undo,
  'redo': redo,
  'cut': function () { copySelection(true); },
  'copy': function () { copySelection(false); },
  'paste': pasteClipboard,
};

/* -------------------------------- events -------------------------------- */
function bindEvents() {
  // --- mouse selection ---
  gridBody.addEventListener('mousedown', function (e) {
    if (e.button !== 0) return;
    var td = e.target.closest ? e.target.closest('td.c') : null;
    if (!td) return;
    if (cellEditor) finishCellEdit('stay');
    var r = parseInt(td.getAttribute('data-r'), 10);
    var c = parseInt(td.getAttribute('data-c'), 10);
    if (e.shiftKey) extendTo(r, c);
    else selectCell(r, c);
    dragging = true;
    e.preventDefault();
  });
  gridBody.addEventListener('mouseover', function (e) {
    if (!dragging) return;
    var td = e.target.closest ? e.target.closest('td.c') : null;
    if (!td) return;
    extendTo(parseInt(td.getAttribute('data-r'), 10), parseInt(td.getAttribute('data-c'), 10));
  });
  document.addEventListener('mouseup', function () { dragging = false; });
  gridBody.addEventListener('dblclick', function (e) {
    var td = e.target.closest ? e.target.closest('td.c') : null;
    if (!td) return;
    dragging = false;
    startCellEdit('');
  });
  gridWrap.addEventListener('scroll', function () { cancelCellEdit(); }, { passive: true });

  // --- formula bar ---
  formulaInput.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') {
      e.preventDefault();
      commitFormulaBar(e.shiftKey ? 'up' : 'down');
    } else if (e.key === 'Escape') {
      e.preventDefault();
      revertFormulaBar();
    }
  });
  formulaInput.addEventListener('focus', function () {
    // Keep the bar in sync if the anchor changed while it was blurred.
    updateFormulaBar();
  });

  // --- hidden file inputs ---
  fileOpenInput.addEventListener('change', function () {
    var f = fileOpenInput.files[0];
    if (!f) return;
    var reader = new FileReader();
    reader.onload = function () { loadWorkbookData(reader.result, null); };
    reader.onerror = function () { showMsg('Could not read file'); };
    reader.readAsText(f);
  });
  fileCsvInput.addEventListener('change', function () {
    var f = fileCsvInput.files[0];
    if (!f) return;
    var reader = new FileReader();
    reader.onload = function () { importCSVData(reader.result); };
    reader.onerror = function () { showMsg('Could not read CSV file'); };
    reader.readAsText(f);
  });

  // --- toolbar ---
  $('btn-new').addEventListener('click', newWorkbook);
  $('btn-open').addEventListener('click', openWorkbook);
  $('btn-save').addEventListener('click', saveWorkbook);
  $('btn-saveas').addEventListener('click', saveWorkbookAs);
  $('btn-import').addEventListener('click', importCSV);
  $('btn-export-csv').addEventListener('click', exportCSV);
  $('btn-export-xlsx').addEventListener('click', exportXLSX);
  $('btn-undo').addEventListener('click', undo);
  $('btn-redo').addEventListener('click', redo);
  $('btn-bold').addEventListener('click', toggleBold);
  $('btn-italic').addEventListener('click', toggleItalic);
  $('sel-numfmt').addEventListener('change', function (e) {
    applyFormat({ numFmt: e.target.value });
    formulaInput.focus();
  });
  $('sel-fontsize').addEventListener('change', function (e) {
    if (!e.target.value) return;
    applyFormat({ fontSize: parseInt(e.target.value, 10) });
    formulaInput.focus();
  });
  $('btn-align-left').addEventListener('click', function () { applyFormat({ align: 'left' }); });
  $('btn-align-center').addEventListener('click', function () { applyFormat({ align: 'center' }); });
  $('btn-align-right').addEventListener('click', function () { applyFormat({ align: 'right' }); });

  // --- sheet tabs ---
  $('btn-add-sheet').addEventListener('click', addSheet);

  // --- keyboard ---
  document.addEventListener('keydown', onKeyDown);
}

function onKeyDown(e) {
  var inFormula = document.activeElement === formulaInput;
  var inEditor = !!cellEditor && document.activeElement === cellEditor;
  var mod = e.ctrlKey || e.metaKey;

  if (inFormula || inEditor) {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (inEditor) finishCellEdit(e.shiftKey ? 'up' : 'down');
      else commitFormulaBar(e.shiftKey ? 'up' : 'down');
    } else if (e.key === 'Escape') {
      e.preventDefault();
      if (inEditor) cancelCellEdit();
      else revertFormulaBar();
    } else if (e.key === 'Tab' && inEditor) {
      e.preventDefault();
      finishCellEdit(e.shiftKey ? 'left' : 'right');
    } else if (mod && (e.key === 's' || e.key === 'S')) {
      // Allow save even while typing in the formula bar: commit first.
      e.preventDefault();
      if (inFormula) commitFormulaBar('stay');
      else if (inEditor) finishCellEdit('stay');
      saveWorkbook();
    }
    return; // otherwise let inputs behave natively (incl. Ctrl+Z text undo)
  }

  var key = e.key;

  // --- Ctrl/Cmd shortcuts ---
  if (mod) {
    var k = key.toLowerCase();
    if (k === 's') { e.preventDefault(); saveWorkbook(); return; }
    if (k === 'o') { e.preventDefault(); openWorkbook(); return; }
    if (k === 'n') { e.preventDefault(); newWorkbook(); return; }
    if (k === 'z' && !e.shiftKey) { e.preventDefault(); undo(); return; }
    if ((k === 'z' && e.shiftKey) || k === 'y') { e.preventDefault(); redo(); return; }
    if (k === 'x') { e.preventDefault(); copySelection(true); return; }
    if (k === 'c') { e.preventDefault(); copySelection(false); return; }
    if (k === 'v') { e.preventDefault(); pasteClipboard(); return; }
    if (k === 'b') { e.preventDefault(); toggleBold(); return; }
    if (k === 'i') { e.preventDefault(); toggleItalic(); return; }
    return;
  }

  // --- navigation & editing ---
  switch (key) {
    case 'ArrowUp': e.preventDefault(); moveAnchor(-1, 0, e.shiftKey); return;
    case 'ArrowDown': e.preventDefault(); moveAnchor(1, 0, e.shiftKey); return;
    case 'ArrowLeft': e.preventDefault(); moveAnchor(0, -1, e.shiftKey); return;
    case 'ArrowRight': e.preventDefault(); moveAnchor(0, 1, e.shiftKey); return;
    case 'Enter': e.preventDefault(); moveAnchor(e.shiftKey ? -1 : 1, 0, false); return;
    case 'Tab': e.preventDefault(); moveAnchor(0, e.shiftKey ? -1 : 1, false); return;
    case 'F2': e.preventDefault(); startCellEdit(''); return;
    case 'Delete':
    case 'Backspace': e.preventDefault(); clearSelection(); return;
    case 'Escape': cancelCellEdit(); return;
  }

  // Typing a printable character starts in-cell editing (Excel-style).
  if (key.length === 1 && !e.altKey) {
    e.preventDefault();
    startCellEdit(key);
  }
}

/* --------------------------------- init --------------------------------- */
function init() {
  buildGrid();
  bindEvents();
  if (hasXLSX) {
    var xbtn = $('btn-export-xlsx');
    if (xbtn) xbtn.hidden = false;
  }
  if (hasApi && typeof window.api.onMenuAction === 'function') {
    window.api.onMenuAction(function (name) {
      var fn = MENU_ACTIONS[name];
      if (fn) fn();
    });
  }
  activeSheet = wb.sheetNames()[0];
  renderTabs();
  selectCell(0, 0, false);
  clearDirty();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}

})();
