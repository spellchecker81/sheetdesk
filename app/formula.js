/*!
 * SheetDesk formula engine (app/formula.js)
 * -----------------------------------------
 * A dependency-free Excel-like formula engine. Works in the browser as a plain
 * <script> (exposes window.SheetDeskFormula) and in Node via module.exports,
 * so the same code is unit-tested directly with node.
 *
 * Features: lexer + recursive-descent parser, cell/range/sheet-qualified
 * references, full operator set, dependency graph with topological recalc,
 * cycle detection, minimal dependent recalculation, volatile functions,
 * Excel serial dates (incl. the 1900 leap-year quirk), and an extensible
 * function registry (see "ADDING A FUNCTION" below).
 */

/* ---------------------------------------------------------------------------
 * ADDING A FUNCTION — the documented pattern
 * ---------------------------------------------------------------------------
 * All functions live in ONE registry object: FUNCTIONS (name -> definition).
 * Lookup is case-insensitive: registerFunction() upper-cases the name, and the
 * evaluator upper-cases every call name before lookup.
 *
 *   registerFunction('MYFUNC', {
 *     minArgs: 2,            // required arg count; violations -> #N/A
 *     maxArgs: 3,            // omit or null for "no upper limit"
 *     lazy: false,           // true: receive UNEVALUATED thunks (needed for
 *                            //   IF / AND / OR / IFERROR / CHOOSE so branches
 *                            //   short-circuit instead of evaluating eagerly)
 *     volatile: false,       // true: recalculated on EVERY change
 *                            //   (RAND, RANDBETWEEN, TODAY, NOW)
 *     propagateErrors: true, // false: your fn receives raw error values
 *                            //   instead of having them short-circuit first
 *                            //   (needed for ISERROR, ISBLANK, ...)
 *     passRef: false,        // true: cell/range args arrive as reference
 *                            //   descriptors {__ref, sheet, c, r, ...}
 *                            //   instead of values (needed for ROW, COLUMN, ...)
 *     astArgs: false,        // true: receive RAW AST nodes, not values
 *                            //   (needed for LET names and LAMBDA params/body;
 *                            //   evaluate them yourself via C.eval(node, ctx))
 *     desc: 'One-line description shown in docs.',
 *     fn: function (args, C) {
 *       // args: evaluated values. Each is one of:
 *       //   number | string | boolean | null (blank cell)
 *       //   { __err: '#DIV/0!' }            (an error value)
 *       //   { __range: true, values: [[v,..],[..],..] }  (a range argument)
 *       //
 *       // C: helper object
 *       //   C.num(v)    -> number or error value  ("" -> #VALUE!, TRUE -> 1)
 *       //   C.str(v)    -> string or error value  (TRUE -> "TRUE")
 *       //   C.bool(v)   -> boolean or error value
 *       //   C.err(code) -> error value, e.g. C.err('#VALUE!')
 *       //   C.scalar(v) -> v, or #VALUE! if v is a range
 *       //   C.flat(args)-> args with every range expanded to its cell values
 *       //
 *       // Return a value, or an error via C.err('#CODE!').
 *       // Do NOT throw for bad input — return an error value instead.
 *     }
 *   });
 *
 * Conventions for range/array arguments:
 *   - Ranges arrive as { __range: true, values: rows }. Use C.flat(args) to get
 *     a simple list of the cells' values, or read values[r][c] directly.
 *   - Follow Excel: aggregate functions (SUM, AVERAGE, ...) IGNORE text and
 *     blank cells inside ranges but PROPAGATE errors; direct (non-range)
 *     arguments are coerced ("5" -> 5, TRUE -> 1) and non-numeric text is #VALUE!.
 *   - A missing/omitted optional argument arrives as undefined.
 * -------------------------------------------------------------------------*/

(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; }
  else { root.SheetDeskFormula = api; }
})(typeof self !== 'undefined' ? self : this, function () {
'use strict';

/* ============================ configuration ============================ */
var MAX_ROWS = 200;
var MAX_COLS = 52; // A..AZ — raise freely; grid UI reads these constants

/* ============================ error values ============================= */
function sdErr(code) { return { __err: code }; }
function isErr(v) { return !!v && typeof v === 'object' && typeof v.__err === 'string'; }
var ERROR_CODES = ['#DIV/0!', '#NAME?', '#VALUE!', '#REF!', '#N/A', '#NUM!', '#NULL!', '#CYCLE!', '#SPILL!', '#CALC!'];

/* ======================= column / reference helpers ==================== */
function indexToCol(i) {
  var s = '';
  i++;
  do { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } while (i > 0);
  return s;
}
function colToIndex(col) {
  var n = 0;
  for (var k = 0; k < col.length; k++) n = n * 26 + (col.charCodeAt(k) - 64);
  return n - 1;
}
// "$A$1" -> {absC, col:'A', absR, row:1} ; null when not a reference
function splitRef(text) {
  var m = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$/.exec(text);
  if (!m) return null;
  return { absC: m[1] === '$', col: m[2].toUpperCase(), absR: m[3] === '$', row: parseInt(m[4], 10) };
}

/* ============================ date helpers ============================= */
// Excel serial dates: 1900-01-01 == 1. Excel (incorrectly) treats 1900 as a
// leap year, so serial 60 == the non-existent 1900-02-29 and every serial
// >= 61 is one higher than the true day count. We reproduce that quirk.
var DATE_EPOCH = Date.UTC(1899, 11, 31); // serial 1 == 1900-01-01
function dateToSerial(y, m, d) {
  var actual = Math.round((Date.UTC(y, m - 1, d) - DATE_EPOCH) / 86400000);
  return actual < 60 ? actual : actual + 1;
}
function serialToDate(s) {
  var actual = s > 60 ? s - 1 : s;
  var d = new Date(DATE_EPOCH + Math.round(actual) * 86400000);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}
function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

/* ============================ number helpers =========================== */
function numRes(v) { // map non-finite results to #NUM! like Excel
  return (typeof v === 'number' && !isFinite(v)) ? sdErr('#NUM!') : v;
}
function numToString(v) {
  if (typeof v !== 'number' || !isFinite(v)) return '#NUM!';
  if (v === 0) return '0';
  return String(parseFloat(v.toPrecision(12))); // kills 0.1+0.2-style noise
}

/* ================================ lexer ================================ */
function tokenize(src) {
  var toks = [], i = 0, n = src.length;
  function isSp(c) { return c === ' ' || c === '\t' || c === '\n' || c === '\r'; }
  while (i < n) {
    var c = src[i];
    if (isSp(c)) { i++; continue; }
    // number (incl. 1e3, .5)
    if ((c >= '0' && c <= '9') || (c === '.' && i + 1 < n && src[i + 1] >= '0' && src[i + 1] <= '9')) {
      var j = i, seenDot = false;
      while (j < n && ((src[j] >= '0' && src[j] <= '9') || (src[j] === '.' && !seenDot))) {
        if (src[j] === '.') seenDot = true;
        j++;
      }
      if (j < n && (src[j] === 'e' || src[j] === 'E')) {
        var k = j + 1;
        if (src[k] === '+' || src[k] === '-') k++;
        if (k < n && src[k] >= '0' && src[k] <= '9') {
          while (k < n && src[k] >= '0' && src[k] <= '9') k++;
          j = k;
        }
      }
      toks.push({ t: 'num', v: parseFloat(src.slice(i, j)) }); i = j; continue;
    }
    // double-quoted string, "" escapes a quote
    if (c === '"') {
      var s = '', p = i + 1;
      while (p < n) {
        if (src[p] === '"') { if (src[p + 1] === '"') { s += '"'; p += 2; } else { p++; break; } }
        else { s += src[p]; p++; }
      }
      toks.push({ t: 'str', v: s }); i = p; continue;
    }
    // single-quoted sheet name 'My Sheet'
    if (c === "'") {
      var q = '', r = i + 1;
      while (r < n) {
        if (src[r] === "'") { if (src[r + 1] === "'") { q += "'"; r += 2; } else { r++; break; } }
        else { q += src[r]; r++; }
      }
      toks.push({ t: 'sheet', v: q }); i = r; continue;
    }
    // error literal #DIV/0! — '/' and '!' are part of the code, not delimiters
    if (c === '#') {
      var e = i;
      while (e < n && /[#A-Za-z0-9/!?]/.test(src[e])) e++;
      toks.push({ t: 'err', v: src.slice(i, e).toUpperCase() }); i = e; continue;
    }
    // identifier (function name, cell ref, TRUE/FALSE, sheet name)
    if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c === '_' || c === '$') {
      var w = i;
      while (w < n) {
        var d = src[w];
        if ((d >= 'A' && d <= 'Z') || (d >= 'a' && d <= 'z') || (d >= '0' && d <= '9') || d === '_' || d === '$' || d === '.') w++;
        else break;
      }
      toks.push({ t: 'ident', v: src.slice(i, w) }); i = w; continue;
    }
    var two = src.substr(i, 2);
    if (two === '<>' || two === '<=' || two === '>=') { toks.push({ t: 'op', v: two }); i += 2; continue; }
    if (c === '{') { toks.push({ t: 'bl', v: c }); i++; continue; }
    if (c === '}') { toks.push({ t: 'br', v: c }); i++; continue; }
    if (c === ';') { toks.push({ t: 'semi', v: c }); i++; continue; }
    if ('+-*/^%=<>&(),:!'.indexOf(c) >= 0) {
      toks.push({ t: (c === ',' ? 'comma' : c === ':' ? 'colon' : c === '!' ? 'bang' : c === '(' ? 'lp' : c === ')' ? 'rp' : 'op'), v: c });
      i++; continue;
    }
    throw { parse: true, msg: 'Unexpected character: ' + c };
  }
  toks.push({ t: 'eof' });
  return toks;
}

/* ================================ parser =============================== */
// Grammar (lowest to highest precedence):
//   comparison (= <> < > <= >=) -> concat (&) -> add (+ -) -> mul (* /)
//   -> pow (^, right-assoc) -> unary (+ -) -> postfix (%) -> primary
function Parser(toks) { this.toks = toks; this.i = 0; }
Parser.prototype.peek = function (n) { return this.toks[this.i + (n || 0)]; };
Parser.prototype.next = function () { return this.toks[this.i++]; };
Parser.prototype.accept = function (t, v) {
  var k = this.peek();
  if (k.t === t && (v === undefined || k.v === v)) { this.i++; return k; }
  return null;
};
Parser.prototype.expect = function (t, v) {
  var k = this.accept(t, v);
  if (!k) throw { parse: true, msg: 'Expected ' + (v || t) };
  return k;
};

function parseRefNode(P, sheetName, refText) {
  var s = splitRef(refText);
  if (!s) throw { parse: true, msg: 'Invalid reference: ' + refText };
  var c = colToIndex(s.col), r = s.row - 1;
  if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return { t: 'err', v: '#REF!' };
  if (P.accept('colon')) {
    var t2 = P.expect('ident');
    if (P.peek().t === 'bang') throw { parse: true, msg: 'Bad range' };
    var s2 = splitRef(t2.v);
    if (!s2) throw { parse: true, msg: 'Invalid reference: ' + t2.v };
    var c2 = colToIndex(s2.col), r2 = s2.row - 1;
    if (c2 < 0 || c2 >= MAX_COLS || r2 < 0 || r2 >= MAX_ROWS) return { t: 'err', v: '#REF!' };
    return {
      t: 'range', sheet: sheetName,
      c1: Math.min(c, c2), r1: Math.min(r, r2), c2: Math.max(c, c2), r2: Math.max(r, r2),
      absC1: s.absC, absR1: s.absR, absC2: s2.absC, absR2: s2.absR
    };
  }
  return { t: 'cell', sheet: sheetName, c: c, r: r, absC: s.absC, absR: s.absR };
}

function parseAtom(P) {
  var tk = P.peek();
  if (tk.t === 'num') { P.next(); return { t: 'num', v: tk.v }; }
  if (tk.t === 'str') { P.next(); return { t: 'str', v: tk.v }; }
  if (tk.t === 'err') { P.next(); return { t: 'err', v: tk.v }; }
  if (tk.t === 'lp') {
    P.next();
    var e = parseComparison(P);
    P.expect('rp');
    return e;
  }
  if (tk.t === 'bl') { // array literal {1,2;3,4} — comma: same row, semicolon: next row
    P.next();
    var rows = [];
    if (P.peek().t !== 'br') {
      for (;;) {
        var ar = [];
        do { ar.push(parseComparison(P)); } while (P.accept('comma'));
        rows.push(ar);
        if (!P.accept('semi')) break;
      }
      P.expect('br');
    } else { P.next(); }
    if (!rows.length) throw { parse: true, msg: 'Empty array literal' };
    var aw = rows[0].length;
    for (var zi = 0; zi < rows.length; zi++)
      if (rows[zi].length !== aw) throw { parse: true, msg: 'Ragged array literal' };
    return { t: 'arraylit', rows: rows };
  }
  if (tk.t === 'sheet') { // 'My Sheet'!A1
    P.next(); P.expect('bang');
    var refTok = P.expect('ident');
    return parseRefNode(P, tk.v, refTok.v);
  }
  if (tk.t === 'ident') {
    var name = tk.v;
    if (P.peek(1).t === 'bang') { // Sheet2!A1
      P.next(); P.next();
      var rt = P.expect('ident');
      return parseRefNode(P, name, rt.v);
    }
    if (P.peek(1).t === 'lp') { // function call
      P.next(); P.next();
      var args = [];
      if (P.peek().t !== 'rp') { do { args.push(parseComparison(P)); } while (P.accept('comma')); }
      P.expect('rp');
      return { t: 'call', name: name.toUpperCase(), args: args };
    }
    if (/^(TRUE|FALSE)$/i.test(name)) { P.next(); return { t: 'bool', v: /^TRUE/i.test(name) }; }
    if (splitRef(name)) { P.next(); return parseRefNode(P, null, name); }
    P.next();
    return { t: 'name', v: name.toUpperCase() }; // unknown name -> #NAME? at eval
  }
  throw { parse: true, msg: 'Unexpected token in formula' };
}
// Immediate invocation: LAMBDA(...)(args). A trailing "(...)" applies the
// value produced by the atom as a LAMBDA value.
function parsePrimary(P) {
  var node = parseAtom(P);
  while (P.peek().t === 'lp') {
    P.next();
    var args = [];
    if (P.peek().t !== 'rp') { do { args.push(parseComparison(P)); } while (P.accept('comma')); }
    P.expect('rp');
    node = { t: 'apply', fn: node, args: args };
  }
  return node;
}
function parsePostfix(P) {
  var n = parsePrimary(P);
  while (P.accept('op', '%')) n = { t: 'pct', x: n };
  return n;
}
function parseUnary(P) {
  // Excel: unary minus binds looser than ^ (so -3^2 == -(3^2) == -9)
  if (P.accept('op', '-')) return { t: 'neg', x: parseUnary(P) };
  if (P.accept('op', '+')) return parseUnary(P);
  return parsePow(P);
}
function parsePow(P) {
  var b = parsePostfix(P);
  if (P.accept('op', '^')) { var e = parseUnary(P); return { t: 'bin', op: '^', l: b, r: e }; }
  return b;
}
function parseMul(P) {
  var l = parseUnary(P);
  for (;;) {
    var k = P.peek();
    if (k.t === 'op' && (k.v === '*' || k.v === '/')) { P.next(); l = { t: 'bin', op: k.v, l: l, r: parseUnary(P) }; }
    else return l;
  }
}
function parseAdd(P) {
  var l = parseMul(P);
  for (;;) {
    var k = P.peek();
    if (k.t === 'op' && (k.v === '+' || k.v === '-')) { P.next(); l = { t: 'bin', op: k.v, l: l, r: parseMul(P) }; }
    else return l;
  }
}
function parseConcat(P) {
  var l = parseAdd(P);
  while (P.accept('op', '&')) l = { t: 'bin', op: '&', l: l, r: parseAdd(P) };
  return l;
}
function parseComparison(P) {
  var l = parseConcat(P);
  for (;;) {
    var k = P.peek();
    if (k.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].indexOf(k.v) >= 0) {
      P.next(); l = { t: 'bin', op: k.v, l: l, r: parseConcat(P) };
    } else return l;
  }
}
function parse(src) {
  var P = new Parser(tokenize(src));
  var e = parseComparison(P);
  if (P.peek().t !== 'eof') throw { parse: true, msg: 'Unexpected trailing input' };
  return e;
}

/* ============================== AST utilities ========================== */
function walkAst(n, f) {
  if (!n || typeof n !== 'object') return;
  f(n);
  if (n.t === 'bin') { walkAst(n.l, f); walkAst(n.r, f); }
  else if (n.t === 'neg' || n.t === 'pct') walkAst(n.x, f);
  else if (n.t === 'apply') { walkAst(n.fn, f); for (var j = 0; j < n.args.length; j++) walkAst(n.args[j], f); }
  else if (n.t === 'arraylit') { for (var q = 0; q < n.rows.length; q++) for (var w = 0; w < n.rows[q].length; w++) walkAst(n.rows[q][w], f); }
  else if (n.t === 'call') { for (var i = 0; i < n.args.length; i++) walkAst(n.args[i], f); }
}
// Dependency keys ("Sheet1!A1") an AST reads. Needs the workbook to normalize
// sheet-name case; unknown sheets are kept as typed (-> #REF! at eval).
function collectDeps(wb, ownerSheet, ast) {
  var deps = {};
  function add(sh, c, r) {
    var s = wb.getSheet(sh);
    deps[(s ? s.name : sh) + '!' + indexToCol(c) + (r + 1)] = 1;
  }
  walkAst(ast, function (n) {
    var sh, c, r;
    if (n.t === 'cell') add(n.sheet || ownerSheet, n.c, n.r);
    else if (n.t === 'range') {
      sh = n.sheet || ownerSheet;
      for (r = n.r1; r <= n.r2; r++) for (c = n.c1; c <= n.c2; c++) add(sh, c, r);
    }
  });
  return deps;
}
function hasVolatile(ast) {
  var found = false;
  walkAst(ast, function (n) {
    if (n.t === 'call' && FUNCTIONS[n.name] && FUNCTIONS[n.name].volatile) found = true;
  });
  return found;
}
// Shift relative references by (dc, dr); absolute ($) parts stay. Used by fill.
function shiftAst(node, dc, dr) {
  if (!node || typeof node !== 'object') return node;
  function shiftCell(c, r, absC, absR) {
    var nc = absC ? c : c + dc, nr = absR ? r : r + dr;
    if (nc < 0 || nc >= MAX_COLS || nr < 0 || nr >= MAX_ROWS) return null;
    return { c: nc, r: nr };
  }
  if (node.t === 'cell') {
    var p = shiftCell(node.c, node.r, node.absC, node.absR);
    if (!p) return { t: 'err', v: '#REF!' };
    return { t: 'cell', sheet: node.sheet, c: p.c, r: p.r, absC: node.absC, absR: node.absR };
  }
  if (node.t === 'range') {
    var p1 = shiftCell(node.c1, node.r1, node.absC1, node.absR1);
    var p2 = shiftCell(node.c2, node.r2, node.absC2, node.absR2);
    if (!p1 || !p2) return { t: 'err', v: '#REF!' };
    return {
      t: 'range', sheet: node.sheet,
      c1: Math.min(p1.c, p2.c), r1: Math.min(p1.r, p2.r),
      c2: Math.max(p1.c, p2.c), r2: Math.max(p1.r, p2.r),
      absC1: node.absC1, absR1: node.absR1, absC2: node.absC2, absR2: node.absR2
    };
  }
  if (node.t === 'bin') return { t: 'bin', op: node.op, l: shiftAst(node.l, dc, dr), r: shiftAst(node.r, dc, dr) };
  if (node.t === 'neg') return { t: 'neg', x: shiftAst(node.x, dc, dr) };
  if (node.t === 'pct') return { t: 'pct', x: shiftAst(node.x, dc, dr) };
  if (node.t === 'call') {
    return { t: 'call', name: node.name, args: node.args.map(function (a) { return shiftAst(a, dc, dr); }) };
  }
  if (node.t === 'arraylit') {
    return { t: 'arraylit', rows: node.rows.map(function (r) { return r.map(function (x) { return shiftAst(x, dc, dr); }); }) };
  }
  return node; // num/str/bool/err/name are position-independent
}
function quoteSheet(s) {
  return /[^A-Za-z0-9_]/.test(s) ? "'" + s.replace(/'/g, "''") + "'" : s;
}
function cellToText(sheet, c, r, absC, absR) {
  return (sheet ? quoteSheet(sheet) + '!' : '') + (absC ? '$' : '') + indexToCol(c) + (absR ? '$' : '') + (r + 1);
}
function astToText(n) {
  if (!n) return '';
  switch (n.t) {
    case 'num': return numToString(n.v);
    case 'str': return '"' + n.v.replace(/"/g, '""') + '"';
    case 'bool': return n.v ? 'TRUE' : 'FALSE';
    case 'err': return n.v;
    case 'name': return n.v;
    case 'cell': return cellToText(n.sheet, n.c, n.r, n.absC, n.absR);
    case 'range':
      return cellToText(n.sheet, n.c1, n.r1, n.absC1, n.absR1) + ':' +
             cellToText(null, n.c2, n.r2, n.absC2, n.absR2);
    case 'bin': return '(' + astToText(n.l) + n.op + astToText(n.r) + ')';
    case 'neg': return '(-' + astToText(n.x) + ')';
    case 'pct': return '(' + astToText(n.x) + '%)';
    case 'call': return n.name + '(' + n.args.map(astToText).join(',') + ')';
    case 'arraylit': return '{' + n.rows.map(function (r) { return r.map(astToText).join(','); }).join(';') + '}';
  }
  return '';
}

/* ============================ value coercion =========================== */
// Excel semantics: blank -> 0 in numeric context; "5" -> 5; TRUE -> 1;
// non-numeric text in arithmetic -> #VALUE!; errors propagate.
function toNum(v) {
  if (isErr(v)) return v;
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    var s = v.trim();
    if (s === '') return sdErr('#VALUE!');
    var n = Number(s);
    return isNaN(n) ? sdErr('#VALUE!') : n;
  }
  return sdErr('#VALUE!'); // ranges etc.
}
function toStr(v) {
  if (isErr(v)) return v;
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numToString(v);
  return sdErr('#VALUE!');
}
function toBool(v) {
  if (isErr(v)) return v;
  if (v == null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') {
    var s = v.trim().toUpperCase();
    if (s === 'TRUE') return true;
    if (s === 'FALSE') return false;
    return sdErr('#VALUE!');
  }
  return sdErr('#VALUE!');
}
// Excel comparison: type order number < text < logical. Within text,
// case-insensitive. Returns negative/0/positive, or an error value.
function cmpValues(a, b) {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (a == null) a = (typeof b === 'string') ? '' : 0;
  if (b == null) b = (typeof a === 'string') ? '' : 0;
  function rank(v) {
    if (typeof v === 'number') return 0;
    if (typeof v === 'string') return 1;
    if (typeof v === 'boolean') return 2;
    return 3;
  }
  var ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (ra === 0) return a - b;
  if (ra === 1) {
    var x = a.toLowerCase(), y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return (a ? 1 : 0) - (b ? 1 : 0);
}
function cmpEq(a, b) { var r = cmpValues(a, b); return isErr(r) ? r : r === 0; }

/* ========================= function registry =========================== */
var FUNCTIONS = {};
function registerFunction(name, def) {
  def = def || {};
  def.name = String(name).toUpperCase();
  FUNCTIONS[def.name] = def;
  return def;
}

/* ---- shared helpers for function implementations ---- */
function flattenArgs(args) { // expand ranges/arrays to their values, in order
  var out = [];
  for (var i = 0; i < args.length; i++) {
    var a = args[i];
    if (isArrayLike(a)) {
      for (var r = 0; r < a.values.length; r++)
        for (var c = 0; c < a.values[r].length; c++) out.push(a.values[r][c]);
    } else out.push(a);
  }
  return out;
}
// Excel aggregate rule: ranges contribute numbers only (text/blank/logical
// ignored, errors propagate); direct args are coerced ("5"->5, TRUE->1).
function collectNums(args) {
  var out = [];
  for (var i = 0; i < args.length; i++) {
    var a = args[i];
    if (isArrayLike(a)) {
      for (var r = 0; r < a.values.length; r++)
        for (var c = 0; c < a.values[r].length; c++) {
          var v = a.values[r][c];
          if (isErr(v)) return v;
          if (typeof v === 'number') out.push(v);
        }
    } else {
      if (isErr(a)) return a;
      if (a == null) continue;
      if (typeof a === 'number') { out.push(a); continue; }
      if (typeof a === 'boolean') { out.push(a ? 1 : 0); continue; }
      var n = toNum(a);
      if (isErr(n)) return n;
      out.push(n);
    }
  }
  return out;
}
function countNonBlank(args) {
  var n = 0, vs = flattenArgs(args);
  for (var i = 0; i < vs.length; i++) { var v = vs[i]; if (v != null && v !== '') n++; }
  return n;
}
function asTable(arg) { return isArrayLike(arg) ? arg.values : [[arg]]; }
function asVector(arg) { // 1-D vector from a range/array (column preferred) or scalar
  if (isArrayLike(arg)) {
    var vs = arg.values, rows = vs.length, cols = vs[0].length;
    if (rows > 1 && cols > 1) return null;
    var out = [];
    if (rows > 1) { for (var r = 0; r < rows; r++) out.push(vs[r][0]); }
    else { for (var c = 0; c < cols; c++) out.push(vs[0][c]); }
    return out;
  }
  return [arg];
}

/* ---- criteria ("=5", ">3", "a*", "<>") for SUMIF/COUNTIF/AVERAGEIF ---- */
function parseCriteria(crit) {
  if (isErr(crit)) return crit;
  if (typeof crit === 'number') return { op: '=', operand: crit, isNum: true, hasWild: false };
  var s = String(crit == null ? '' : crit);
  var m = /^(<=|>=|<>|<|>|=)([\s\S]*)$/.exec(s), op, operand;
  if (m) { op = m[1]; operand = m[2]; } else { op = '='; operand = s; }
  var num = operand === '' ? NaN : Number(operand);
  return { op: op, operand: operand, num: num, isNum: operand !== '' && !isNaN(num), hasWild: /[*?]/.test(operand) };
}
function wildcardRe(pat) {
  return new RegExp('^' + String(pat).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
}
function matchCriteria(cellVal, cr) {
  if (isErr(cellVal)) return false;
  if (cr.op === '=' && cr.operand === '') return cellVal == null || cellVal === '';
  if (cr.op === '<>' && cr.operand === '') return !(cellVal == null || cellVal === '');
  if (cr.hasWild) {
    var ok = wildcardRe(cr.operand).test(cellVal == null ? '' : String(cellVal));
    return cr.op === '=' ? ok : cr.op === '<>' ? !ok : false;
  }
  var v = (cellVal == null || cellVal === '') ? (cr.isNum ? 0 : '') : cellVal;
  var r;
  if (cr.isNum && typeof v === 'number') r = v - cr.num;
  else if (!cr.isNum) {
    var x = String(v).toLowerCase(), y = String(cr.operand).toLowerCase();
    r = x < y ? -1 : x > y ? 1 : 0;
    if (typeof v === 'number' && cr.operand !== '') { r = -1; } // number < text
  }
  else r = (typeof v === 'string') ? 1 : (v - cr.num); // text > number
  switch (cr.op) {
    case '=': return r === 0;
    case '<>': return r !== 0;
    case '<': return r < 0;
    case '>': return r > 0;
    case '<=': return r <= 0;
    case '>=': return r >= 0;
  }
  return false;
}

/* ---- TEXT(value, format) formatting ---- */
var MONTHS_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function pad2(n) { return (n < 10 ? '0' : '') + n; }
function formatNumberCode(v, code) {
  var pct = code.indexOf('%') >= 0;
  var decMatch = /\.(0+)/.exec(code);
  var dec = decMatch ? decMatch[1].length : 0;
  var comma = code.indexOf(',') >= 0;
  var prefix = /^[^0#?.,%]+/.exec(code);
  var scaled = v * (pct ? 100 : 1);
  var fixed = scaled.toFixed(dec);
  var parts = fixed.split('.'), ip = parts[0], neg = '';
  if (ip[0] === '-') { neg = '-'; ip = ip.slice(1); }
  if (comma) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (prefix ? prefix[0] : '') + neg + ip + (dec ? '.' + parts[1] : '') + (pct ? '%' : '');
}
function formatDateCode(v, code) {
  var d = serialToDate(Math.trunc(v));
  var M = MONTHS_ABBR[d.m - 1];
  switch (code) {
    case 'YYYY-MM-DD': return d.y + '-' + pad2(d.m) + '-' + pad2(d.d);
    case 'YYYY/MM/DD': return d.y + '/' + pad2(d.m) + '/' + pad2(d.d);
    case 'MM/DD/YYYY': return pad2(d.m) + '/' + pad2(d.d) + '/' + d.y;
    case 'M/D/YYYY': return d.m + '/' + d.d + '/' + d.y;
    case 'DD/MM/YYYY': return pad2(d.d) + '/' + pad2(d.m) + '/' + d.y;
    case 'DD-MMM-YYYY': return pad2(d.d) + '-' + M + '-' + d.y;
    case 'MMM D, YYYY': return M + ' ' + d.d + ', ' + d.y;
  }
  return null;
}
function formatText(v, code) {
  if (isErr(v)) return v;
  code = String(code == null ? 'General' : code).toUpperCase();
  if (code === 'GENERAL' || code === 'G/GENERAL') {
    if (v == null) return '';
    return typeof v === 'number' ? numToString(v) : toStr(v);
  }
  if (typeof v === 'number') {
    var dfmt = formatDateCode(v, code);
    if (dfmt !== null) return dfmt;
    if (/^[ $€£¥]?[#0,.]+%?$/.test(code.replace(/ /g, ''))) return formatNumberCode(v, code.replace(/ /g, ''));
  }
  return toStr(v);
}

/* ============================ math functions =========================== */
registerFunction('SUM', { minArgs: 1, maxArgs: null, desc: 'Sum of numbers; text in ranges ignored.',
  fn: function (a) { var s = collectNums(a); if (isErr(s)) return s; var t = 0; for (var i = 0; i < s.length; i++) t += s[i]; return numRes(t); } });
registerFunction('AVERAGE', { minArgs: 1, maxArgs: null, desc: 'Arithmetic mean; #DIV/0! when no numbers.',
  fn: function (a) { var s = collectNums(a); if (isErr(s)) return s; if (!s.length) return sdErr('#DIV/0!'); var t = 0; for (var i = 0; i < s.length; i++) t += s[i]; return numRes(t / s.length); } });
registerFunction('MIN', { minArgs: 1, maxArgs: null, desc: 'Smallest number; 0 when none.',
  fn: function (a) { var s = collectNums(a); if (isErr(s)) return s; return s.length ? Math.min.apply(null, s) : 0; } });
registerFunction('MAX', { minArgs: 1, maxArgs: null, desc: 'Largest number; 0 when none.',
  fn: function (a) { var s = collectNums(a); if (isErr(s)) return s; return s.length ? Math.max.apply(null, s) : 0; } });
registerFunction('COUNT', { minArgs: 1, maxArgs: null, desc: 'Count of numeric values.',
  fn: function (a) { var s = collectNums(a); if (isErr(s)) return s; return s.length; } });
registerFunction('COUNTA', { minArgs: 1, maxArgs: null, desc: 'Count of non-blank values.',
  fn: function (a) { return countNonBlank(a); } });
registerFunction('PRODUCT', { minArgs: 1, maxArgs: null, desc: 'Product of numbers.',
  fn: function (a) { var s = collectNums(a); if (isErr(s)) return s; if (!s.length) return 0; var t = 1; for (var i = 0; i < s.length; i++) t *= s[i]; return numRes(t); } });
registerFunction('ABS', { minArgs: 1, maxArgs: 1, desc: 'Absolute value.',
  fn: function (a, C) { var x = C.num(a[0]); return isErr(x) ? x : Math.abs(x); } });
function roundHalfAway(x, n) {
  var f = Math.pow(10, n);
  return (x < 0 ? -1 : 1) * Math.round(Math.abs(x) * f) / f;
}
registerFunction('ROUND', { minArgs: 1, maxArgs: 2, desc: 'Round half away from zero.',
  fn: function (a, C) { var x = C.num(a[0]); if (isErr(x)) return x; var n = a.length > 1 ? C.num(a[1]) : 0; if (isErr(n)) return n; return numRes(roundHalfAway(x, Math.trunc(n))); } });
registerFunction('ROUNDUP', { minArgs: 1, maxArgs: 2, desc: 'Round away from zero.',
  fn: function (a, C) { var x = C.num(a[0]); if (isErr(x)) return x; var n = a.length > 1 ? C.num(a[1]) : 0; if (isErr(n)) return n; n = Math.trunc(n); var f = Math.pow(10, n); return numRes((x < 0 ? -1 : 1) * Math.ceil(Math.abs(x) * f) / f); } });
registerFunction('ROUNDDOWN', { minArgs: 1, maxArgs: 2, desc: 'Round toward zero.',
  fn: function (a, C) { var x = C.num(a[0]); if (isErr(x)) return x; var n = a.length > 1 ? C.num(a[1]) : 0; if (isErr(n)) return n; n = Math.trunc(n); var f = Math.pow(10, n); return numRes((x < 0 ? -1 : 1) * Math.floor(Math.abs(x) * f) / f); } });
registerFunction('INT', { minArgs: 1, maxArgs: 1, desc: 'Round down to integer.',
  fn: function (a, C) { var x = C.num(a[0]); return isErr(x) ? x : Math.floor(x); } });
registerFunction('TRUNC', { minArgs: 1, maxArgs: 2, desc: 'Truncate toward zero.',
  fn: function (a, C) { var x = C.num(a[0]); if (isErr(x)) return x; var n = a.length > 1 ? C.num(a[1]) : 0; if (isErr(n)) return n; n = Math.trunc(n); var f = Math.pow(10, n); return numRes(Math.trunc(x * f) / f); } });
registerFunction('MOD', { minArgs: 2, maxArgs: 2, desc: 'Remainder with divisor sign (Excel MOD).',
  fn: function (a, C) { var x = C.num(a[0]); if (isErr(x)) return x; var d = C.num(a[1]); if (isErr(d)) return d; if (d === 0) return C.err('#DIV/0!'); return numRes(x - d * Math.floor(x / d)); } });
registerFunction('POWER', { minArgs: 2, maxArgs: 2, desc: 'x raised to y.',
  fn: function (a, C) { var x = C.num(a[0]); if (isErr(x)) return x; var y = C.num(a[1]); if (isErr(y)) return y; var r = Math.pow(x, y); return isNaN(r) ? C.err('#NUM!') : numRes(r); } });
registerFunction('SQRT', { minArgs: 1, maxArgs: 1, desc: 'Square root; #NUM! for negatives.',
  fn: function (a, C) { var x = C.num(a[0]); if (isErr(x)) return x; if (x < 0) return C.err('#NUM!'); return Math.sqrt(x); } });
registerFunction('PI', { minArgs: 0, maxArgs: 0, desc: 'The constant pi.', fn: function () { return Math.PI; } });
registerFunction('RAND', { minArgs: 0, maxArgs: 0, volatile: true, desc: 'Random [0,1); recalculates always.', fn: function () { return Math.random(); } });
registerFunction('RANDBETWEEN', { minArgs: 2, maxArgs: 2, volatile: true, desc: 'Random integer in [bottom, top].',
  fn: function (a, C) { var x = C.num(a[0]); if (isErr(x)) return x; var y = C.num(a[1]); if (isErr(y)) return y; x = Math.trunc(x); y = Math.trunc(y); if (x > y) return C.err('#NUM!'); return Math.floor(Math.random() * (y - x + 1)) + x; } });

/* ---- conditional aggregates ---- */
registerFunction('SUMIF', { minArgs: 2, maxArgs: 3, desc: 'Sum of sum_range where range matches criteria.',
  fn: function (a, C) {
    // a[0] = criteria range, a[1] = criteria (scalar), a[2] = optional sum range
    if (a[1] && a[1].__range) return C.err('#VALUE!');
    var range = asTable(a[0]), sumR = a.length > 2 ? asTable(a[2]) : range;
    var crit = parseCriteria(a[1]);
    var total = 0;
    for (var r = 0; r < range.length; r++)
      for (var c = 0; c < range[r].length; c++) {
        if (matchCriteria(range[r][c], crit)) {
          var v = (sumR[r] && sumR[r][c] != null) ? sumR[r][c] : null;
          if (isErr(v)) return v;
          if (typeof v === 'number') total += v;
        }
      }
    return numRes(total);
  } });
registerFunction('COUNTIF', { minArgs: 2, maxArgs: 2, desc: 'Count of cells matching criteria.',
  fn: function (a) {
    var range = asTable(a[0]);
    var crit = parseCriteria(a[1] && a[1].__range ? sdErr('#VALUE!') : a[1]);
    if (isErr(crit)) return crit;
    var n = 0;
    for (var r = 0; r < range.length; r++)
      for (var c = 0; c < range[r].length; c++)
        if (matchCriteria(range[r][c], crit)) n++;
    return n;
  } });
registerFunction('AVERAGEIF', { minArgs: 2, maxArgs: 3, desc: 'Mean of avg_range where range matches criteria.',
  fn: function (a) {
    var range = asTable(a[0]), avgR = a.length > 2 ? asTable(a[2]) : range;
    var crit = parseCriteria(a[1] && a[1].__range ? sdErr('#VALUE!') : a[1]);
    if (isErr(crit)) return crit;
    var total = 0, n = 0;
    for (var r = 0; r < range.length; r++)
      for (var c = 0; c < range[r].length; c++) {
        if (matchCriteria(range[r][c], crit)) {
          var v = (avgR[r] && avgR[r][c] != null) ? avgR[r][c] : null;
          if (isErr(v)) return v;
          if (typeof v === 'number') { total += v; n++; }
        }
      }
    return n ? numRes(total / n) : sdErr('#DIV/0!');
  } });

/* ============================ logic functions ========================== */
registerFunction('IF', { minArgs: 2, maxArgs: 3, lazy: true, desc: 'Branch on condition (lifts over arrays).',
  fn: function (a, C) {
    var c = a[0]();
    if (isErr(c)) return c;
    if (isArrayLike(c)) {
      var t = a[1](); if (isErr(t)) return t;
      var f = a.length > 2 ? a[2]() : false; if (isErr(f)) return f;
      var A = to2D(c), B = to2D(t), D = to2D(f);
      var sh = broadcastDims([A, B, D]); if (isErr(sh)) return sh;
      var out = [];
      for (var i = 0; i < sh.rows; i++) {
        var row = [];
        for (var j = 0; j < sh.cols; j++) {
          var bv = C.bool(at2D(A, i, j));
          row.push(isErr(bv) ? bv : (bv ? at2D(B, i, j) : at2D(D, i, j)));
        }
        out.push(row);
      }
      return { __array: true, rows: sh.rows, cols: sh.cols, values: out };
    }
    var b = C.bool(c);
    if (isErr(b)) return b;
    if (b) return a[1]();
    return a.length > 2 ? a[2]() : false;
  } });
registerFunction('AND', { minArgs: 1, maxArgs: null, lazy: true, desc: 'TRUE if all args true (short-circuits).',
  fn: function (a, C) {
    for (var i = 0; i < a.length; i++) { var v = C.bool(a[i]()); if (isErr(v)) return v; if (!v) return false; }
    return true;
  } });
registerFunction('OR', { minArgs: 1, maxArgs: null, lazy: true, desc: 'TRUE if any arg true (short-circuits).',
  fn: function (a, C) {
    for (var i = 0; i < a.length; i++) { var v = C.bool(a[i]()); if (isErr(v)) return v; if (v) return true; }
    return false;
  } });
registerFunction('NOT', { minArgs: 1, maxArgs: 1, desc: 'Logical negation.',
  fn: function (a, C) { var v = C.bool(a[0]); return isErr(v) ? v : !v; } });
registerFunction('XOR', { minArgs: 1, maxArgs: null, desc: 'TRUE when an odd number of args are true.',
  fn: function (a, C) {
    var n = 0;
    for (var i = 0; i < a.length; i++) { var v = C.bool(a[i]); if (isErr(v)) return v; if (v) n++; }
    return n % 2 === 1;
  } });
registerFunction('IFERROR', { minArgs: 2, maxArgs: 2, lazy: true, desc: 'Fallback value when first arg is an error.',
  fn: function (a) { var v = a[0](); return isErr(v) ? a[1]() : v; } });

/* ============================ text functions =========================== */
registerFunction('CONCAT', { minArgs: 1, maxArgs: 254, desc: 'Concatenate values and ranges.',
  fn: function (a, C) {
    var vs = flattenArgs(a), out = '';
    for (var i = 0; i < vs.length; i++) { var s = C.str(vs[i]); if (isErr(s)) return s; out += s; }
    return out;
  } });
registerFunction('TEXTJOIN', { minArgs: 3, maxArgs: null, desc: 'Join text with delimiter; optionally skip empties.',
  fn: function (a, C) {
    var d = C.str(C.scalar(a[0])); if (isErr(d)) return d;
    var ig = C.bool(C.scalar(a[1])); if (isErr(ig)) return ig;
    var vs = flattenArgs(a.slice(2)), parts = [];
    for (var i = 0; i < vs.length; i++) {
      var s = C.str(vs[i]); if (isErr(s)) return s;
      if (ig && s === '') continue;
      parts.push(s);
    }
    return parts.join(d);
  } });
registerFunction('LEFT', { minArgs: 1, maxArgs: 2, desc: 'First n characters.',
  fn: function (a, C) { var s = C.str(a[0]); if (isErr(s)) return s; var n = a.length > 1 ? C.num(a[1]) : 1; if (isErr(n)) return n; if (n < 0) return C.err('#VALUE!'); return s.substring(0, Math.trunc(n)); } });
registerFunction('RIGHT', { minArgs: 1, maxArgs: 2, desc: 'Last n characters.',
  fn: function (a, C) { var s = C.str(a[0]); if (isErr(s)) return s; var n = a.length > 1 ? C.num(a[1]) : 1; if (isErr(n)) return n; if (n < 0) return C.err('#VALUE!'); n = Math.trunc(n); return s.substring(Math.max(0, s.length - n)); } });
registerFunction('MID', { minArgs: 3, maxArgs: 3, desc: 'Characters from 1-based start.',
  fn: function (a, C) { var s = C.str(a[0]); if (isErr(s)) return s; var st = C.num(a[1]); if (isErr(st)) return st; var n = C.num(a[2]); if (isErr(n)) return n; st = Math.trunc(st); n = Math.trunc(n); if (st < 1 || n < 0) return C.err('#VALUE!'); return s.substr(st - 1, n); } });
registerFunction('LEN', { minArgs: 1, maxArgs: 1, desc: 'Character count.',
  fn: function (a, C) { var s = C.str(a[0]); return isErr(s) ? s : s.length; } });
registerFunction('TRIM', { minArgs: 1, maxArgs: 1, desc: 'Strip leading/trailing spaces; collapse inner runs.',
  fn: function (a, C) { var s = C.str(a[0]); return isErr(s) ? s : s.replace(/\s+/g, ' ').replace(/^ | $/g, ''); } });
registerFunction('UPPER', { minArgs: 1, maxArgs: 1, desc: 'Uppercase.',
  fn: function (a, C) { var s = C.str(a[0]); return isErr(s) ? s : s.toUpperCase(); } });
registerFunction('LOWER', { minArgs: 1, maxArgs: 1, desc: 'Lowercase.',
  fn: function (a, C) { var s = C.str(a[0]); return isErr(s) ? s : s.toLowerCase(); } });
registerFunction('SUBSTITUTE', { minArgs: 3, maxArgs: 4, desc: 'Replace old with new; optional instance number.',
  fn: function (a, C) {
    var s = C.str(a[0]); if (isErr(s)) return s;
    var old = C.str(a[1]); if (isErr(old)) return old;
    var nw = C.str(a[2]); if (isErr(nw)) return nw;
    if (old === '') return s;
    if (a.length < 4) return s.split(old).join(nw);
    var inst = Math.trunc(C.num(a[3])); if (isErr(inst)) return inst;
    if (inst < 1) return C.err('#VALUE!');
    var idx = -1, from = 0;
    for (var k = 0; k < inst; k++) { idx = s.indexOf(old, from); if (idx < 0) return s; from = idx + old.length; }
    return s.slice(0, idx) + nw + s.slice(idx + old.length);
  } });
registerFunction('TEXT', { minArgs: 2, maxArgs: 2, desc: 'Format a value as text ("0.00", "yyyy-mm-dd", ...).',
  fn: function (a, C) { var v = a[0]; var code = C.str(a[1]); if (isErr(code)) return code; return formatText(v, code); } });
registerFunction('VALUE', { minArgs: 1, maxArgs: 1, desc: 'Parse text as a number.',
  fn: function (a, C) {
    var s = C.str(a[0]); if (isErr(s)) return s;
    s = s.trim();
    if (s === '') return C.err('#VALUE!');
    if (/%$/.test(s)) { var p = Number(s.slice(0, -1)); return isNaN(p) ? C.err('#VALUE!') : p / 100; }
    var n = Number(s);
    return isNaN(n) ? C.err('#VALUE!') : n;
  } });

/* ========================= date & time functions ======================= */
registerFunction('TODAY', { minArgs: 0, maxArgs: 0, volatile: true, desc: "Today's serial date.",
  fn: function () { var d = new Date(); return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); } });
registerFunction('NOW', { minArgs: 0, maxArgs: 0, volatile: true, desc: 'Now as serial date + time fraction.',
  fn: function () { var d = new Date(); return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + d.getMilliseconds() / 1000) / 86400; } });
registerFunction('DATE', { minArgs: 3, maxArgs: 3, desc: 'Serial for y/m/d (month overflow allowed).',
  fn: function (a, C) {
    var y = C.num(a[0]); if (isErr(y)) return y;
    var m = C.num(a[1]); if (isErr(m)) return m;
    var d = C.num(a[2]); if (isErr(d)) return d;
    y = Math.trunc(y); m = Math.trunc(m); d = Math.trunc(d);
    if (y >= 0 && y <= 29) y += 2000; else if (y >= 30 && y <= 99) y += 1900;
    if (y < 1900 || y > 9999) return C.err('#NUM!');
    y += Math.floor((m - 1) / 12); m = ((m - 1) % 12 + 12) % 12 + 1;
    return dateToSerial(y, m, d); // Date.UTC inside handles day overflow
  } });
function serialParts(v, C) {
  var s = C.num(v); if (isErr(s)) return s;
  if (s < 0) return C.err('#NUM!');
  return serialToDate(Math.trunc(s));
}
registerFunction('YEAR', { minArgs: 1, maxArgs: 1, desc: 'Year of a serial date.',
  fn: function (a, C) { var p = serialParts(a[0], C); return isErr(p) ? p : p.y; } });
registerFunction('MONTH', { minArgs: 1, maxArgs: 1, desc: 'Month of a serial date.',
  fn: function (a, C) { var p = serialParts(a[0], C); return isErr(p) ? p : p.m; } });
registerFunction('DAY', { minArgs: 1, maxArgs: 1, desc: 'Day of a serial date.',
  fn: function (a, C) { var p = serialParts(a[0], C); return isErr(p) ? p : p.d; } });
function addMonths(p, months) {
  var total = (p.m - 1) + months;
  var y = p.y + Math.floor(total / 12), m = ((total % 12) + 12) % 12 + 1;
  return { y: y, m: m, d: Math.min(p.d, daysInMonth(y, m)) };
}
registerFunction('EDATE', { minArgs: 2, maxArgs: 2, desc: 'Serial n months before/after.',
  fn: function (a, C) { var p = serialParts(a[0], C); if (isErr(p)) return p; var m = C.num(a[1]); if (isErr(m)) return m; var q = addMonths(p, Math.trunc(m)); return dateToSerial(q.y, q.m, q.d); } });
registerFunction('EOMONTH', { minArgs: 2, maxArgs: 2, desc: 'Last day of the month n months away.',
  fn: function (a, C) { var p = serialParts(a[0], C); if (isErr(p)) return p; var m = C.num(a[1]); if (isErr(m)) return m; var q = addMonths(p, Math.trunc(m)); return dateToSerial(q.y, q.m, daysInMonth(q.y, q.m)); } });
registerFunction('WEEKDAY', { minArgs: 1, maxArgs: 2, desc: 'Day of week (type 1: Sun=1..Sat=7).',
  fn: function (a, C) {
    var s = C.num(a[0]); if (isErr(s)) return s;
    var t = a.length > 1 ? C.num(a[1]) : 1; if (isErr(t)) return t;
    t = Math.trunc(t); s = Math.trunc(s);
    if (t === 1) return ((s - 1) % 7 + 7) % 7 + 1;
    if (t === 2) return ((s + 5) % 7 + 7) % 7 + 1;
    if (t === 3) return ((s + 5) % 7 + 7) % 7;
    return C.err('#NUM!');
  } });

/* ============================ lookup functions ========================= */
registerFunction('VLOOKUP', { minArgs: 3, maxArgs: 4, desc: 'Vertical lookup; exact unless 4th arg TRUE.',
  fn: function (a, C) {
    var key = C.scalar(a[0]); if (isErr(key)) return key;
    var rows = asTable(a[1]);
    var idx = Math.trunc(C.num(a[2])); if (isErr(idx)) return idx;
    if (idx < 1) return C.err('#VALUE!');
    if (!rows.length || idx > rows[0].length) return C.err('#REF!');
    var approx = a.length > 3 ? C.bool(a[3]) : true; if (isErr(approx)) return approx;
    var best = -1;
    for (var i = 0; i < rows.length; i++) {
      var cell = rows[i][0];
      if (isErr(cell)) continue;
      if (approx) { var r = cmpValues(cell, key); if (isErr(r)) continue; if (r <= 0) best = i; }
      else { var e = cmpEq(cell, key); if (isErr(e)) continue; if (e) { best = i; break; } }
    }
    return best < 0 ? C.err('#N/A') : rows[best][idx - 1];
  } });
registerFunction('HLOOKUP', { minArgs: 3, maxArgs: 4, desc: 'Horizontal lookup; exact unless 4th arg TRUE.',
  fn: function (a, C) {
    var key = C.scalar(a[0]); if (isErr(key)) return key;
    var rows = asTable(a[1]);
    var idx = Math.trunc(C.num(a[2])); if (isErr(idx)) return idx;
    if (idx < 1) return C.err('#VALUE!');
    if (!rows.length || idx > rows.length) return C.err('#REF!');
    var approx = a.length > 3 ? C.bool(a[3]) : true; if (isErr(approx)) return approx;
    var best = -1;
    for (var j = 0; j < rows[0].length; j++) {
      var cell = rows[0][j];
      if (isErr(cell)) continue;
      if (approx) { var r = cmpValues(cell, key); if (isErr(r)) continue; if (r <= 0) best = j; }
      else { var e = cmpEq(cell, key); if (isErr(e)) continue; if (e) { best = j; break; } }
    }
    return best < 0 ? C.err('#N/A') : rows[idx - 1][best];
  } });
registerFunction('MATCH', { minArgs: 2, maxArgs: 3, desc: '1-based position of lookup in a vector.',
  fn: function (a, C) {
    var key = C.scalar(a[0]); if (isErr(key)) return key;
    var vec = asVector(a[1]);
    if (!vec) return C.err('#N/A');
    var type = a.length > 2 ? Math.trunc(C.num(a[2])) : 1; if (isErr(type)) return type;
    var best = -1, bestVal = null;
    for (var i = 0; i < vec.length; i++) {
      var v = vec[i];
      if (isErr(v)) continue;
      if (type === 0) { var e = cmpEq(v, key); if (!isErr(e) && e) return i + 1; }
      else {
        var r = cmpValues(v, key); if (isErr(r)) continue;
        if (type === 1 && r <= 0 && (best < 0 || cmpValues(v, bestVal) > 0)) { best = i; bestVal = v; }
        if (type === -1 && r >= 0 && (best < 0 || cmpValues(v, bestVal) < 0)) { best = i; bestVal = v; }
      }
    }
    return best < 0 ? C.err('#N/A') : best + 1;
  } });
registerFunction('INDEX', { minArgs: 2, maxArgs: 3, desc: 'Value at 1-based row/col of an array.',
  fn: function (a, C) {
    var t = asTable(a[0]);
    var r = Math.trunc(C.num(a[1])); if (isErr(r)) return r;
    var c = a.length > 2 ? Math.trunc(C.num(a[2])) : 1; if (isErr(c)) return c;
    if (r < 1 || c < 1) return C.err('#VALUE!');
    if (r > t.length || c > t[0].length) return C.err('#REF!');
    return t[r - 1][c - 1];
  } });
registerFunction('XLOOKUP', { minArgs: 3, maxArgs: 4, desc: 'Exact match in lookup vector; optional if-not-found.',
  fn: function (a, C) {
    var key = C.scalar(a[0]); if (isErr(key)) return key;
    var lv = asVector(a[1]), rv = asVector(a[2]);
    if (!lv || !rv) return C.err('#VALUE!');
    if (lv.length !== rv.length) return C.err('#VALUE!');
    for (var i = 0; i < lv.length; i++) {
      if (isErr(lv[i])) continue;
      var e = cmpEq(lv[i], key);
      if (!isErr(e) && e) return rv[i];
    }
    return a.length > 3 ? a[3] : C.err('#N/A');
  } });
registerFunction('CHOOSE', { minArgs: 2, maxArgs: 254, lazy: true, desc: 'Return the nth value (only it is evaluated).',
  fn: function (a, C) {
    var n = C.num(a[0]()); if (isErr(n)) return n;
    n = Math.trunc(n);
    if (n < 1 || n >= a.length) return C.err('#VALUE!');
    return a[n]();
  } });

/* ---- reference functions (receive reference descriptors) ---- */
registerFunction('ROW', { minArgs: 0, maxArgs: 1, passRef: true, desc: 'Row number of a reference (or this cell).',
  fn: function (a, C) {
    if (!a.length) return C.cellPos.r + 1;
    var d = a[0];
    return (d && d.__ref) ? d.r + 1 : C.err('#VALUE!');
  } });
registerFunction('COLUMN', { minArgs: 0, maxArgs: 1, passRef: true, desc: 'Column number of a reference (or this cell).',
  fn: function (a, C) {
    if (!a.length) return C.cellPos.c + 1;
    var d = a[0];
    return (d && d.__ref) ? d.c + 1 : C.err('#VALUE!');
  } });
registerFunction('ROWS', { minArgs: 1, maxArgs: 1, passRef: true, desc: 'Row count of a reference.',
  fn: function (a, C) { var d = a[0]; if (d && d.__ref === 'range') return d.r2 - d.r1 + 1; return 1; } });
registerFunction('COLUMNS', { minArgs: 1, maxArgs: 1, passRef: true, desc: 'Column count of a reference.',
  fn: function (a, C) { var d = a[0]; if (d && d.__ref === 'range') return d.c2 - d.c1 + 1; return 1; } });

/* ============================ info functions =========================== */
registerFunction('ISNUMBER', { minArgs: 1, maxArgs: 1, propagateErrors: false, desc: 'TRUE for numbers.',
  fn: function (a) { return typeof a[0] === 'number'; } });
registerFunction('ISTEXT', { minArgs: 1, maxArgs: 1, propagateErrors: false, desc: 'TRUE for text.',
  fn: function (a) { return typeof a[0] === 'string'; } });
registerFunction('ISBLANK', { minArgs: 1, maxArgs: 1, propagateErrors: false, desc: 'TRUE for blank cells.',
  fn: function (a) {
    var v = a[0];
    if (v == null) return true;
    if (isArrayLike(v)) {
      var d = arrayDims(v);
      return d.rows === 1 && d.cols === 1 && v.values[0][0] == null;
    }
    return false;
  } });
registerFunction('ISERROR', { minArgs: 1, maxArgs: 1, propagateErrors: false, desc: 'TRUE for any error value.',
  fn: function (a) { return isErr(a[0]); } });

/* ================== example custom function (extensible!) ============== */
// Copy this pattern to add your own. See the "ADDING A FUNCTION" block at the
// top of this file and the README section "Adding a function".
registerFunction('PROFITMARGIN', { minArgs: 2, maxArgs: 2,
  desc: 'Example custom function: (revenue - cost) / revenue.',
  fn: function (args, C) {
    var revenue = C.num(args[0]); if (isErr(revenue)) return revenue;
    var cost = C.num(args[1]); if (isErr(cost)) return cost;
    if (revenue === 0) return C.err('#DIV/0!');
    return (revenue - cost) / revenue;
  } });

/* ==================== dynamic array text functions ===================== */
function textBeforeAfter(isBefore, a, C) {
  var text = C.str(C.scalar(a[0])); if (isErr(text)) return text;
  var delim = C.str(C.scalar(a[1])); if (isErr(delim)) return delim;
  var inst = a.length > 2 ? Math.trunc(C.num(a[2])) : 1; if (isErr(inst)) return inst;
  var matchMode = a.length > 3 ? Math.trunc(C.num(a[3])) : 0; if (isErr(matchMode)) return matchMode;
  var matchEnd = a.length > 4 ? C.bool(a[4]) : false; if (isErr(matchEnd)) return matchEnd;
  var ifNotFound = a.length > 5 ? a[5] : undefined;
  if (delim === '') return C.err('#VALUE!');
  if (inst === 0) return C.err('#VALUE!');
  if (matchMode !== 0 && matchMode !== 1) return C.err('#VALUE!');
  var t = matchMode === 1 ? text.toLowerCase() : text;
  var d = matchMode === 1 ? delim.toLowerCase() : delim;
  var idxs = [], pos = 0;
  for (;;) {
    var p = t.indexOf(d, pos);
    if (p < 0) break;
    idxs.push(p);
    pos = p + d.length;
  }
  if (matchEnd) idxs.push(text.length); // end of text counts as a delimiter
  function notFound() { return ifNotFound === undefined ? C.err('#N/A') : ifNotFound; }
  if (!idxs.length) return notFound();
  var k;
  if (inst > 0) {
    if (inst > idxs.length) return notFound();
    k = idxs[inst - 1];
  } else {
    if (-inst > idxs.length) return notFound();
    k = idxs[idxs.length + inst];
  }
  return isBefore ? text.slice(0, k) : text.slice(k + delim.length);
}
registerFunction('TEXTBEFORE', { minArgs: 2, maxArgs: 6,
  desc: 'Text before the nth delimiter (negative n = from the end).',
  fn: function (a, C) { return textBeforeAfter(true, a, C); } });
registerFunction('TEXTAFTER', { minArgs: 2, maxArgs: 6,
  desc: 'Text after the nth delimiter (negative n = from the end).',
  fn: function (a, C) { return textBeforeAfter(false, a, C); } });

/* ======================= dynamic array functions ====================== */
registerFunction('FILTER', { minArgs: 2, maxArgs: 3,
  desc: 'Rows (or columns) of array where include is true; #CALC! when empty.',
  fn: function (a, C) {
    var d = arrayDims(isArrayLike(a[0]) ? a[0] : [[a[0]]]);
    var vals = isArrayLike(a[0]) ? arrayValues(a[0]) : [[a[0]]];
    var vec = asVector(a[1]);
    if (!vec) return C.err('#VALUE!');
    var flags = [];
    for (var i = 0; i < vec.length; i++) {
      var b = toBool(vec[i]);
      if (isErr(b)) return b;
      flags.push(!!b);
    }
    var out;
    if (flags.length === d.rows) {
      out = [];
      for (var r = 0; r < d.rows; r++) if (flags[r]) out.push(vals[r].slice());
    } else if (flags.length === d.cols) {
      out = [];
      for (var r2 = 0; r2 < d.rows; r2++) {
        var row = [];
        for (var c = 0; c < d.cols; c++) if (flags[c]) row.push(vals[r2][c]);
        out.push(row);
      }
    } else return C.err('#VALUE!');
    if (!out.length || !out[0].length) {
      if (a.length > 2) return a[2]; // if_empty (array ok -> spills)
      return C.err('#CALC!');
    }
    return { __array: true, rows: out.length, cols: out[0].length, values: out };
  } });
registerFunction('SORT', { minArgs: 1, maxArgs: 4,
  desc: 'Sort rows by a column (by_col=1 sorts columns by a row).',
  fn: function (a, C) {
    var d = arrayDims(isArrayLike(a[0]) ? a[0] : [[a[0]]]);
    var vals = isArrayLike(a[0]) ? arrayValues(a[0]) : [[a[0]]];
    var idx = a.length > 1 ? Math.trunc(C.num(C.scalar(a[1]))) : 1; if (isErr(idx)) return idx;
    var order = a.length > 2 ? Math.trunc(C.num(C.scalar(a[2]))) : 1; if (isErr(order)) return order;
    var byCol = a.length > 3 ? C.bool(a[3]) : false; if (isErr(byCol)) return byCol;
    if (order !== 1 && order !== -1) return C.err('#VALUE!');
    var m = vals.map(function (r) { return r.slice(); });
    if (byCol) m = transpose(m);
    if (idx < 1 || idx > m[0].length) return C.err('#VALUE!');
    var kp = idx - 1;
    var dec = m.map(function (r, i) { return { r: r, i: i }; });
    dec.sort(function (x, y) {
      var cr = sortCmp(x.r[kp], y.r[kp]);
      return cr !== 0 ? order * cr : x.i - y.i;
    });
    var out = dec.map(function (x) { return x.r; });
    if (byCol) out = transpose(out);
    return { __array: true, rows: out.length, cols: out[0].length, values: out };
  } });
registerFunction('SORTBY', { minArgs: 2, maxArgs: 254,
  desc: 'Sort rows by one or more key arrays (multi-key).',
  fn: function (a, C) {
    var d = arrayDims(isArrayLike(a[0]) ? a[0] : [[a[0]]]);
    var vals = isArrayLike(a[0]) ? arrayValues(a[0]) : [[a[0]]];
    var keys = [], i = 1;
    while (i < a.length) {
      var kv = a[i++];
      var kvec = asVector(kv);
      if (!kvec) return C.err('#VALUE!');
      var ord = 1;
      if (i < a.length && !isArrayLike(a[i]) && typeof a[i] === 'number') {
        ord = Math.trunc(a[i++]);
      } else if (i < a.length && !isArrayLike(a[i])) {
        var oo = C.num(a[i]); if (isErr(oo)) return oo;
        ord = Math.trunc(oo); i++;
      }
      if (ord !== 1 && ord !== -1) return C.err('#VALUE!');
      keys.push({ vals: kvec, order: ord });
    }
    for (var k = 0; k < keys.length; k++)
      if (keys[k].vals.length !== d.rows) return C.err('#VALUE!');
    var dec = vals.map(function (r, ix) { return { r: r.slice(), ix: ix }; });
    dec.sort(function (x, y) {
      for (var k2 = 0; k2 < keys.length; k2++) {
        var cr = sortCmp(keys[k2].vals[x.ix], keys[k2].vals[y.ix]);
        if (cr !== 0) return keys[k2].order * cr;
      }
      return x.ix - y.ix;
    });
    var out = dec.map(function (x) { return x.r; });
    return { __array: true, rows: out.length, cols: out[0].length, values: out };
  } });
registerFunction('UNIQUE', { minArgs: 1, maxArgs: 3,
  desc: 'Distinct rows (by_col=1: columns; exactly_once=1: singletons).',
  fn: function (a, C) {
    var vals = isArrayLike(a[0]) ? arrayValues(a[0]) : [[a[0]]];
    var byCol = a.length > 1 ? C.bool(a[1]) : false; if (isErr(byCol)) return byCol;
    var once = a.length > 2 ? C.bool(a[2]) : false; if (isErr(once)) return once;
    var m = byCol ? transpose(vals.map(function (r) { return r.slice(); }))
                  : vals.map(function (r) { return r.slice(); });
    function key(r) {
      return r.map(function (v) {
        if (isErr(v)) return 'e' + v.__err;
        if (v == null || v === '') return 'z';
        if (typeof v === 'number') return 'n' + v;
        if (typeof v === 'string') return 's' + v.toLowerCase();
        return 'b' + (v ? 1 : 0);
      }).join('');
    }
    var counts = {}, order = [];
    m.forEach(function (r) {
      var k = key(r);
      if (!Object.prototype.hasOwnProperty.call(counts, k)) { counts[k] = 0; order.push({ k: k, r: r }); }
      counts[k]++;
    });
    var out = [];
    order.forEach(function (o) { if (!once || counts[o.k] === 1) out.push(o.r.slice()); });
    if (!out.length) return C.err('#CALC!');
    if (byCol) out = transpose(out);
    return { __array: true, rows: out.length, cols: out[0].length, values: out };
  } });
function chooseIdx(n, size, C) {
  if (n === 0) return C.err('#VALUE!');
  var i = n > 0 ? n - 1 : size + n;
  return (i < 0 || i >= size) ? C.err('#VALUE!') : i;
}
registerFunction('CHOOSECOLS', { minArgs: 2, maxArgs: 254,
  desc: 'Pick columns by number (negative = from the end).',
  fn: function (a, C) {
    var d = arrayDims(isArrayLike(a[0]) ? a[0] : [[a[0]]]);
    var vals = isArrayLike(a[0]) ? arrayValues(a[0]) : [[a[0]]];
    var cols = [];
    for (var i = 1; i < a.length; i++) {
      var n = Math.trunc(C.num(C.scalar(a[i]))); if (isErr(n)) return n;
      var ix = chooseIdx(n, d.cols, C); if (isErr(ix)) return ix;
      cols.push(ix);
    }
    var out = vals.map(function (r) { return cols.map(function (c) { return r[c]; }); });
    return { __array: true, rows: out.length, cols: cols.length, values: out };
  } });
registerFunction('CHOOSEROWS', { minArgs: 2, maxArgs: 254,
  desc: 'Pick rows by number (negative = from the end).',
  fn: function (a, C) {
    var d = arrayDims(isArrayLike(a[0]) ? a[0] : [[a[0]]]);
    var vals = isArrayLike(a[0]) ? arrayValues(a[0]) : [[a[0]]];
    var rows = [];
    for (var i = 1; i < a.length; i++) {
      var n = Math.trunc(C.num(C.scalar(a[i]))); if (isErr(n)) return n;
      var ix = chooseIdx(n, d.rows, C); if (isErr(ix)) return ix;
      rows.push(ix);
    }
    var out = rows.map(function (r) { return vals[r].slice(); });
    return { __array: true, rows: out.length, cols: d.cols, values: out };
  } });
registerFunction('DROP', { minArgs: 2, maxArgs: 3,
  desc: 'Drop leading rows/cols (negative = from the end); #CALC! when empty.',
  fn: function (a, C) {
    var d = arrayDims(isArrayLike(a[0]) ? a[0] : [[a[0]]]);
    var vals = isArrayLike(a[0]) ? arrayValues(a[0]) : [[a[0]]];
    var dr = Math.trunc(C.num(C.scalar(a[1]))); if (isErr(dr)) return dr;
    var dc = a.length > 2 ? Math.trunc(C.num(C.scalar(a[2]))) : 0; if (isErr(dc)) return dc;
    var r1 = dr >= 0 ? dr : 0, r2 = dr >= 0 ? d.rows : d.rows + dr;
    var c1 = dc >= 0 ? dc : 0, c2 = dc >= 0 ? d.cols : d.cols + dc;
    if (r2 <= r1 || c2 <= c1) return C.err('#CALC!');
    var out = [];
    for (var r = r1; r < r2; r++) out.push(vals[r].slice(c1, c2));
    return { __array: true, rows: out.length, cols: out[0].length, values: out };
  } });
registerFunction('TAKE', { minArgs: 2, maxArgs: 3,
  desc: 'Take leading rows/cols (negative = from the end).',
  fn: function (a, C) {
    var d = arrayDims(isArrayLike(a[0]) ? a[0] : [[a[0]]]);
    var vals = isArrayLike(a[0]) ? arrayValues(a[0]) : [[a[0]]];
    var tr = Math.trunc(C.num(C.scalar(a[1]))); if (isErr(tr)) return tr;
    var tc = a.length > 2 ? Math.trunc(C.num(C.scalar(a[2]))) : d.cols; if (isErr(tc)) return tc;
    if (tr === 0 || tc === 0) return C.err('#CALC!');
    var r1, r2, c1, c2;
    if (tr > 0) { r1 = 0; r2 = Math.min(tr, d.rows); } else { r1 = Math.max(0, d.rows + tr); r2 = d.rows; }
    if (tc > 0) { c1 = 0; c2 = Math.min(tc, d.cols); } else { c1 = Math.max(0, d.cols + tc); c2 = d.cols; }
    var out = [];
    for (var r = r1; r < r2; r++) out.push(vals[r].slice(c1, c2));
    return { __array: true, rows: out.length, cols: out[0].length, values: out };
  } });
registerFunction('HSTACK', { minArgs: 1, maxArgs: 254,
  desc: 'Append arrays side-by-side; shorter ones padded with #N/A.',
  fn: function (a, C) {
    var arrs = [], rows = 0, cols = 0;
    for (var i = 0; i < a.length; i++) {
      var d = to2D(a[i]);
      arrs.push(d); rows = Math.max(rows, d.rows); cols += d.cols;
    }
    var out = [];
    for (var r = 0; r < rows; r++) {
      var row = [];
      for (var k = 0; k < arrs.length; k++) {
        var d2 = arrs[k];
        for (var c = 0; c < d2.cols; c++)
          row.push(r < d2.rows ? d2.values[r][c] : C.err('#N/A'));
      }
      out.push(row);
    }
    return { __array: true, rows: rows, cols: cols, values: out };
  } });
registerFunction('VSTACK', { minArgs: 1, maxArgs: 254,
  desc: 'Stack arrays top-to-bottom; narrower ones padded with #N/A.',
  fn: function (a, C) {
    var arrs = [], rows = 0, cols = 0;
    for (var i = 0; i < a.length; i++) {
      var d = to2D(a[i]);
      arrs.push(d); cols = Math.max(cols, d.cols); rows += d.rows;
    }
    var out = [];
    for (var k = 0; k < arrs.length; k++) {
      var d2 = arrs[k];
      for (var r = 0; r < d2.rows; r++) {
        var row = d2.values[r].slice();
        while (row.length < cols) row.push(C.err('#N/A'));
        out.push(row);
      }
    }
    return { __array: true, rows: rows, cols: cols, values: out };
  } });

/* ============================ LET and LAMBDA =========================== */
registerFunction('LET', { minArgs: 3, maxArgs: 254, astArgs: true,
  desc: 'Bind name/value pairs in order, then evaluate calc. Bindings are eager and sequential.',
  fn: function (nodes, C) {
    if (nodes.length % 2 === 0) return C.err('#N/A'); // need pairs + calc
    var scope = { parent: C.scope || null, vars: {} };
    var C2 = Object.create(C);
    C2.scope = scope;
    for (var i = 0; i < nodes.length - 1; i += 2) {
      var nm = nodes[i];
      if (!nm || nm.t !== 'name' || splitRef(nm.v)) return C.err('#NAME?');
      var val = C2.eval(nodes[i + 1], C2); // eager, in order: later names see earlier ones
      if (isErr(val)) return val;
      scope.vars[nm.v.toUpperCase()] = { value: val, evaluated: true };
    }
    return C2.eval(nodes[nodes.length - 1], C2);
  } });
registerFunction('LAMBDA', { minArgs: 1, maxArgs: 254, astArgs: true,
  desc: 'Define a function value. Call via LAMBDA(...)(args) or a LET-bound name. Recursion allowed (depth cap 1000).',
  fn: function (nodes, C) {
    var params = [];
    for (var i = 0; i < nodes.length - 1; i++) {
      var p = nodes[i];
      if (!p || p.t !== 'name' || splitRef(p.v)) return C.err('#VALUE!');
      params.push(p.v.toUpperCase());
    }
    return { __lambda: true, params: params, body: nodes[nodes.length - 1], defScope: C.scope || null };
  } });
// NOTE: the MAP/REDUCE/SCAN/MAKEARRAY/BYROW/BYCOL helper family is deliberately
// not implemented yet (see README). Typing one yields #NAME? rather than a
// half-working version.

/* ===================== extended error functions ======================== */
registerFunction('IFNA', { minArgs: 2, maxArgs: 2, lazy: true,
  desc: 'Fallback only for #N/A; all other errors propagate.',
  fn: function (a) { var v = a[0](); return (isErr(v) && v.__err === '#N/A') ? a[1]() : v; } });
registerFunction('ISERR', { minArgs: 1, maxArgs: 1, propagateErrors: false,
  desc: 'TRUE for any error except #N/A.',
  fn: function (a) { return isErr(a[0]) && a[0].__err !== '#N/A'; } });
registerFunction('ISNA', { minArgs: 1, maxArgs: 1, propagateErrors: false,
  desc: 'TRUE for #N/A.',
  fn: function (a) { return isErr(a[0]) && a[0].__err === '#N/A'; } });
registerFunction('ERROR.TYPE', { minArgs: 1, maxArgs: 1, propagateErrors: false,
  desc: '1-7 for #NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A; else #N/A.',
  fn: function (a, C) {
    var v = a[0];
    if (!isErr(v)) return C.err('#N/A');
    switch (v.__err) {
      case '#NULL!': return 1;
      case '#DIV/0!': return 2;
      case '#VALUE!': return 3;
      case '#REF!': return 4;
      case '#NAME?': return 5;
      case '#NUM!': return 6;
      case '#N/A': return 7;
      default: return C.err('#N/A'); // #SPILL!, #CALC!, #CYCLE!, ...
    }
  } });

/* ===================== extended statistical functions ================== */
registerFunction('MEDIAN', { minArgs: 1, maxArgs: null, desc: 'Median; #NUM! when no numbers.',
  fn: function (a) { var s = collectNums(a); return isErr(s) ? s : medianOf(s); } });
function modeFn(a) { var s = collectNums(a); return isErr(s) ? s : modeOf(s); }
registerFunction('MODE', { minArgs: 1, maxArgs: null, desc: 'Most frequent value; #N/A when nothing repeats.', fn: modeFn });
registerFunction('MODE.SNGL', { minArgs: 1, maxArgs: null, desc: 'Alias of MODE.', fn: modeFn });
registerFunction('LARGE', { minArgs: 2, maxArgs: 2, desc: 'k-th largest value.',
  fn: function (a, C) {
    var s = collectNums([a[0]]); if (isErr(s)) return s;
    var k = Math.trunc(C.num(C.scalar(a[1]))); if (isErr(k)) return k;
    return largeOf(s, k);
  } });
registerFunction('SMALL', { minArgs: 2, maxArgs: 2, desc: 'k-th smallest value.',
  fn: function (a, C) {
    var s = collectNums([a[0]]); if (isErr(s)) return s;
    var k = Math.trunc(C.num(C.scalar(a[1]))); if (isErr(k)) return k;
    return smallOf(s, k);
  } });
registerFunction('STDEV.S', { minArgs: 1, maxArgs: null, desc: 'Sample standard deviation.',
  fn: function (a) { var s = collectNums(a); return isErr(s) ? s : numRes(stdevS(s)); } });
registerFunction('STDEV.P', { minArgs: 1, maxArgs: null, desc: 'Population standard deviation.',
  fn: function (a) { var s = collectNums(a); return isErr(s) ? s : numRes(stdevP(s)); } });
registerFunction('VAR.S', { minArgs: 1, maxArgs: null, desc: 'Sample variance.',
  fn: function (a) { var s = collectNums(a); if (isErr(s)) return s; var d = stdevS(s); return isErr(d) ? d : numRes(d * d); } });
registerFunction('VAR.P', { minArgs: 1, maxArgs: null, desc: 'Population variance.',
  fn: function (a) { var s = collectNums(a); if (isErr(s)) return s; var d = stdevP(s); return isErr(d) ? d : numRes(d * d); } });
registerFunction('PERCENTILE.INC', { minArgs: 2, maxArgs: 2, desc: 'k-th percentile, k in [0,1].',
  fn: function (a, C) {
    var s = collectNums([a[0]]); if (isErr(s)) return s;
    var k = C.num(C.scalar(a[1])); if (isErr(k)) return k;
    return percentileInc(s, k);
  } });
registerFunction('PERCENTILE.EXC', { minArgs: 2, maxArgs: 2, desc: 'k-th percentile, k in (0,1).',
  fn: function (a, C) {
    var s = collectNums([a[0]]); if (isErr(s)) return s;
    var k = C.num(C.scalar(a[1])); if (isErr(k)) return k;
    return percentileExc(s, k);
  } });
registerFunction('QUARTILE.INC', { minArgs: 2, maxArgs: 2, desc: 'Quartile 0-4 (inclusive).',
  fn: function (a, C) {
    var s = collectNums([a[0]]); if (isErr(s)) return s;
    var q = Math.trunc(C.num(C.scalar(a[1]))); if (isErr(q)) return q;
    if (q < 0 || q > 4) return C.err('#NUM!');
    return percentileInc(s, q / 4);
  } });
registerFunction('QUARTILE.EXC', { minArgs: 2, maxArgs: 2, desc: 'Quartile 1-3 (exclusive).',
  fn: function (a, C) {
    var s = collectNums([a[0]]); if (isErr(s)) return s;
    var q = Math.trunc(C.num(C.scalar(a[1]))); if (isErr(q)) return q;
    if (q < 1 || q > 3) return C.err('#NUM!');
    return percentileExc(s, q / 4);
  } });
registerFunction('AGGREGATE', { minArgs: 3, maxArgs: 4,
  desc: 'Aggregate with options 0-7 to ignore errors / nested AGGREGATEs. (No hidden rows in v1: options 1,3,5,7 behave like 0,2,4,6.)',
  fn: function (a, C) {
    var fnum = C.num(a[0]); if (isErr(fnum)) return fnum; fnum = Math.trunc(fnum);
    var opt = C.num(a[1]); if (isErr(opt)) return opt; opt = Math.trunc(opt);
    if (fnum < 1 || fnum > 19 || opt < 0 || opt > 7) return C.err('#VALUE!');
    var ignoreErr = (opt === 2 || opt === 3 || opt === 6 || opt === 7);
    var ignoreNested = (opt <= 3);
    // v1 has no hidden rows, so options 1,3,5,7 behave like 0,2,4,6.
    var src = a[2], rawVals = [];
    if (isArrayLike(src)) {
      var d = arrayDims(src), vs = arrayValues(src);
      for (var r = 0; r < d.rows; r++)
        for (var c = 0; c < d.cols; c++) {
          // Per Excel, nested ignores apply to plain references, not arrays
          // produced by calculations.
          if (ignoreNested && src.__range && C.wb.cellHasAggregate(src.sheet, src.c1 + c, src.r1 + r)) continue;
          rawVals.push(vs[r][c]);
        }
    } else rawVals.push(src);
    var vals = [];
    for (var i = 0; i < rawVals.length; i++) {
      var v = rawVals[i];
      if (isErr(v)) { if (ignoreErr) continue; return v; }
      vals.push(v);
    }
    function nums() {
      var out = [];
      for (var i = 0; i < vals.length; i++) if (typeof vals[i] === 'number') out.push(vals[i]);
      return out;
    }
    var k, ns, i2, t, p, n;
    switch (fnum) {
      case 1: ns = nums(); return ns.length ? numRes(meanOf(ns)) : C.err('#DIV/0!');
      case 2: return nums().length;
      case 3: n = 0; for (i2 = 0; i2 < vals.length; i2++) if (vals[i2] != null && vals[i2] !== '') n++; return n;
      case 4: ns = nums(); return ns.length ? Math.max.apply(null, ns) : 0;
      case 5: ns = nums(); return ns.length ? Math.min.apply(null, ns) : 0;
      case 6: ns = nums(); if (!ns.length) return 0; p = 1; for (i2 = 0; i2 < ns.length; i2++) p *= ns[i2]; return numRes(p);
      case 7: return numRes(stdevS(nums()));
      case 8: return numRes(stdevP(nums()));
      case 9: ns = nums(); t = 0; for (i2 = 0; i2 < ns.length; i2++) t += ns[i2]; return numRes(t);
      case 10: { var d7 = stdevS(nums()); return isErr(d7) ? d7 : numRes(d7 * d7); }
      case 11: { var d8 = stdevP(nums()); return isErr(d8) ? d8 : numRes(d8 * d8); }
      case 12: return medianOf(nums());
      case 13: return modeOf(nums());
      case 14: case 15: case 16: case 17: case 18: case 19:
        if (a.length < 4) return C.err('#VALUE!'); // k required
        k = C.num(C.scalar(a[3])); if (isErr(k)) return k;
        ns = nums();
        if (fnum === 14) return largeOf(ns, Math.trunc(k));
        if (fnum === 15) return smallOf(ns, Math.trunc(k));
        if (fnum === 16) return percentileInc(ns, k);
        if (fnum === 17) { var q17 = Math.trunc(k); if (q17 < 0 || q17 > 4) return C.err('#NUM!'); return percentileInc(ns, q17 / 4); }
        if (fnum === 18) return percentileExc(ns, k);
        var q19 = Math.trunc(k); if (q19 < 1 || q19 > 3) return C.err('#NUM!'); return percentileExc(ns, q19 / 4);
    }
    return C.err('#VALUE!');
  } });

/* ====================== dynamic array / spill helpers ================== */
// Array values: { __array: true, rows, cols, values }. Ranges are references
// ({ __range: true, ... }); both are "array-like" for elementwise lifting.
function isArrayLike(v) { return !!(v && (v.__range || v.__array)); }
function arrayValues(v) { return v.values; }
function arrayDims(v) {
  return v.__array ? { rows: v.rows, cols: v.cols }
                   : { rows: v.values.length, cols: v.values[0].length };
}
function to2D(v) {
  if (v && v.__array) return { rows: v.rows, cols: v.cols, values: v.values };
  if (v && v.__range) return { rows: v.values.length, cols: v.values[0].length, values: v.values };
  return { rows: 1, cols: 1, values: [[v]] };
}
function broadcastDims(list) {
  var rows = 1, cols = 1, i;
  for (i = 0; i < list.length; i++) {
    rows = Math.max(rows, list[i].rows);
    cols = Math.max(cols, list[i].cols);
  }
  for (i = 0; i < list.length; i++) {
    if ((list[i].rows !== rows && list[i].rows !== 1) ||
        (list[i].cols !== cols && list[i].cols !== 1)) return sdErr('#VALUE!');
  }
  return { rows: rows, cols: cols };
}
function at2D(x, i, j) { return x.values[i % x.rows][j % x.cols]; }
function mapArray(v, f) {
  var d = to2D(v), out = [];
  for (var i = 0; i < d.rows; i++) {
    var row = [];
    for (var j = 0; j < d.cols; j++) row.push(f(d.values[i][j]));
    out.push(row);
  }
  return { __array: true, rows: d.rows, cols: d.cols, values: out };
}
function transpose(m) {
  var out = [];
  for (var c = 0; c < m[0].length; c++) {
    var row = [];
    for (var r = 0; r < m.length; r++) row.push(m[r][c]);
    out.push(row);
  }
  return out;
}
// Sort comparator: blanks last, errors last, then number < text < logical.
function sortCmp(a, b) {
  var ae = isErr(a), be = isErr(b);
  if (ae && be) return 0;
  if (ae) return 1;
  if (be) return -1;
  var an = (a == null || a === ''), bn = (b == null || b === '');
  if (an && bn) return 0;
  if (an) return 1;
  if (bn) return -1;
  return cmpValues(a, b);
}
// LET/LAMBDA lexical scope lookup. Returns the bound value or undefined.
function scopeLookup(scope, name) {
  var key = String(name).toUpperCase();
  while (scope) {
    if (scope.vars && Object.prototype.hasOwnProperty.call(scope.vars, key))
      return scope.vars[key].value;
    scope = scope.parent;
  }
  return undefined;
}
function applyLambda(lam, argVals, C) {
  if (argVals.length !== lam.params.length) return sdErr('#N/A');
  var depth = (C.callDepth || 0) + 1;
  if (depth > 1000) return sdErr('#CALC!'); // recursion guard
  var scope = { parent: lam.defScope, vars: {} };
  for (var i = 0; i < lam.params.length; i++)
    scope.vars[lam.params[i]] = { value: argVals[i], evaluated: true };
  var C2 = Object.create(C);
  C2.scope = scope;
  C2.callDepth = depth;
  return evalNode(lam.body, C2);
}
// Shared numeric summaries used by standalone fns and AGGREGATE.
function meanOf(nums) { var t = 0; for (var i = 0; i < nums.length; i++) t += nums[i]; return t / nums.length; }
function numsSorted(nums) { return nums.slice().sort(function (x, y) { return x - y; }); }
function medianOf(nums) {
  if (!nums.length) return sdErr('#NUM!');
  var s = numsSorted(nums), n = s.length, m = Math.floor(n / 2);
  return n % 2 ? s[m] : numRes((s[m - 1] + s[m]) / 2);
}
function modeOf(nums) {
  if (!nums.length) return sdErr('#N/A');
  var counts = {}, best = nums[0], bestN = 0;
  for (var i = 0; i < nums.length; i++) {
    var k = 'v' + nums[i];
    counts[k] = (counts[k] || 0) + 1;
    if (counts[k] > bestN) { bestN = counts[k]; best = nums[i]; }
  }
  return bestN > 1 ? best : sdErr('#N/A');
}
function largeOf(nums, k) {
  if (!nums.length || k < 1 || k > nums.length) return sdErr('#NUM!');
  return numsSorted(nums)[nums.length - k];
}
function smallOf(nums, k) {
  if (!nums.length || k < 1 || k > nums.length) return sdErr('#NUM!');
  return numsSorted(nums)[k - 1];
}
function stdevS(nums) {
  if (nums.length < 2) return sdErr('#DIV/0!');
  var m = meanOf(nums), s = 0;
  for (var i = 0; i < nums.length; i++) s += (nums[i] - m) * (nums[i] - m);
  return Math.sqrt(s / (nums.length - 1));
}
function stdevP(nums) {
  if (nums.length < 1) return sdErr('#DIV/0!');
  var m = meanOf(nums), s = 0;
  for (var i = 0; i < nums.length; i++) s += (nums[i] - m) * (nums[i] - m);
  return Math.sqrt(s / nums.length);
}
function percentileInc(nums, k) {
  if (!nums.length) return sdErr('#NUM!');
  if (typeof k !== 'number' || k < 0 || k > 1) return sdErr('#NUM!');
  var s = numsSorted(nums), n = s.length;
  if (n === 1) return s[0];
  var rank = k * (n - 1), lo = Math.floor(rank), hi = Math.ceil(rank);
  return numRes(s[lo] + (s[hi] - s[lo]) * (rank - lo));
}
function percentileExc(nums, k) {
  if (!nums.length) return sdErr('#NUM!');
  if (typeof k !== 'number' || k <= 0 || k >= 1) return sdErr('#NUM!');
  var s = numsSorted(nums), n = s.length, rank = k * (n + 1);
  if (rank < 1 || rank > n) return sdErr('#NUM!');
  var lo = Math.floor(rank) - 1, hi = Math.ceil(rank) - 1;
  return numRes(s[lo] + (s[hi] - s[lo]) * (rank - Math.floor(rank)));
}
// Does a formula AST contain an AGGREGATE/SUBTOTAL call? (for AGGREGATE options)
function astHasAggregateCall(ast) {
  var found = false;
  walkAst(ast, function (n) {
    if (n.t === 'call' && (n.name === 'AGGREGATE' || n.name === 'SUBTOTAL')) found = true;
  });
  return found;
}

/* ============================ AST evaluator ============================ */
// Scalar binary operation (operands are guaranteed non-array here).
function scalarBin(op, L, R) {
  var a, b, r, ls, rs, e;
  switch (op) {
    case '+': a = toNum(L); if (isErr(a)) return a; b = toNum(R); if (isErr(b)) return b; return numRes(a + b);
    case '-': a = toNum(L); if (isErr(a)) return a; b = toNum(R); if (isErr(b)) return b; return numRes(a - b);
    case '*': a = toNum(L); if (isErr(a)) return a; b = toNum(R); if (isErr(b)) return b; return numRes(a * b);
    case '/': a = toNum(L); if (isErr(a)) return a; b = toNum(R); if (isErr(b)) return b; if (b === 0) return sdErr('#DIV/0!'); return numRes(a / b);
    case '^': a = toNum(L); if (isErr(a)) return a; b = toNum(R); if (isErr(b)) return b; var p = Math.pow(a, b); return isNaN(p) ? sdErr('#NUM!') : numRes(p);
    case '&': ls = toStr(L); if (isErr(ls)) return ls; rs = toStr(R); if (isErr(rs)) return rs; return ls + rs;
    case '=': return cmpEq(L, R);
    case '<>': e = cmpEq(L, R); return isErr(e) ? e : !e;
    default:
      r = cmpValues(L, R); if (isErr(r)) return r;
      return op === '<' ? r < 0 : op === '>' ? r > 0 : op === '<=' ? r <= 0 : r >= 0;
  }
}
function evalBin(n, C) {
  var L = evalNode(n.l, C); if (isErr(L)) return L;
  var R = evalNode(n.r, C); if (isErr(R)) return R;
  if (isArrayLike(L) || isArrayLike(R)) {
    // Excel 365 elementwise lifting with scalar/1-row/1-col broadcasting
    var A = to2D(L), B = to2D(R);
    var sh = broadcastDims([A, B]); if (isErr(sh)) return sh;
    var out = [];
    for (var i = 0; i < sh.rows; i++) {
      var row = [];
      for (var j = 0; j < sh.cols; j++) row.push(scalarBin(n.op, at2D(A, i, j), at2D(B, i, j)));
      out.push(row);
    }
    return { __array: true, rows: sh.rows, cols: sh.cols, values: out };
  }
  return scalarBin(n.op, L, R);
}
function refDescriptor(node, C) {
  if (node.t === 'cell') {
    var sh = node.sheet ? C.resolveSheet(node.sheet) : C.sheet;
    if (!sh) return sdErr('#REF!');
    return { __ref: 'cell', sheet: sh, c: node.c, r: node.r };
  }
  if (node.t === 'range') {
    var sh2 = node.sheet ? C.resolveSheet(node.sheet) : C.sheet;
    if (!sh2) return sdErr('#REF!');
    return { __ref: 'range', sheet: sh2, c1: node.c1, r1: node.r1, c2: node.c2, r2: node.r2, c: node.c1, r: node.r1 };
  }
  return evalNode(node, C);
}
function evalCall(n, C) {
  var meta = FUNCTIONS[n.name], i, v;
  if (!meta) {
    // Maybe a LET/LAMBDA-bound lambda used as f(args)
    var lam = scopeLookup(C.scope, n.name);
    if (lam && lam.__lambda) {
      var argVals = [];
      for (i = 0; i < n.args.length; i++) {
        v = evalNode(n.args[i], C);
        if (isErr(v)) return v;
        argVals.push(v);
      }
      return applyLambda(lam, argVals, C);
    }
    return sdErr('#NAME?');
  }
  var argc = n.args.length;
  if (argc < (meta.minArgs || 0) || (meta.maxArgs != null && argc > meta.maxArgs)) return sdErr('#N/A');
  var args = [];
  if (meta.astArgs) {
    args = n.args; // raw AST nodes (LET names, LAMBDA params/body)
  } else if (meta.lazy) {
    args = n.args.map(function (a) { return function () { return evalNode(a, C); }; });
  } else if (meta.passRef) {
    for (i = 0; i < argc; i++) args.push(refDescriptor(n.args[i], C));
  } else {
    for (i = 0; i < argc; i++) args.push(evalNode(n.args[i], C));
    if (meta.propagateErrors !== false) { for (i = 0; i < argc; i++) if (isErr(args[i])) return args[i]; }
  }
  return meta.fn(args, C);
}
function evalNode(n, C) {
  switch (n.t) {
    case 'num': case 'str': case 'bool': return n.v;
    case 'err': return sdErr(n.v);
    case 'name': {
      var sv = scopeLookup(C.scope, n.v); // LET/LAMBDA bindings first
      return sv === undefined ? sdErr('#NAME?') : sv;
    }
    case 'apply': { // immediate invocation: <lambda-valued expr>(args)
      var f = evalNode(n.fn, C);
      if (isErr(f)) return f;
      if (!f || !f.__lambda) return sdErr('#VALUE!');
      var av = [];
      for (var q = 0; q < n.args.length; q++) {
        var vv = evalNode(n.args[q], C);
        if (isErr(vv)) return vv;
        av.push(vv);
      }
      return applyLambda(f, av, C);
    }
    case 'neg': {
      var nx = evalNode(n.x, C);
      if (isErr(nx)) return nx;
      if (isArrayLike(nx)) return mapArray(nx, function (x) { var t = toNum(x); return isErr(t) ? t : numRes(-t); });
      var a = toNum(nx); return isErr(a) ? a : numRes(-a);
    }
    case 'pct': {
      var px = evalNode(n.x, C);
      if (isErr(px)) return px;
      if (isArrayLike(px)) return mapArray(px, function (x) { var t = toNum(x); return isErr(t) ? t : numRes(t / 100); });
      var b = toNum(px); return isErr(b) ? b : numRes(b / 100);
    }
    case 'bin': return evalBin(n, C);
    case 'arraylit': {
      var av2 = [];
      for (var ai = 0; ai < n.rows.length; ai++) {
        var arow = [];
        for (var aj = 0; aj < n.rows[ai].length; aj++) arow.push(evalNode(n.rows[ai][aj], C));
        av2.push(arow);
      }
      return { __array: true, rows: av2.length, cols: av2[0].length, values: av2 };
    }
    case 'cell': {
      var sh = n.sheet ? C.resolveSheet(n.sheet) : C.sheet;
      if (!sh) return sdErr('#REF!');
      return C.wb.readCellValue(sh, n.c, n.r, C);
    }
    case 'range': {
      var sh2 = n.sheet ? C.resolveSheet(n.sheet) : C.sheet;
      if (!sh2) return sdErr('#REF!');
      return {
        __range: true, sheet: sh2, c1: n.c1, r1: n.r1, c2: n.c2, r2: n.r2,
        values: C.wb.readRangeValues(sh2, n.c1, n.r1, n.c2, n.r2, C)
      };
    }
    case 'call': return evalCall(n, C);
  }
  return sdErr('#VALUE!');
}
function makeCtx(wb, sheetName, key, cellPos) {
  var ctx = {
    wb: wb, sheet: sheetName, key: key, cellPos: cellPos,
    scope: null, callDepth: 0,
    num: toNum, str: toStr, bool: toBool, err: sdErr,
    scalar: function (v) { return isArrayLike(v) ? sdErr('#VALUE!') : v; },
    flat: flattenArgs,
    resolveSheet: function (nm) { var s = wb.getSheet(nm); return s ? s.name : null; }
  };
  ctx.eval = function (node, cx) { return evalNode(node, cx || ctx); };
  return ctx;
}
// A non-formula cell's literal value.
function parseLiteral(raw) {
  var s = String(raw).trim();
  if (s === '') return null;
  if (/^(TRUE|FALSE)$/i.test(s)) return /^TRUE/i.test(s);
  if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) return parseFloat(s);
  if (ERROR_CODES.indexOf(s.toUpperCase()) >= 0) return sdErr(s.toUpperCase());
  return String(raw);
}

/* ============================== Workbook =============================== */
function Workbook() {
  this.sheets = [];
  this.fwd = {};       // cellKey -> {depKey: 1}
  this.rev = {};       // cellKey -> {dependentKey: 1}
  this.volatiles = {}; // cellKey -> 1
  this.spills = {};    // anchorKey -> {sheet, c, r, rows, cols, anchorRef}
  this.addSheet('Sheet1');
}
Workbook.prototype.uniqueName = function (base) {
  base = String(base || 'Sheet').trim() || 'Sheet';
  var name = base, i = 2;
  while (this.getSheet(name)) name = base + ' (' + (i++) + ')';
  return name;
};
Workbook.prototype.addSheet = function (name) {
  name = this.uniqueName(name);
  this.sheets.push({ name: name, cells: {} });
  return name;
};
Workbook.prototype.getSheet = function (name) {
  if (name == null) return null;
  var want = String(name).toLowerCase();
  for (var i = 0; i < this.sheets.length; i++)
    if (this.sheets[i].name.toLowerCase() === want) return this.sheets[i];
  return null;
};
Workbook.prototype.sheetNames = function () {
  return this.sheets.map(function (s) { return s.name; });
};
// Rename a sheet and rewrite qualified references in every formula.
Workbook.prototype.renameSheet = function (oldName, newName) {
  var sh = this.getSheet(oldName);
  if (!sh) return false;
  newName = String(newName == null ? '' : newName).trim();
  if (!newName || /['!]/.test(newName)) return false;
  if (sh.name === newName) return true;
  if (this.getSheet(newName)) return false; // name taken
  var oldQ = oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  var reQuoted = new RegExp("'" + oldQ + "'!", 'g');
  var reBare = new RegExp("(^|[^A-Za-z0-9_.'\"])" + oldQ + "!", 'gi');
  for (var i = 0; i < this.sheets.length; i++) {
    var sht = this.sheets[i], refs = Object.keys(sht.cells);
    for (var k = 0; k < refs.length; k++) {
      var cell = sht.cells[refs[k]];
      if (cell.raw && cell.raw.charAt(0) === '=') {
        cell.raw = cell.raw
          .replace(reQuoted, "'" + newName + "'!")
          .replace(reBare, function (m, p) { return p + newName + '!'; });
      }
    }
  }
  sh.name = newName;
  this.reparseAll();
  return true;
};
Workbook.prototype.removeSheet = function (name) {
  if (this.sheets.length <= 1) return false;
  var want = String(name).toLowerCase(), idx = -1;
  for (var i = 0; i < this.sheets.length; i++)
    if (this.sheets[i].name.toLowerCase() === want) idx = i;
  if (idx < 0) return false;
  this.sheets.splice(idx, 1);
  this.reparseAll(); // refs to the deleted sheet become #REF!
  return true;
};
Workbook.prototype.dropDeps = function (key) {
  var deps = this.fwd[key] || {}, self = this;
  Object.keys(deps).forEach(function (d) { if (self.rev[d]) delete self.rev[d][key]; });
  delete this.fwd[key];
};
Workbook.prototype.parseCellInto = function (sheetName, key, cell) {
  this.dropDeps(key);
  delete cell.value;
  cell.ast = null; cell.deps = {}; cell.isFormula = false;
  cell.volatile = false; cell.parseError = null;
  var raw = cell.raw || '';
  if (raw.charAt(0) === '=') {
    cell.isFormula = true;
    try { cell.ast = parse(raw.slice(1)); }
    catch (e) { cell.parseError = '#VALUE!'; }
    if (cell.ast) {
      cell.deps = collectDeps(this, sheetName, cell.ast);
      cell.volatile = hasVolatile(cell.ast);
      var self = this;
      Object.keys(cell.deps).forEach(function (d) {
        (self.rev[d] = self.rev[d] || {})[key] = 1;
      });
    }
  }
  this.fwd[key] = cell.deps;
  if (cell.volatile) this.volatiles[key] = 1; else delete this.volatiles[key];
};
Workbook.prototype.setCell = function (sheetName, c, r, raw) {
  var sh = this.getSheet(sheetName);
  if (!sh) return;
  var ref = indexToCol(c) + (r + 1), key = sh.name + '!' + ref;
  if (raw == null || String(raw) === '') {
    if (sh.cells[ref]) {
      delete sh.cells[ref];
      this.dropDeps(key);
      delete this.spills[key];
      this.recalcFrom(key);
      this.recheckSpills(sh.name, c, r, key);
    }
    return;
  }
  var cell = sh.cells[ref] || (sh.cells[ref] = {});
  cell.raw = String(raw);
  this.parseCellInto(sh.name, key, cell);
  this.recalcFrom(key);
  this.recheckSpills(sh.name, c, r, key);
};
// After an edit at (c,r): any OTHER spill covering that cell may have become
// blocked (or a #SPILL! anchor may now fit) -> recompute those anchors.
Workbook.prototype.recheckSpills = function (sheetName, c, r, exceptKey) {
  var self = this, need = {};
  var k, s;
  for (k in this.spills) {
    if (k === exceptKey) continue;
    s = this.spills[k];
    if (s.sheet === sheetName && c >= s.c && c < s.c + s.cols && r >= s.r && r < s.r + s.rows)
      need[k] = 1;
  }
  var sh = this.getSheet(sheetName);
  if (sh) {
    var refs = Object.keys(sh.cells);
    for (var i = 0; i < refs.length; i++) {
      var cell = sh.cells[refs[i]], ak = sh.name + '!' + refs[i];
      if (ak !== exceptKey && cell.value && isErr(cell.value) && cell.value.__err === '#SPILL!')
        need[ak] = 1;
    }
  }
  Object.keys(need).forEach(function (ak2) { self.recalcFrom(ak2); });
};
Workbook.prototype.setFormat = function (sheetName, c, r, patch) {
  var sh = this.getSheet(sheetName);
  if (!sh) return;
  var ref = indexToCol(c) + (r + 1);
  var cell = sh.cells[ref] || (sh.cells[ref] = { raw: '' });
  cell.fmt = Object.assign({}, cell.fmt, patch);
};
Workbook.prototype.getFormat = function (sheetName, c, r) {
  var sh = this.getSheet(sheetName);
  if (!sh) return {};
  var cell = sh.cells[indexToCol(c) + (r + 1)];
  return (cell && cell.fmt) || {};
};
Workbook.prototype.getRaw = function (sheetName, c, r) {
  var sh = this.getSheet(sheetName);
  if (!sh) return '';
  var cell = sh.cells[indexToCol(c) + (r + 1)];
  return cell ? cell.raw : '';
};
Workbook.prototype.getValue = function (sheetName, c, r) {
  return this.readCellValue(sheetName, c, r);
};
// Spill lookup: is (c,r) inside a spill range on this sheet?
// Returns {key, anchorRef, vr, vc, anchor} or null.
Workbook.prototype.spillAt = function (sheetName, c, r) {
  for (var k in this.spills) {
    var s = this.spills[k];
    if (s.sheet !== sheetName) continue;
    if (c >= s.c && c < s.c + s.cols && r >= s.r && r < s.r + s.rows)
      return { key: k, anchorRef: s.anchorRef, vr: r - s.r, vc: c - s.c, anchor: (c === s.c && r === s.r) };
  }
  return null;
};
// Place (or clear) the spill for an anchor cell that just evaluated to v.
// Returns the value to store: the array, or #SPILL! when blocked/out of grid.
Workbook.prototype.updateSpill = function (key, v, pos) {
  var ix = key.lastIndexOf('!');
  var sh = this.getSheet(key.slice(0, ix));
  delete this.spills[key];
  if (!sh || !v || !v.__array) return v;
  var rows = v.rows, cols = v.cols;
  if (pos.r + rows > MAX_ROWS || pos.c + cols > MAX_COLS) return sdErr('#SPILL!');
  var rr, cc, k, s;
  for (rr = pos.r; rr < pos.r + rows; rr++)
    for (cc = pos.c; cc < pos.c + cols; cc++) {
      if (rr === pos.r && cc === pos.c) continue; // the anchor itself is fine
      if (sh.cells[indexToCol(cc) + (rr + 1)]) return sdErr('#SPILL!');
    }
  for (k in this.spills) { // no overlapping a different spill
    s = this.spills[k];
    if (s.sheet !== sh.name) continue;
    if (pos.r < s.r + s.rows && pos.r + rows > s.r &&
        pos.c < s.c + s.cols && pos.c + cols > s.c) return sdErr('#SPILL!');
  }
  this.spills[key] = {
    sheet: sh.name, c: pos.c, r: pos.r, rows: rows, cols: cols,
    anchorRef: indexToCol(pos.c) + (pos.r + 1)
  };
  return v;
};
Workbook.prototype.readCellValue = function (sheetName, c, r) {
  var sh = this.getSheet(sheetName);
  if (!sh) return sdErr('#REF!');
  var ref = indexToCol(c) + (r + 1);
  var sp = this.spillAt(sheetName, c, r);
  if (sp && !sp.anchor) {
    // Display mirror of another cell's spill (not directly editable).
    var acell = sh.cells[sp.anchorRef];
    if (acell && !('value' in acell)) this.evalKey(sp.key);
    var av = acell ? acell.value : null;
    if (!av || !av.__array) return null; // blocked anchor -> blank, like Excel
    return av.values[sp.vr][sp.vc];
  }
  var cell = sh.cells[ref];
  if (!cell) return null;
  if (!('value' in cell)) this.evalKey(sh.name + '!' + ref);
  var v = cell.value;
  if (v && v.__array) {
    if (isErr(v)) return v;
    return v.values[0][0]; // implicit intersection: anchor ref -> top-left value
  }
  return v == null ? null : v;
};
Workbook.prototype.readRangeValues = function (sheetName, c1, r1, c2, r2) {
  var rows = [], self = this;
  for (var r = r1; r <= r2; r++) {
    var row = [];
    for (var c = c1; c <= c2; c++) row.push(self.readCellValue(sheetName, c, r));
    rows.push(row);
  }
  return rows;
};
Workbook.prototype.evalKey = function (key) {
  var i = key.lastIndexOf('!');
  var sh = this.getSheet(key.slice(0, i));
  if (!sh) return;
  var ref = key.slice(i + 1), cell = sh.cells[ref];
  if (!cell) return;
  var v;
  if (cell.parseError) v = sdErr(cell.parseError);
  else if (!cell.isFormula) v = parseLiteral(cell.raw);
  else {
    var rc = splitRef(ref);
    var pos = { c: colToIndex(rc.col), r: rc.row - 1 };
    v = evalNode(cell.ast, makeCtx(this, sh.name, key, pos));
    if (v && v.__range) { // a bare range reference spills its values, Excel 365-style
      v = { __array: true, rows: v.values.length, cols: v.values[0].length, values: v.values };
    }
    v = this.updateSpill(key, v, pos);
  }
  cell.value = v;
};
// Recalculate exactly the cells affected by a change (dependents + volatiles),
// in topological order; anything left over is in a cycle -> #CYCLE!.
Workbook.prototype.recalcSet = function (affected) {
  var keys = Object.keys(affected), i, k, d;
  var indeg = {};
  for (i = 0; i < keys.length; i++) indeg[keys[i]] = 0;
  for (i = 0; i < keys.length; i++) {
    var f = this.fwd[keys[i]] || {};
    for (d in f) if (affected[d]) indeg[keys[i]]++;
  }
  var ready = keys.filter(function (x) { return indeg[x] === 0; });
  var done = {};
  while (ready.length) {
    k = ready.pop(); done[k] = 1; this.evalKey(k);
    var rdeps = this.rev[k] || {};
    for (d in rdeps) if (affected[d] && !done[d] && --indeg[d] === 0) ready.push(d);
  }
  for (i = 0; i < keys.length; i++) {
    if (!done[keys[i]]) {
      var ix = keys[i].lastIndexOf('!');
      var sh = this.getSheet(keys[i].slice(0, ix));
      if (sh) { var cell = sh.cells[keys[i].slice(ix + 1)]; if (cell) cell.value = sdErr('#CYCLE!'); }
    }
  }
};
Workbook.prototype.recalcFrom = function (startKey) {
  var affected = {}, queue = [startKey], k;
  affected[startKey] = 1;
  var self = this;
  Object.keys(this.volatiles).forEach(function (vk) {
    if (!affected[vk]) { affected[vk] = 1; queue.push(vk); }
  });
  function pushSpillCells(ak) { // dependents may sit on any cell of the spill
    var s = self.spills[ak];
    if (!s) return;
    for (var rr = s.r; rr < s.r + s.rows; rr++)
      for (var cc = s.c; cc < s.c + s.cols; cc++) {
        var q = s.sheet + '!' + indexToCol(cc) + (rr + 1);
        if (!affected[q]) { affected[q] = 1; queue.push(q); }
      }
  }
  pushSpillCells(startKey);
  while (queue.length) {
    k = queue.pop();
    pushSpillCells(k);
    var rdeps = this.rev[k] || {};
    Object.keys(rdeps).forEach(function (d) {
      if (!affected[d]) { affected[d] = 1; queue.push(d); }
    });
  }
  this.recalcSet(affected);
};
Workbook.prototype.recalcAll = function () {
  var affected = {}, self = this;
  this.sheets.forEach(function (sh) {
    Object.keys(sh.cells).forEach(function (ref) {
      var cell = sh.cells[ref];
      if (cell.isFormula || cell.parseError) affected[sh.name + '!' + ref] = 1;
    });
  });
  this.recalcSet(affected);
};
Workbook.prototype.reparseAll = function () {
  this.fwd = {}; this.rev = {}; this.volatiles = {}; this.spills = {};
  var self = this;
  this.sheets.forEach(function (sh) {
    Object.keys(sh.cells).forEach(function (ref) {
      self.parseCellInto(sh.name, sh.name + '!' + ref, sh.cells[ref]);
    });
  });
  this.recalcAll();
};
// Native file format: {app:'sheetdesk', version:1, sheets:[{name, cells:{A1:{raw, fmt}}}]}
Workbook.prototype.toJSON = function () {
  return {
    app: 'sheetdesk', version: 1,
    sheets: this.sheets.map(function (sh) {
      var cells = {};
      Object.keys(sh.cells).forEach(function (ref) {
        var cell = sh.cells[ref], o = { raw: cell.raw };
        if (cell.fmt && Object.keys(cell.fmt).length) o.fmt = cell.fmt;
        cells[ref] = o;
      });
      return { name: sh.name, cells: cells };
    })
  };
};
// True when the cell holds a formula calling AGGREGATE/SUBTOTAL
// (used by AGGREGATE's options 0-3 "ignore nested" semantics).
Workbook.prototype.cellHasAggregate = function (sheetName, c, r) {
  var sh = this.getSheet(sheetName);
  if (!sh) return false;
  var cell = sh.cells[indexToCol(c) + (r + 1)];
  return !!(cell && cell.isFormula && cell.ast && astHasAggregateCall(cell.ast));
};
// Bounding box of real cells + spill extents (for export/status).
Workbook.prototype.usedRange = function (sheetName) {
  var sh = this.getSheet(sheetName);
  if (!sh) return null;
  var r1 = MAX_ROWS, c1 = MAX_COLS, r2 = -1, c2 = -1;
  Object.keys(sh.cells).forEach(function (ref) {
    var rc = splitRef(ref);
    if (!rc) return;
    var c = colToIndex(rc.col), r = rc.row - 1;
    if (r < r1) r1 = r; if (c < c1) c1 = c;
    if (r > r2) r2 = r; if (c > c2) c2 = c;
  });
  for (var k in this.spills) {
    var s = this.spills[k];
    if (s.sheet !== sheetName) continue;
    if (s.r < r1) r1 = s.r; if (s.c < c1) c1 = s.c;
    if (s.r + s.rows - 1 > r2) r2 = s.r + s.rows - 1;
    if (s.c + s.cols - 1 > c2) c2 = s.c + s.cols - 1;
  }
  return r2 < 0 ? null : { r1: r1, c1: c1, r2: r2, c2: c2 };
};
Workbook.prototype.loadJSON = function (data) {
  this.sheets = []; this.fwd = {}; this.rev = {}; this.volatiles = {}; this.spills = {};
  var self = this;
  ((data && data.sheets) || []).forEach(function (s) {
    var name = self.addSheet(s.name || 'Sheet');
    var sh = self.getSheet(name);
    Object.keys(s.cells || {}).forEach(function (ref) {
      var c = s.cells[ref] || {};
      sh.cells[ref] = { raw: String(c.raw == null ? '' : c.raw), fmt: c.fmt || undefined };
    });
  });
  if (!this.sheets.length) this.addSheet('Sheet1');
  this.reparseAll();
};

/* ================================ exports ============================== */
return {
  Workbook: Workbook,
  parse: parse,
  parseFormula: parse,
  astToText: astToText,
  shiftAst: shiftAst,
  registerFunction: registerFunction,
  FUNCTIONS: FUNCTIONS,
  colToIndex: colToIndex,
  indexToCol: indexToCol,
  splitRef: splitRef,
  MAX_ROWS: MAX_ROWS,
  MAX_COLS: MAX_COLS,
  dateToSerial: dateToSerial,
  serialToDate: serialToDate,
  numToString: numToString,
  formatText: formatText,
  isErr: isErr,
  ERROR_CODES: ERROR_CODES
};
});
