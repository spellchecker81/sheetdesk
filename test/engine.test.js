/* SheetDesk engine tests — plain node, no framework.
 * Run: node test/engine.test.js   (from ~/workspace/sheetdesk)
 */
'use strict';
const F = require('../app/formula.js');

let pass = 0, fail = 0;
const failures = [];
function t(name, fn) {
  try { fn(); pass++; }
  catch (e) { fail++; failures.push('FAIL: ' + name + ' — ' + (e && e.message)); }
}
function eq(a, b, msg) {
  if (a !== b) throw new Error((msg ? msg + ': ' : '') + 'expected ' + JSON.stringify(b) + ', got ' + JSON.stringify(a));
}
function eqErr(v, code, msg) {
  if (!F.isErr(v) || v.__err !== code)
    throw new Error((msg ? msg + ': ' : '') + 'expected error ' + code + ', got ' + JSON.stringify(v));
}
function approx(a, b, eps, msg) {
  if (typeof a !== 'number' || Math.abs(a - b) > (eps || 1e-9))
    throw new Error((msg ? msg + ': ' : '') + 'expected ~' + b + ', got ' + JSON.stringify(a));
}
function fresh() { return new F.Workbook(); }
function set(wb, ref, raw, sheet) {
  const s = F.splitRef(ref);
  wb.setCell(sheet || 'Sheet1', F.colToIndex(s.col), s.row - 1, raw);
}
function get(wb, ref, sheet) {
  const s = F.splitRef(ref);
  return wb.getValue(sheet || 'Sheet1', F.colToIndex(s.col), s.row - 1);
}
function raw(wb, ref, sheet) {
  const s = F.splitRef(ref);
  return wb.getRaw(sheet || 'Sheet1', F.colToIndex(s.col), s.row - 1);
}

/* ---------------- arithmetic, precedence, strings ---------------- */
t('arithmetic precedence', () => {
  const wb = fresh();
  set(wb, 'A1', '=2+3*4'); eq(get(wb, 'A1'), 14);
  set(wb, 'A2', '=(2+3)*4'); eq(get(wb, 'A2'), 20);
  set(wb, 'A3', '=2^3^2'); eq(get(wb, 'A3'), 512);          // right-assoc
  set(wb, 'A4', '=-3^2'); eq(get(wb, 'A4'), -9);            // Excel: -(3^2)
  set(wb, 'A5', '=(-3)^2'); eq(get(wb, 'A5'), 9);
  set(wb, 'A6', '=2^-3'); eq(get(wb, 'A6'), 0.125);
  set(wb, 'A7', '=--5'); eq(get(wb, 'A7'), 5);
  set(wb, 'A8', '=2*-3'); eq(get(wb, 'A8'), -6);
  set(wb, 'A9', '=50%'); eq(get(wb, 'A9'), 0.5);
  set(wb, 'A10', '=200*10%'); eq(get(wb, 'A10'), 20);
  set(wb, 'A11', '=7/2'); eq(get(wb, 'A11'), 3.5);
  set(wb, 'A12', '=10-4-3'); eq(get(wb, 'A12'), 3);
});
t('string concat & comparisons', () => {
  const wb = fresh();
  set(wb, 'A1', '="a"&"b"'); eq(get(wb, 'A1'), 'ab');
  set(wb, 'A2', '=1&2'); eq(get(wb, 'A2'), '12');
  set(wb, 'A3', '="x"&TRUE'); eq(get(wb, 'A3'), 'xTRUE');
  set(wb, 'A4', '=1="1"'); eq(get(wb, 'A4'), false);        // Excel: different types
  set(wb, 'A5', '="a"="A"'); eq(get(wb, 'A5'), true);       // case-insensitive
  set(wb, 'A6', '="b">"a"'); eq(get(wb, 'A6'), true);
  set(wb, 'A7', '=1<2'); eq(get(wb, 'A7'), true);
  set(wb, 'A8', '=5>=5'); eq(get(wb, 'A8'), true);
  set(wb, 'A9', '=3<>4'); eq(get(wb, 'A9'), true);
  set(wb, 'A10', '=TRUE+1'); eq(get(wb, 'A10'), 2);
  set(wb, 'A11', '="5"+1'); eq(get(wb, 'A11'), 6);          // numeric text coerces
});
t('cell refs and ranges', () => {
  const wb = fresh();
  set(wb, 'A1', '5');
  set(wb, 'B1', '=A1*2'); eq(get(wb, 'B1'), 10);
  set(wb, 'B2', '=$A$1+1'); eq(get(wb, 'B2'), 6);           // absolute ref evaluates
  set(wb, 'A2', '1'); set(wb, 'A3', '2'); set(wb, 'A4', '3');
  set(wb, 'B3', '=SUM(A2:A4)'); eq(get(wb, 'B3'), 6);
  // minimal recalc through the graph
  set(wb, 'C1', '=B1+1'); set(wb, 'D1', '=C1+1');
  eq(get(wb, 'D1'), 12);
  set(wb, 'A1', '10'); eq(get(wb, 'B1'), 20); eq(get(wb, 'C1'), 21); eq(get(wb, 'D1'), 22);
});
t('shiftAst respects absolute refs (fill support)', () => {
  const ast = F.parseFormula('$A$1+A1');
  const shifted = F.shiftAst(ast, 1, 1);
  eq(F.astToText(shifted), '($A$1+B2)');
});
t('cross-sheet refs', () => {
  const wb = fresh();
  wb.addSheet('Data');
  set(wb, 'A1', '42', 'Data');
  set(wb, 'A1', '=Data!A1*2'); eq(get(wb, 'A1'), 84);
  wb.addSheet('My Sheet');
  set(wb, 'A1', '7', 'My Sheet');
  set(wb, 'A2', "='My Sheet'!A1+1"); eq(get(wb, 'A2'), 8);
  eq(wb.renameSheet('Data', 'Info'), true);
  eq(get(wb, 'A1'), 84);                                  // ref rewritten on rename
  eq(raw(wb, 'A1').indexOf('Info!') >= 0, true);
});
t('errors: propagation, #NAME?, #REF! on sheet delete', () => {
  const wb = fresh();
  set(wb, 'A1', '=1/0'); eqErr(get(wb, 'A1'), '#DIV/0!');
  set(wb, 'B1', '=A1+1'); eqErr(get(wb, 'B1'), '#DIV/0!'); // propagates
  set(wb, 'C1', '=SUM(A1,5)'); eqErr(get(wb, 'C1'), '#DIV/0!');
  set(wb, 'D1', '=FOOBAR(1)'); eqErr(get(wb, 'D1'), '#NAME?');
  set(wb, 'E1', '=NOSUCHCELL+1'); eqErr(get(wb, 'E1'), '#NAME?');
  set(wb, 'F1', '=(1+'); eqErr(get(wb, 'F1'), '#VALUE!');  // syntax error
  set(wb, 'G1', '=#NULL!'); eqErr(get(wb, 'G1'), '#NULL!');
  set(wb, 'H1', '=SQRT(-1)'); eqErr(get(wb, 'H1'), '#NUM!');
  set(wb, 'I1', '=VLOOKUP(9,A1:A1,1,FALSE)'); eqErr(get(wb, 'I1'), '#N/A');
  wb.addSheet('Gone');
  set(wb, 'A1', '1', 'Gone');
  set(wb, 'J1', '=Gone!A1+1'); eq(get(wb, 'J1'), 2);
  wb.removeSheet('Gone');
  eqErr(get(wb, 'J1'), '#REF!');
});
t('cycle detection', () => {
  const wb = fresh();
  set(wb, 'A1', '=B1+1');
  set(wb, 'B1', '=A1+1');
  eqErr(get(wb, 'A1'), '#CYCLE!');
  eqErr(get(wb, 'B1'), '#CYCLE!');
  set(wb, 'A1', '5');                                     // break the cycle
  eq(get(wb, 'A1'), 5); eq(get(wb, 'B1'), 6);
});
t('error literal in comparisons is an error, blank coerces to 0', () => {
  const wb = fresh();
  set(wb, 'A1', '=Z99+1'); eq(get(wb, 'A1'), 1);           // blank -> 0
  set(wb, 'A2', '=Z99=0'); eq(get(wb, 'A2'), true);
});

/* ------------------------- math functions ------------------------- */
t('SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/PRODUCT', () => {
  const wb = fresh();
  set(wb, 'A1', '1'); set(wb, 'A2', '2'); set(wb, 'A3', 'x'); set(wb, 'A4', 'TRUE');
  set(wb, 'B1', '=SUM(A1:A4)'); eq(get(wb, 'B1'), 3);      // text/logical ignored in ranges
  set(wb, 'B2', '=SUM(1,2,3)'); eq(get(wb, 'B2'), 6);
  set(wb, 'B3', '=SUM("5",TRUE)'); eq(get(wb, 'B3'), 6);   // direct args coerce
  set(wb, 'B4', '=AVERAGE(A1:A2)'); eq(get(wb, 'B4'), 1.5);
  set(wb, 'B5', '=AVERAGE(Z1:Z2)'); eqErr(get(wb, 'B5'), '#DIV/0!');
  set(wb, 'B6', '=MIN(A1:A4)'); eq(get(wb, 'B6'), 1);
  set(wb, 'B7', '=MAX(A1:A4)'); eq(get(wb, 'B7'), 2);
  set(wb, 'B8', '=MAX(Z1:Z2)'); eq(get(wb, 'B8'), 0);
  set(wb, 'B9', '=COUNT(A1:A4)'); eq(get(wb, 'B9'), 2);
  set(wb, 'B10', '=COUNT(1,TRUE)'); eq(get(wb, 'B10'), 2);
  set(wb, 'B11', '=COUNTA(A1:A4)'); eq(get(wb, 'B11'), 4);
  set(wb, 'B12', '=PRODUCT(2,3,4)'); eq(get(wb, 'B12'), 24);
  set(wb, 'B13', '=SUM("abc")'); eqErr(get(wb, 'B13'), '#VALUE!');
});
t('ABS/ROUND family/INT/TRUNC/MOD/POWER/SQRT', () => {
  const wb = fresh();
  const cases = [
    ['=ABS(-5)', 5], ['=ROUND(2.5,0)', 3], ['=ROUND(-2.5,0)', -3],
    ['=ROUND(123.456,1)', 123.5], ['=ROUND(123.456,-1)', 120],
    ['=ROUNDUP(2.1,0)', 3], ['=ROUNDUP(-2.1,0)', -3],
    ['=ROUNDDOWN(2.9,0)', 2], ['=ROUNDDOWN(-2.9,0)', -2],
    ['=INT(2.9)', 2], ['=INT(-2.1)', -3],
    ['=TRUNC(2.99)', 2], ['=TRUNC(-2.99)', -2],
    ['=MOD(10,3)', 1], ['=MOD(-10,3)', 2],
    ['=POWER(2,10)', 1024], ['=SQRT(16)', 4], ['=PI()', Math.PI],
  ];
  cases.forEach(([f, v], i) => { set(wb, 'A' + (i + 1), f); eq(get(wb, 'A' + (i + 1)), v, f); });
  set(wb, 'B1', '=MOD(10,0)'); eqErr(get(wb, 'B1'), '#DIV/0!');
  set(wb, 'B2', '=SQRT(-4)'); eqErr(get(wb, 'B2'), '#NUM!');
  set(wb, 'B3', '=POWER(-1,0.5)'); eqErr(get(wb, 'B3'), '#NUM!');
});
t('SUMIF/COUNTIF/AVERAGEIF', () => {
  const wb = fresh();
  ['1', '2', '3', '4'].forEach((v, i) => set(wb, 'A' + (i + 1), v));
  ['10', '20', '30', '40'].forEach((v, i) => set(wb, 'B' + (i + 1), v));
  set(wb, 'C1', '=SUMIF(A1:A4,">2")'); eq(get(wb, 'C1'), 7);
  set(wb, 'C2', '=SUMIF(A1:A4,">2",B1:B4)'); eq(get(wb, 'C2'), 70);
  set(wb, 'C3', '=COUNTIF(A1:A4,">2")'); eq(get(wb, 'C3'), 2);
  set(wb, 'C4', '=COUNTIF(A1:A4,"<>")'); eq(get(wb, 'C4'), 4);
  set(wb, 'C5', '=AVERAGEIF(A1:A4,">2")'); eq(get(wb, 'C5'), 3.5);
  set(wb, 'C6', '=AVERAGEIF(A1:A4,">99")'); eqErr(get(wb, 'C6'), '#DIV/0!');
  ['apple', 'apricot', 'banana'].forEach((v, i) => set(wb, 'D' + (i + 1), v));
  set(wb, 'C7', '=COUNTIF(D1:D3,"a*")'); eq(get(wb, 'C7'), 2);
  set(wb, 'C8', '=COUNTIF(D1:D3,"APPLE")'); eq(get(wb, 'C8'), 1); // case-insensitive
});

/* ------------------------- logic functions -------------------------- */
t('IF/AND/OR/NOT/XOR/IFERROR', () => {
  const wb = fresh();
  set(wb, 'A1', '=IF(1>2,"y","n")'); eq(get(wb, 'A1'), 'n');
  set(wb, 'A2', '=IF(TRUE,1,1/0)'); eq(get(wb, 'A2'), 1);   // short-circuit
  set(wb, 'A3', '=IF(FALSE,1)'); eq(get(wb, 'A3'), false); // omitted else -> FALSE
  set(wb, 'A4', '=AND(TRUE,1)'); eq(get(wb, 'A4'), true);
  set(wb, 'A5', '=AND(TRUE,FALSE)'); eq(get(wb, 'A5'), false);
  set(wb, 'A6', '=AND(FALSE,1/0)'); eq(get(wb, 'A6'), false);
  set(wb, 'A7', '=OR(FALSE,0)'); eq(get(wb, 'A7'), false);
  set(wb, 'A8', '=OR(FALSE,1/0)'); eqErr(get(wb, 'A8'), '#DIV/0!');
  set(wb, 'A9', '=OR(TRUE,1/0)'); eq(get(wb, 'A9'), true);
  set(wb, 'A10', '=NOT(TRUE)'); eq(get(wb, 'A10'), false);
  set(wb, 'A11', '=XOR(TRUE,FALSE)'); eq(get(wb, 'A11'), true);
  set(wb, 'A12', '=XOR(TRUE,TRUE)'); eq(get(wb, 'A12'), false);
  set(wb, 'A13', '=IFERROR(1/0,"oops")'); eq(get(wb, 'A13'), 'oops');
  set(wb, 'A14', '=IFERROR(5,"x")'); eq(get(wb, 'A14'), 5);
});
t('IFNA/ISERR/ISNA/ERROR.TYPE', () => {
  const wb = fresh();
  set(wb, 'A1', '1'); set(wb, 'A2', '2'); set(wb, 'A3', '3');
  set(wb, 'B1', '=VLOOKUP(9,A1:A3,1,FALSE)'); eqErr(get(wb, 'B1'), '#N/A');
  set(wb, 'B2', '=IFNA(VLOOKUP(9,A1:A3,1,FALSE),"miss")'); eq(get(wb, 'B2'), 'miss');
  set(wb, 'B3', '=IFNA(1/0,"x")'); eqErr(get(wb, 'B3'), '#DIV/0!'); // only #N/A caught
  set(wb, 'B4', '=IFNA(5,"x")'); eq(get(wb, 'B4'), 5);
  set(wb, 'B5', '=ISERR(1/0)'); eq(get(wb, 'B5'), true);
  set(wb, 'B6', '=ISERR(VLOOKUP(9,A1:A3,1,FALSE))'); eq(get(wb, 'B6'), false);
  set(wb, 'B7', '=ISERR(5)'); eq(get(wb, 'B7'), false);
  set(wb, 'B8', '=ISNA(VLOOKUP(9,A1:A3,1,FALSE))'); eq(get(wb, 'B8'), true);
  set(wb, 'B9', '=ISNA(1/0)'); eq(get(wb, 'B9'), false);
  set(wb, 'C1', '=ERROR.TYPE(#NULL!)'); eq(get(wb, 'C1'), 1);
  set(wb, 'C2', '=ERROR.TYPE(1/0)'); eq(get(wb, 'C2'), 2);
  set(wb, 'C3', '=ERROR.TYPE("a"+1)'); eq(get(wb, 'C3'), 3);
  set(wb, 'C4', '=ERROR.TYPE(A1000)'); eq(get(wb, 'C4'), 4);  // out-of-grid -> #REF!
  set(wb, 'C5', '=ERROR.TYPE(FOO(1))'); eq(get(wb, 'C5'), 5);
  set(wb, 'C6', '=ERROR.TYPE(SQRT(-1))'); eq(get(wb, 'C6'), 6);
  set(wb, 'C7', '=ERROR.TYPE(VLOOKUP(9,A1:A3,1,FALSE))'); eq(get(wb, 'C7'), 7);
  set(wb, 'C8', '=ERROR.TYPE(5)'); eqErr(get(wb, 'C8'), '#N/A');
});

/* ------------------------- text functions ----------------------------- */
t('CONCAT/TEXTJOIN/LEFT/RIGHT/MID/LEN/TRIM/UPPER/LOWER/SUBSTITUTE', () => {
  const wb = fresh();
  set(wb, 'A1', '=CONCAT("a",1,TRUE)'); eq(get(wb, 'A1'), 'a1TRUE');
  set(wb, 'A2', '=TEXTJOIN(",",TRUE,"a","","b")'); eq(get(wb, 'A2'), 'a,b');
  set(wb, 'A3', '=TEXTJOIN(",",FALSE,"a","","b")'); eq(get(wb, 'A3'), 'a,,b');
  set(wb, 'A4', '=TEXTJOIN(" - ",TRUE,"x","y")'); eq(get(wb, 'A4'), 'x - y');
  set(wb, 'A5', '=LEFT("hello",2)'); eq(get(wb, 'A5'), 'he');
  set(wb, 'A6', '=RIGHT("hello",2)'); eq(get(wb, 'A6'), 'lo');
  set(wb, 'A7', '=MID("hello",2,3)'); eq(get(wb, 'A7'), 'ell');
  set(wb, 'A8', '=MID("hello",2,99)'); eq(get(wb, 'A8'), 'ello');
  set(wb, 'A9', '=LEN("hello")'); eq(get(wb, 'A9'), 5);
  set(wb, 'A10', '=TRIM("  a  b  ")'); eq(get(wb, 'A10'), 'a b');
  set(wb, 'A11', '=UPPER("abc")'); eq(get(wb, 'A11'), 'ABC');
  set(wb, 'A12', '=LOWER("ABC")'); eq(get(wb, 'A12'), 'abc');
  set(wb, 'A13', '=SUBSTITUTE("aaa","a","b",2)'); eq(get(wb, 'A13'), 'aba');
  set(wb, 'A14', '=SUBSTITUTE("aaa","a","b")'); eq(get(wb, 'A14'), 'bbb');
  set(wb, 'A15', '=LEFT("hi",-1)'); eqErr(get(wb, 'A15'), '#VALUE!');
});
t('TEXT/VALUE', () => {
  const wb = fresh();
  set(wb, 'A1', '=TEXT(0.5,"0%")'); eq(get(wb, 'A1'), '50%');
  set(wb, 'A2', '=TEXT(46303,"yyyy-mm-dd")'); eq(get(wb, 'A2'), '2026-10-08');
  set(wb, 'A3', '=TEXT(1234.5,"#,##0.00")'); eq(get(wb, 'A3'), '1,234.50');
  set(wb, 'A4', '=TEXT(3.14159,"0.00")'); eq(get(wb, 'A4'), '3.14');
  set(wb, 'A5', '=TEXT(46303,"mm/dd/yyyy")'); eq(get(wb, 'A5'), '10/08/2026');
  set(wb, 'B1', '=VALUE("123")'); eq(get(wb, 'B1'), 123);
  set(wb, 'B2', '=VALUE("50%")'); eq(get(wb, 'B2'), 0.5);
  set(wb, 'B3', '=VALUE("abc")'); eqErr(get(wb, 'B3'), '#VALUE!');
});
t('TEXTBEFORE/TEXTAFTER', () => {
  const wb = fresh();
  set(wb, 'A1', '=TEXTBEFORE("a,b,c",",")'); eq(get(wb, 'A1'), 'a');
  set(wb, 'A2', '=TEXTBEFORE("a,b,c",",",2)'); eq(get(wb, 'A2'), 'a,b');
  set(wb, 'A3', '=TEXTBEFORE("a,b,c",",",-1)'); eq(get(wb, 'A3'), 'a,b');
  set(wb, 'A4', '=TEXTAFTER("a,b,c",",")'); eq(get(wb, 'A4'), 'b,c');
  set(wb, 'A5', '=TEXTAFTER("a,b,c",",",-1)'); eq(get(wb, 'A5'), 'c');
  set(wb, 'A6', '=TEXTAFTER("a,b,c",",",2)'); eq(get(wb, 'A6'), 'c');
  set(wb, 'A7', '=TEXTBEFORE("a,b",",",2,0,1)'); eq(get(wb, 'A7'), 'a,b');   // match_end
  set(wb, 'A8', '=TEXTAFTER("a,b",",",2,0,1)'); eq(get(wb, 'A8'), '');       // match_end
  set(wb, 'A9', '=TEXTBEFORE("abc","x",1,0,0,"nf")'); eq(get(wb, 'A9'), 'nf'); // if_not_found
  set(wb, 'A10', '=TEXTBEFORE("abc","x")'); eqErr(get(wb, 'A10'), '#N/A');
  set(wb, 'A11', '=TEXTBEFORE("abc",",",0)'); eqErr(get(wb, 'A11'), '#VALUE!');
  set(wb, 'A12', '=TEXTBEFORE("aXb","x",1,1)'); eq(get(wb, 'A12'), 'a');      // match_mode=1
  set(wb, 'A13', '=TEXTBEFORE("aXb","x",1,0)'); eqErr(get(wb, 'A13'), '#N/A'); // case-sensitive miss
  set(wb, 'A14', '=TEXTBEFORE("abc","")'); eqErr(get(wb, 'A14'), '#VALUE!');
});

/* ------------------------- date functions ----------------------------- */
t('DATE serial correctness (independently anchored)', () => {
  // Excel serial 2023-01-01 = 44927 (well-known anchor). 2026-10-08 is
  // 365+366+365 days later to 2026-01-01, +273 (Jan-Sep) +7 (Oct 1->8).
  const expected = 44927 + 365 + 366 + 365 + 273 + 7;
  eq(expected, 46303);
  const wb = fresh();
  set(wb, 'A1', '=DATE(2026,10,8)'); eq(get(wb, 'A1'), expected);
  set(wb, 'A2', '=DATE(2026,13,1)');           // month overflow
  set(wb, 'A3', '=DATE(2027,1,1)');
  eq(get(wb, 'A2'), get(wb, 'A3'));
  set(wb, 'A4', '=YEAR(DATE(2026,10,8))'); eq(get(wb, 'A4'), 2026);
  set(wb, 'A5', '=MONTH(DATE(2026,10,8))'); eq(get(wb, 'A5'), 10);
  set(wb, 'A6', '=DAY(DATE(2026,10,8))'); eq(get(wb, 'A6'), 8);
  set(wb, 'A7', '=WEEKDAY(DATE(2026,10,8))'); eq(get(wb, 'A7'), 5); // Thursday
  set(wb, 'A8', '=WEEKDAY(DATE(2026,10,8),2)'); eq(get(wb, 'A8'), 4);
  set(wb, 'A9', '=WEEKDAY(DATE(2026,10,8),3)'); eq(get(wb, 'A9'), 3);
  set(wb, 'A10', '=EDATE(DATE(2026,1,31),1)');
  set(wb, 'A11', '=DATE(2026,2,28)');
  eq(get(wb, 'A10'), get(wb, 'A11'));          // clamped to month end
  set(wb, 'A12', '=EOMONTH(DATE(2026,2,15),0)');
  eq(get(wb, 'A12'), get(wb, 'A11'));
  set(wb, 'A13', '=DATE(1899,1,1)'); eqErr(get(wb, 'A13'), '#NUM!');
});
t('volatile functions return plausible values', () => {
  const wb = fresh();
  set(wb, 'A1', '=TODAY()');
  const today = get(wb, 'A1');
  eq(typeof today, 'number');
  eq(today >= 46303 && today < 50000, true, 'TODAY plausible');
  set(wb, 'A2', '=NOW()');
  const now = get(wb, 'A2');
  eq(typeof now, 'number');
  eq(now >= today && now < today + 1, true, 'NOW plausible');
  set(wb, 'A3', '=RAND()');
  const r = get(wb, 'A3');
  eq(r >= 0 && r < 1, true, 'RAND range');
  set(wb, 'A4', '=RANDBETWEEN(1,6)');
  const rb = get(wb, 'A4');
  eq(Number.isInteger(rb) && rb >= 1 && rb <= 6, true, 'RANDBETWEEN range');
});

/* ------------------------- lookup functions --------------------------- */
t('VLOOKUP/HLOOKUP exact and approximate', () => {
  const wb = fresh();
  set(wb, 'A1', '1'); set(wb, 'B1', 'a');
  set(wb, 'A2', '2'); set(wb, 'B2', 'b');
  set(wb, 'A3', '3'); set(wb, 'B3', 'c');
  set(wb, 'C1', '=VLOOKUP(2,A1:B3,2,FALSE)'); eq(get(wb, 'C1'), 'b');
  set(wb, 'C2', '=VLOOKUP(9,A1:B3,2,FALSE)'); eqErr(get(wb, 'C2'), '#N/A');
  set(wb, 'C3', '=VLOOKUP(2.5,A1:B3,2)'); eq(get(wb, 'C3'), 'b'); // approximate
  set(wb, 'C4', '=VLOOKUP(0.5,A1:B3,2)'); eqErr(get(wb, 'C4'), '#N/A');
  set(wb, 'C5', '=VLOOKUP(2,A1:B3,9,FALSE)'); eqErr(get(wb, 'C5'), '#REF!');
  set(wb, 'C6', '=VLOOKUP(2,A1:B3,0,FALSE)'); eqErr(get(wb, 'C6'), '#VALUE!');
  set(wb, 'D1', '=HLOOKUP("b",A1:B3,2,FALSE)');
  // HLOOKUP searches the FIRST ROW: A1=1,B1="a" -> "b" not found
  eqErr(get(wb, 'D1'), '#N/A');
  set(wb, 'E1', 'x'); set(wb, 'F1', 'y');
  set(wb, 'E2', '10'); set(wb, 'F2', '20');
  set(wb, 'D2', '=HLOOKUP("y",E1:F2,2,FALSE)'); eq(get(wb, 'D2'), 20);
});
t('MATCH/INDEX/XLOOKUP/CHOOSE', () => {
  const wb = fresh();
  set(wb, 'A1', '1'); set(wb, 'B1', 'a');
  set(wb, 'A2', '2'); set(wb, 'B2', 'b');
  set(wb, 'A3', '3'); set(wb, 'B3', 'c');
  set(wb, 'C1', '=MATCH(2,A1:A3,0)'); eq(get(wb, 'C1'), 2);
  set(wb, 'C2', '=MATCH("b",B1:B3,0)'); eq(get(wb, 'C2'), 2);
  set(wb, 'C3', '=MATCH(2.5,A1:A3,1)'); eq(get(wb, 'C3'), 2);
  set(wb, 'C4', '=MATCH(9,A1:A3,0)'); eqErr(get(wb, 'C4'), '#N/A');
  set(wb, 'C5', '=INDEX(A1:B3,2,2)'); eq(get(wb, 'C5'), 'b');
  set(wb, 'C6', '=INDEX(A1:B3,9,1)'); eqErr(get(wb, 'C6'), '#REF!');
  set(wb, 'C7', '=XLOOKUP(2,A1:A3,B1:B3)'); eq(get(wb, 'C7'), 'b');
  set(wb, 'C8', '=XLOOKUP(9,A1:A3,B1:B3,"nf")'); eq(get(wb, 'C8'), 'nf');
  set(wb, 'C9', '=XLOOKUP(9,A1:A3,B1:B3)'); eqErr(get(wb, 'C9'), '#N/A');
  set(wb, 'C10', '=CHOOSE(2,"a","b","c")'); eq(get(wb, 'C10'), 'b');
  set(wb, 'C11', '=CHOOSE(5,"a")'); eqErr(get(wb, 'C11'), '#VALUE!');
});
t('ROW/COLUMN/ROWS/COLUMNS', () => {
  const wb = fresh();
  set(wb, 'B3', '=ROW()'); eq(get(wb, 'B3'), 3);
  set(wb, 'B2', '=COLUMN()'); eq(get(wb, 'B2'), 2);
  set(wb, 'C1', '=ROW(A5)'); eq(get(wb, 'C1'), 5);
  set(wb, 'C2', '=COLUMN(C1)'); eq(get(wb, 'C2'), 3);
  set(wb, 'C3', '=ROWS(A1:B3)'); eq(get(wb, 'C3'), 3);
  set(wb, 'C4', '=COLUMNS(A1:B3)'); eq(get(wb, 'C4'), 2);
  set(wb, 'C5', '=ROWS(5)'); eq(get(wb, 'C5'), 1);
});
t('ISNUMBER/ISTEXT/ISBLANK/ISERROR', () => {
  const wb = fresh();
  set(wb, 'A1', '5');
  set(wb, 'B1', '=ISNUMBER(A1)'); eq(get(wb, 'B1'), true);
  set(wb, 'B2', '=ISNUMBER("5")'); eq(get(wb, 'B2'), false);
  set(wb, 'B3', '=ISTEXT("x")'); eq(get(wb, 'B3'), true);
  set(wb, 'B4', '=ISBLANK(Z99)'); eq(get(wb, 'B4'), true);
  set(wb, 'B5', '=ISBLANK(A1)'); eq(get(wb, 'B5'), false);
  set(wb, 'B6', '=ISERROR(1/0)'); eq(get(wb, 'B6'), true);
  set(wb, 'B7', '=ISERROR(1)'); eq(get(wb, 'B7'), false);
});
t('nested functions and case-insensitivity', () => {
  const wb = fresh();
  set(wb, 'A1', '=SUM(IF(1,"2",0),MAX(1,2))'); eq(get(wb, 'A1'), 4);
  set(wb, 'A2', '=sUm(1,2)'); eq(get(wb, 'A2'), 3);
  set(wb, 'A3', '=LeFt("hello",2)'); eq(get(wb, 'A3'), 'he');
  set(wb, 'A4', '=vlookup(2,A1:A2,1,FALSE)');
  eqErr(get(wb, 'A4'), '#N/A'); // A1:A2 hold formulas, no match for 2
});
t('function registry is extensible', () => {
  eq(typeof F.FUNCTIONS['PROFITMARGIN'], 'object');
  eq(typeof F.FUNCTIONS['SUM'], 'object');
  F.registerFunction('DOUBLEIT', {
    minArgs: 1, maxArgs: 1, desc: 'test double',
    fn: (args, C) => { const x = C.num(args[0]); return F.isErr(x) ? x : x * 2; }
  });
  const wb = fresh();
  set(wb, 'A1', '=DOUBLEIT(21)'); eq(get(wb, 'A1'), 42);
  set(wb, 'A2', '=doubleit(2)'); eq(get(wb, 'A2'), 4); // case-insensitive
  set(wb, 'A3', '=PROFITMARGIN(100,60)'); eq(get(wb, 'A3'), 0.4);
  set(wb, 'A4', '=PROFITMARGIN(0,5)'); eqErr(get(wb, 'A4'), '#DIV/0!');
  delete F.FUNCTIONS['DOUBLEIT'];
});

/* ---------------- dynamic array functions ---------------- */
t('FILTER basic, multi-condition, empty -> #CALC!, if_empty', () => {
  const wb = fresh();
  ['5', '2', '8', '1', '9'].forEach((v, i) => set(wb, 'A' + (i + 1), v));
  set(wb, 'C1', '=FILTER(A1:A5,A1:A5>2)');
  eq(get(wb, 'C1'), 5); eq(get(wb, 'C2'), 8); eq(get(wb, 'C3'), 9);
  eq(get(wb, 'C4'), null);
  eq(raw(wb, 'C2'), '');                       // spilled cells hold no formula
  set(wb, 'E1', '=FILTER(A1:A5,(A1:A5>2)*(A1:A5<9))');  // multiple conditions
  eq(get(wb, 'E1'), 5); eq(get(wb, 'E2'), 8); eq(get(wb, 'E3'), null);
  set(wb, 'G1', '=FILTER(A1:A5,A1:A5>100)');
  eqErr(get(wb, 'G1'), '#CALC!');               // empty result
  set(wb, 'H1', '=FILTER(A1:A5,A1:A5>100,"none")');
  eq(get(wb, 'H1'), 'none');                   // if_empty
});
t('FILTER can filter columns with a horizontal include', () => {
  const wb = fresh();
  set(wb, 'A1', '1'); set(wb, 'B1', '2'); set(wb, 'C1', '3');
  set(wb, 'A2', '4'); set(wb, 'B2', '5'); set(wb, 'C2', '6');
  set(wb, 'E1', '=FILTER(A1:C2,{1,0,1})');
  eq(get(wb, 'E1'), 1); eq(get(wb, 'F1'), 3);
  eq(get(wb, 'E2'), 4); eq(get(wb, 'F2'), 6);
  eq(get(wb, 'G1'), null);
});
t('SORT ascending/descending and nested multi-key', () => {
  const wb = fresh();
  ['3', '1', '4', '1', '5'].forEach((v, i) => set(wb, 'B' + (i + 1), v));
  set(wb, 'D1', '=SORT(B1:B5)');
  [1, 1, 3, 4, 5].forEach((v, i) => eq(get(wb, 'D' + (i + 1)), v));
  set(wb, 'F1', '=SORT(B1:B5,1,-1)');
  [5, 4, 3, 1, 1].forEach((v, i) => eq(get(wb, 'F' + (i + 1)), v));
  // multi-key via nesting: sort by col 2, then col 1
  set(wb, 'A1', '3'); set(wb, 'B1', 'c');
  set(wb, 'A2', '1'); set(wb, 'B2', 'b');
  set(wb, 'A3', '1'); set(wb, 'B3', 'a');
  set(wb, 'A4', '2'); set(wb, 'B4', 'd');
  set(wb, 'D1', '=SORT(SORT(A1:B4,2),1)');
  eq(get(wb, 'D1'), 1); eq(get(wb, 'E1'), 'a');
  eq(get(wb, 'D2'), 1); eq(get(wb, 'E2'), 'b');
  eq(get(wb, 'D3'), 2); eq(get(wb, 'E3'), 'd');
  eq(get(wb, 'D4'), 3); eq(get(wb, 'E4'), 'c');
});
t('SORTBY multi-key', () => {
  const wb = fresh();
  set(wb, 'A1', '2'); set(wb, 'B1', 'b');
  set(wb, 'A2', '1'); set(wb, 'B2', 'd');
  set(wb, 'A3', '2'); set(wb, 'B3', 'a');
  set(wb, 'A4', '1'); set(wb, 'B4', 'c');
  set(wb, 'D1', '=SORTBY(A1:B4,A1:A4,1,B1:B4,-1)');
  eq(get(wb, 'D1'), 1); eq(get(wb, 'E1'), 'd');
  eq(get(wb, 'D2'), 1); eq(get(wb, 'E2'), 'c');
  eq(get(wb, 'D3'), 2); eq(get(wb, 'E3'), 'b');
  eq(get(wb, 'D4'), 2); eq(get(wb, 'E4'), 'a');
});
t('UNIQUE rows, by_col, exactly_once', () => {
  const wb = fresh();
  ['1', '2', '2', '3', '1'].forEach((v, i) => set(wb, 'A' + (i + 1), v));
  set(wb, 'C1', '=UNIQUE(A1:A5)');
  eq(get(wb, 'C1'), 1); eq(get(wb, 'C2'), 2); eq(get(wb, 'C3'), 3);
  eq(get(wb, 'C4'), null);
  set(wb, 'E1', '=UNIQUE(A1:A5,0,1)');          // exactly_once
  eq(get(wb, 'E1'), 3); eq(get(wb, 'E2'), null);
  set(wb, 'A7', '7'); set(wb, 'B7', '8'); set(wb, 'C7', '7');
  set(wb, 'E7', '=UNIQUE(A7:C7,1)');           // by_col
  eq(get(wb, 'E7'), 7); eq(get(wb, 'F7'), 8); eq(get(wb, 'G7'), null);
});
t('CHOOSECOLS/CHOOSEROWS incl. negative indices', () => {
  const wb = fresh();
  set(wb, 'A1', '1'); set(wb, 'B1', '2'); set(wb, 'C1', '3');
  set(wb, 'A2', '4'); set(wb, 'B2', '5'); set(wb, 'C2', '6');
  set(wb, 'E1', '=CHOOSECOLS(A1:C2,3,1)');
  eq(get(wb, 'E1'), 3); eq(get(wb, 'F1'), 1);
  eq(get(wb, 'E2'), 6); eq(get(wb, 'F2'), 4);
  set(wb, 'H1', '=CHOOSECOLS(A1:C2,-1)');       // last column
  eq(get(wb, 'H1'), 3); eq(get(wb, 'H2'), 6);
  set(wb, 'J1', '=CHOOSEROWS(A1:C2,2,-2)');
  eq(get(wb, 'J1'), 4); eq(get(wb, 'K1'), 5); eq(get(wb, 'L1'), 6);
  eq(get(wb, 'J2'), 1); eq(get(wb, 'K2'), 2); eq(get(wb, 'L2'), 3);
  set(wb, 'N1', '=CHOOSECOLS(A1:C2,9)'); eqErr(get(wb, 'N1'), '#VALUE!');
});
t('DROP/TAKE incl. negative args', () => {
  const wb = fresh();
  ['1', '2', '3', '4', '5'].forEach((v, i) => set(wb, 'A' + (i + 1), v));
  set(wb, 'C1', '=DROP(A1:A5,2)');
  eq(get(wb, 'C1'), 3); eq(get(wb, 'C2'), 4); eq(get(wb, 'C3'), 5); eq(get(wb, 'C4'), null);
  set(wb, 'E1', '=DROP(A1:A5,-2)');            // drop from end
  eq(get(wb, 'E1'), 1); eq(get(wb, 'E2'), 2); eq(get(wb, 'E3'), 3); eq(get(wb, 'E4'), null);
  set(wb, 'G1', '=DROP(A1:A5,9)'); eqErr(get(wb, 'G1'), '#CALC!');
  set(wb, 'I1', '=TAKE(A1:A5,2)');
  eq(get(wb, 'I1'), 1); eq(get(wb, 'I2'), 2); eq(get(wb, 'I3'), null);
  set(wb, 'K1', '=TAKE(A1:A5,-2)');            // take from end
  eq(get(wb, 'K1'), 4); eq(get(wb, 'K2'), 5); eq(get(wb, 'K3'), null);
});
t('HSTACK/VSTACK with #N/A padding and nesting', () => {
  const wb = fresh();
  set(wb, 'A1', '1'); set(wb, 'A2', '2');
  set(wb, 'B1', '10'); set(wb, 'B2', '20'); set(wb, 'B3', '30');
  set(wb, 'D1', '=HSTACK(A1:A2,B1:B3)');
  eq(get(wb, 'D1'), 1); eq(get(wb, 'E1'), 10);
  eq(get(wb, 'D2'), 2); eq(get(wb, 'E2'), 20);
  eqErr(get(wb, 'D3'), '#N/A'); eq(get(wb, 'E3'), 30);  // padded
  set(wb, 'G1', '1'); set(wb, 'H1', '2');
  set(wb, 'G3', '5'); set(wb, 'G4', '6');
  set(wb, 'J1', '=VSTACK(G1:H1,G3:G4)');
  eq(get(wb, 'J1'), 1); eq(get(wb, 'K1'), 2);
  eq(get(wb, 'J2'), 5); eqErr(get(wb, 'K2'), '#N/A');
  eq(get(wb, 'J3'), 6); eqErr(get(wb, 'K3'), '#N/A');
  // nesting with another dynamic-array result
  ['1', '2', '2', '3'].forEach((v, i) => set(wb, 'M' + (i + 1), v));
  set(wb, 'O1', '=VSTACK(UNIQUE(M1:M4),99)');
  eq(get(wb, 'O1'), 1); eq(get(wb, 'O2'), 2);
  eq(get(wb, 'O3'), 3); eq(get(wb, 'O4'), 99);
});
t('elementwise operator lifting (dynamic-array arithmetic)', () => {
  const wb = fresh();
  set(wb, 'A1', '1'); set(wb, 'A2', '2'); set(wb, 'A3', '3');
  set(wb, 'C1', '=A1:A3*10');
  eq(get(wb, 'C1'), 10); eq(get(wb, 'C2'), 20); eq(get(wb, 'C3'), 30);
  set(wb, 'E1', '=IF(A1:A3>1,"big","small")');
  eq(get(wb, 'E1'), 'small'); eq(get(wb, 'E2'), 'big'); eq(get(wb, 'E3'), 'big');
});

/* ------------------------- spill behavior ------------------------- */
t('spill placement across cells; anchor shows top-left', () => {
  const wb = fresh();
  set(wb, 'A1', '=HSTACK(1,2,3)');
  eq(get(wb, 'A1'), 1);
  eq(get(wb, 'B1'), 2);
  eq(get(wb, 'C1'), 3);
  eq(get(wb, 'D1'), null);
  eq(raw(wb, 'B1'), '');                       // spilled cells hold no raw content
  const sp = wb.spillAt('Sheet1', 1, 0);
  eq(sp && sp.anchor, false);
  eq(sp && sp.anchorRef, 'A1');
});
t('#SPILL! when blocked, cleared when unblocked', () => {
  const wb = fresh();
  set(wb, 'A1', '=HSTACK(1,2,3)');
  eq(get(wb, 'C1'), 3);
  set(wb, 'B1', 'x');                          // block the spill
  eqErr(get(wb, 'A1'), '#SPILL!');
  eq(get(wb, 'B1'), 'x');                      // blocker keeps its own value
  set(wb, 'B1', '');                           // unblock
  eq(get(wb, 'A1'), 1);
  eq(get(wb, 'B1'), 2);
  eq(get(wb, 'C1'), 3);
});
t('spill edits: typing into a spill range blocks the source spill', () => {
  const wb = fresh();
  set(wb, 'A1', '=VSTACK(1,2,3)');
  eq(get(wb, 'A3'), 3);
  set(wb, 'A2', '9');                          // type inside the spill range
  eqErr(get(wb, 'A1'), '#SPILL!');
  eq(get(wb, 'A2'), 9);
  set(wb, 'A2', '');
  eq(get(wb, 'A1'), 1);
  eq(get(wb, 'A2'), 2);
});
t('formulas referencing spilled cells depend on the anchor', () => {
  const wb = fresh();
  set(wb, 'A1', '=HSTACK(1,2,3)');
  set(wb, 'D1', '=B1*10');                     // B1 is a spilled cell
  eq(get(wb, 'D1'), 20);
  set(wb, 'A1', '=HSTACK(5,6,7)');             // anchor changes -> dependent recalcs
  eq(get(wb, 'D1'), 60);
  eq(get(wb, 'B1'), 6);
});
t('anchor reference uses implicit intersection (top-left)', () => {
  const wb = fresh();
  set(wb, 'A1', '=HSTACK(1,2)');
  set(wb, 'C1', '=A1+1');
  eq(get(wb, 'C1'), 2);
});
t('bare range reference spills', () => {
  const wb = fresh();
  set(wb, 'A1', '7'); set(wb, 'A2', '8');
  set(wb, 'C1', '=A1:A2');
  eq(get(wb, 'C1'), 7); eq(get(wb, 'C2'), 8);
});

/* ------------------------- LET and LAMBDA --------------------------- */
t('LET sequencing and shadowing', () => {
  const wb = fresh();
  set(wb, 'A1', '=LET(x,5,x*2)'); eq(get(wb, 'A1'), 10);
  set(wb, 'A2', '=LET(x,5,y,x*2,x+y)'); eq(get(wb, 'A2'), 15); // later sees earlier
  set(wb, 'A3', '=LET(x,1,x,2,x)'); eq(get(wb, 'A3'), 2);     // shadowing
  set(wb, 'A4', '=LET(x,1/0,5)'); eqErr(get(wb, 'A4'), '#DIV/0!'); // eager
  set(wb, 'A5', '=LET(x,y,y,1,x)'); eqErr(get(wb, 'A5'), '#NAME?'); // forward ref fails
  set(wb, 'A6', '=LET(data,A1:A2,SUM(data))');
  set(wb, 'B1', '3'); set(wb, 'B2', '4');
  // A6 was set before B1:B2 existed; re-point at B1:B2
  set(wb, 'A6', '=LET(data,B1:B2,SUM(data))'); eq(get(wb, 'A6'), 7);
});
t('LAMBDA immediate invocation', () => {
  const wb = fresh();
  set(wb, 'A1', '=LAMBDA(x,x*2)(5)'); eq(get(wb, 'A1'), 10);
  set(wb, 'A2', '=LAMBDA(a,b,a+b)(3,4)'); eq(get(wb, 'A2'), 7);
  set(wb, 'A3', '=LAMBDA(x,x*2)'); // a bare lambda is a value, not callable output
  eq(get(wb, 'A3') && get(wb, 'A3').__lambda, true);
});
t('LAMBDA via LET and recursion (factorial)', () => {
  const wb = fresh();
  set(wb, 'A1', '=LET(f,LAMBDA(x,x+1),f(41))'); eq(get(wb, 'A1'), 42);
  set(wb, 'A2', '=LET(f,LAMBDA(n,IF(n<=1,1,n*f(n-1))),f(5))');
  eq(get(wb, 'A2'), 120);
  set(wb, 'A3', '=LET(add,LAMBDA(a,b,a+b),add(add(1,2),add(3,4)))');
  eq(get(wb, 'A3'), 10);
  set(wb, 'A4', '=LET(f,LAMBDA(x,x*2),f(1,2))'); eqErr(get(wb, 'A4'), '#N/A'); // arity
});

/* ----------------- extended stats & AGGREGATE ----------------------- */
t('MEDIAN/MODE/LARGE/SMALL', () => {
  const wb = fresh();
  set(wb, 'A1', '=MEDIAN(1,3,2)'); eq(get(wb, 'A1'), 2);
  set(wb, 'A2', '=MEDIAN(1,2,3,4)'); eq(get(wb, 'A2'), 2.5);
  set(wb, 'A3', '=MEDIAN(Z1)'); eqErr(get(wb, 'A3'), '#NUM!');
  set(wb, 'A4', '=MODE(1,2,2,3)'); eq(get(wb, 'A4'), 2);
  set(wb, 'A5', '=MODE(1,2,3)'); eqErr(get(wb, 'A5'), '#N/A');
  set(wb, 'A6', '=MODE.SNGL(1,1,2)'); eq(get(wb, 'A6'), 1);
  ['5', '2', '8', '1', '9'].forEach((v, i) => set(wb, 'B' + (i + 1), v));
  set(wb, 'C1', '=LARGE(B1:B5,2)'); eq(get(wb, 'C1'), 8);
  set(wb, 'C2', '=SMALL(B1:B5,2)'); eq(get(wb, 'C2'), 2);
  set(wb, 'C3', '=LARGE(B1:B5,0)'); eqErr(get(wb, 'C3'), '#NUM!');
  set(wb, 'C4', '=SMALL(B1:B5,9)'); eqErr(get(wb, 'C4'), '#NUM!');
});
t('STDEV/VAR sample vs population + PERCENTILE/QUARTILE', () => {
  const wb = fresh();
  const ds = '=STDEV.S(2,4,4,4,5,5,7,9)';
  set(wb, 'A1', ds); approx(get(wb, 'A1'), 2.138089935, 1e-6);
  set(wb, 'A2', '=STDEV.P(2,4,4,4,5,5,7,9)'); approx(get(wb, 'A2'), 2, 1e-9);
  set(wb, 'A3', '=VAR.S(2,4,4,4,5,5,7,9)'); approx(get(wb, 'A3'), 32 / 7, 1e-9);
  set(wb, 'A4', '=VAR.P(2,4,4,4,5,5,7,9)'); eq(get(wb, 'A4'), 4);
  set(wb, 'A5', '=STDEV.S(5)'); eqErr(get(wb, 'A5'), '#DIV/0!');
  set(wb, 'A6', '=VAR.P(Z1)'); eqErr(get(wb, 'A6'), '#DIV/0!');
  ['1', '2', '3', '4'].forEach((v, i) => set(wb, 'B' + (i + 1), v));
  set(wb, 'C1', '=PERCENTILE.INC(B1:B4,0.5)'); eq(get(wb, 'C1'), 2.5);
  set(wb, 'C2', '=PERCENTILE.EXC(B1:B4,0.25)'); eq(get(wb, 'C2'), 1.25);
  set(wb, 'C3', '=PERCENTILE.INC(B1:B4,2)'); eqErr(get(wb, 'C3'), '#NUM!');
  set(wb, 'C4', '=QUARTILE.INC(B1:B4,2)'); eq(get(wb, 'C4'), 2.5);
  set(wb, 'C5', '=QUARTILE.EXC(B1:B4,2)'); eq(get(wb, 'C5'), 2.5);
  set(wb, 'C6', '=QUARTILE.EXC(B1:B4,0)'); eqErr(get(wb, 'C6'), '#NUM!');
});
t('AGGREGATE ignore semantics (options 0-7)', () => {
  const wb = fresh();
  set(wb, 'A1', '1'); set(wb, 'A2', '=1/0'); set(wb, 'A3', '3'); set(wb, 'A4', '4');
  set(wb, 'B1', '=AGGREGATE(9,6,A1:A4)'); eq(get(wb, 'B1'), 8);   // ignore errors
  set(wb, 'B2', '=AGGREGATE(9,2,A1:A4)'); eq(get(wb, 'B2'), 8);   // ignore errors
  set(wb, 'B3', '=AGGREGATE(9,4,A1:A4)'); eqErr(get(wb, 'B3'), '#DIV/0!'); // propagate
  set(wb, 'B4', '=AGGREGATE(9,0,A1:A4)'); eqErr(get(wb, 'B4'), '#DIV/0!'); // propagate
  set(wb, 'B5', '=AGGREGATE(1,6,A1:A4)'); approx(get(wb, 'B5'), 8 / 3, 1e-9);
  set(wb, 'B6', '=AGGREGATE(14,6,A1:A4,2)'); eq(get(wb, 'B6'), 3); // LARGE k=2
  set(wb, 'B7', '=AGGREGATE(14,6,A1:A4)'); eqErr(get(wb, 'B7'), '#VALUE!'); // k missing
  set(wb, 'B8', '=AGGREGATE(99,6,A1:A4)'); eqErr(get(wb, 'B8'), '#VALUE!');
  set(wb, 'B9', '=AGGREGATE(9,9,A1:A4)'); eqErr(get(wb, 'B9'), '#VALUE!');
  set(wb, 'B10', '=AGGREGATE(12,6,A1:A4)'); eq(get(wb, 'B10'), 3); // MEDIAN
  set(wb, 'B11', '=AGGREGATE(16,6,A1:A4,0.5)'); eq(get(wb, 'B11'), 3); // PERCENTILE.INC
  // nested AGGREGATE: options 0-3 ignore it, 4-7 do not
  set(wb, 'C1', '=AGGREGATE(9,6,A1:A4)');       // = 8
  set(wb, 'C2', '=AGGREGATE(9,0,C1:C1)'); eq(get(wb, 'C2'), 0);  // nested ignored
  set(wb, 'C3', '=AGGREGATE(9,4,C1:C1)'); eq(get(wb, 'C3'), 8);  // not ignored
  set(wb, 'C4', '=AGGREGATE(2,6,A1:A4)'); eq(get(wb, 'C4'), 3); // COUNT
  set(wb, 'C5', '=AGGREGATE(4,6,A1:A4)'); eq(get(wb, 'C5'), 4); // MAX
});
t('workbook JSON round-trip preserves formulas and spill rebuilds', () => {
  const wb = fresh();
  set(wb, 'A1', '5');
  set(wb, 'B1', '=A1*2');
  set(wb, 'C1', '=HSTACK(1,2)');
  wb.setFormat('Sheet1', 0, 0, { bold: true });
  const json = JSON.stringify(wb.toJSON());
  const wb2 = fresh();
  wb2.loadJSON(JSON.parse(json));
  eq(wb2.getValue('Sheet1', 1, 0), 10);
  eq(wb2.getValue('Sheet1', 2, 0), 1);
  eq(wb2.getValue('Sheet1', 3, 0), 2);   // spill rebuilt after load
  eq(wb2.getFormat('Sheet1', 0, 0).bold, true);
});

t('array literals {1,2;3,4}', () => {
  const wb = fresh();
  set(wb, 'A1', '={1,2;3,4}');
  eq(get(wb, 'A1'), 1); eq(get(wb, 'B1'), 2);
  eq(get(wb, 'A2'), 3); eq(get(wb, 'B2'), 4);
  set(wb, 'D1', '=SUM({1,2;3,4})'); eq(get(wb, 'D1'), 10);
  set(wb, 'D2', '={1,2;3}'); eqErr(get(wb, 'D2'), '#VALUE!'); // ragged -> parse error
  set(wb, 'D3', '=HSTACK({1;2},{3,4})');
  eq(get(wb, 'D3'), 1); eq(get(wb, 'E3'), 3); eq(get(wb, 'F3'), 4);
  eq(get(wb, 'D4'), 2); eqErr(get(wb, 'E4'), '#N/A'); eqErr(get(wb, 'F4'), '#N/A');
});

/* ------------------------------- summary ------------------------------ */
console.log('\n' + pass + ' passed, ' + fail + ' failed, ' + (pass + fail) + ' total.');
if (failures.length) {
  console.log('--- failures ---');
  failures.forEach(f => console.log(f));
  process.exit(1);
} else {
  console.log('ALL TESTS PASSED');
}
