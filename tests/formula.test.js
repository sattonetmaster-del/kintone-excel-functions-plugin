/*
 * Excel関数プラグイン - 式エンジン単体テスト
 * 実行方法: node tests/formula.test.js
 */
'use strict';

const assert = require('node:assert');
const ExcelFormula = require('../src/formula.js');

let passed = 0;
let failed = 0;
const failures = [];

function test(name, fn) {
  try {
    fn();
    passed++;
    console.log('ok   - ' + name);
  } catch (e) {
    failed++;
    failures.push(name + ': ' + e.message);
    console.log('FAIL - ' + name + ': ' + e.message);
  }
}

// フィールド参照をマップで解決するテスト用リゾルバ
function resolver(map) {
  return function (ref) {
    if (Object.prototype.hasOwnProperty.call(map, ref)) return map[ref];
    throw new Error('フィールド「' + ref + '」が見つかりません');
  };
}

function ev(formula, map) {
  return ExcelFormula.evaluate(formula, resolver(map || {}));
}

function approx(actual, expected, eps) {
  assert.ok(
    Math.abs(actual - expected) < (eps || 1e-9),
    'expected ' + expected + ', got ' + actual
  );
}

/* ---------- 四則演算と優先順位 ---------- */

test('四則演算： 1 + 2 * 3 = 7（乗算が先）', () => {
  assert.strictEqual(ev('1 + 2 * 3'), 7);
});

test('括弧： (1 + 2) * 3 = 9', () => {
  assert.strictEqual(ev('(1 + 2) * 3'), 9);
});

test('減算は左結合： 10 - 4 - 3 = 3', () => {
  assert.strictEqual(ev('10 - 4 - 3'), 3);
});

test('べき乗は右結合： 2 ^ 3 ^ 2 = 512', () => {
  assert.strictEqual(ev('2 ^ 3 ^ 2'), 512);
});

test('単項マイナス： -5 + 3 = -2', () => {
  assert.strictEqual(ev('-5 + 3'), -2);
});

test('剰余： 10 % 3 = 1', () => {
  assert.strictEqual(ev('10 % 3'), 1);
});

test('除算： 10 / 4 = 2.5', () => {
  assert.strictEqual(ev('10 / 4'), 2.5);
});

test('浮動小数点の整形： 0.1 + 0.2 = 0.3', () => {
  assert.strictEqual(ev('0.1 + 0.2'), 0.3);
});

/* ---------- フィールド参照 ---------- */

test('フィールド参照と加算', () => {
  assert.strictEqual(ev('[単価] * [数量]', { '単価': 100, '数量': 3 }), 300);
});

test('kintoneの数値文字列を数値として扱う', () => {
  assert.strictEqual(ev('[金額] * 2', { '金額': '150' }), 300);
});

test('空フィールドとの四則演算は空（null）を伝播', () => {
  assert.strictEqual(ev('[空] + 1', { '空': null }), null);
  assert.strictEqual(ev('[空] + 1', { '空': '' }), null);
});

/* ---------- 集計関数（スカラー引数） ---------- */

test('SUM（複数引数）', () => {
  assert.strictEqual(ev('SUM(1, 2, 3, 4)'), 10);
});

test('AVERAGE', () => {
  assert.strictEqual(ev('AVERAGE(10, 20, 30)'), 20);
});

test('MAX / MIN', () => {
  assert.strictEqual(ev('MAX(3, 9, 5)'), 9);
  assert.strictEqual(ev('MIN(3, 9, 5)'), 3);
});

test('COUNT は数値の個数（空・非数値は数えない）', () => {
  assert.strictEqual(ev('COUNT(5, [空], "abc", 10)', { '空': null }), 2);
});

test('関数名は大文字小文字不問', () => {
  assert.strictEqual(ev('average(2, 4)'), 3);
  assert.strictEqual(ev('Sum(1, 2)'), 3);
});

/* ---------- テーブル列の集計 ---------- */

const tableMap = { '明細.単価': [100, 200, 300], '明細.数量': [1, 2, 3] };

test('SUM（テーブル列）', () => {
  assert.strictEqual(ev('SUM([明細.単価])', tableMap), 600);
});

test('AVERAGE（テーブル列）', () => {
  assert.strictEqual(ev('AVERAGE([明細.単価])', tableMap), 200);
});

test('MAX / MIN（テーブル列）', () => {
  assert.strictEqual(ev('MAX([明細.数量])', tableMap), 3);
  assert.strictEqual(ev('MIN([明細.数量])', tableMap), 1);
});

test('COUNT（テーブル列）', () => {
  assert.strictEqual(ev('COUNT([明細.単価])', tableMap), 3);
});

test('集計は空・非数値を無視する（ExcelのAVERAGEに倣う）', () => {
  assert.strictEqual(ev('AVERAGE([t.c])', { 't.c': [10, null, '', 'abc', 20] }), 15);
  assert.strictEqual(ev('SUM([t.c])', { 't.c': [10, null, '', 20] }), 30);
});

test('スカラーとテーブル列の混在引数', () => {
  assert.strictEqual(ev('SUM([明細.単価], 50)', tableMap), 650);
});

test('テーブル列同士の直接演算はエラー', () => {
  assert.throws(() => ev('[明細.単価] + 1', tableMap), /集計関数/);
});

/* ---------- STDEV / STDEVP（既知データで検算） ---------- */

// 古典的な検算データ： 平均5、母集団標準偏差2、標本標準偏差 sqrt(32/7)
const sampleData = { 't.v': [2, 4, 4, 4, 5, 5, 7, 9] };

test('STDEVP（母集団標準偏差）= 2', () => {
  approx(ev('STDEVP([t.v])', sampleData), 2);
});

test('STDEV（標本標準偏差）= sqrt(32/7)', () => {
  approx(ev('STDEV([t.v])', sampleData), Math.sqrt(32 / 7));
});

test('STDEV は数値1つではエラー', () => {
  assert.throws(() => ev('STDEV([t.v])', { 't.v': [5] }), /2つ以上/);
});

test('AVERAGE は対象数値が無いとエラー', () => {
  assert.throws(() => ev('AVERAGE([t.v])', { 't.v': [null, ''] }), /対象となる数値がありません/);
});

/* ---------- ROUND ---------- */

test('ROUND(3.14159, 2) = 3.14', () => {
  assert.strictEqual(ev('ROUND(3.14159, 2)'), 3.14);
});

test('ROUND(2.5, 0) = 3（Excel式の四捨五入）', () => {
  assert.strictEqual(ev('ROUND(2.5, 0)'), 3);
});

test('ROUND(-2.5, 0) = -3（0.5は絶対値方向）', () => {
  assert.strictEqual(ev('ROUND(-2.5, 0)'), -3);
});

test('ROUND(123.456, -1) = 120（負の桁数）', () => {
  assert.strictEqual(ev('ROUND(123.456, -1)'), 120);
});

test('ROUND の桁数省略時は 0', () => {
  assert.strictEqual(ev('ROUND(3.6)'), 4);
});

test('ROUND とフィールド参照の組み合わせ', () => {
  approx(ev('ROUND([金額] * 0.1, 0)', { '金額': 12345 }), 1235);
});

/* ---------- IF と比較 ---------- */

test('IF（数値比較・真）', () => {
  assert.strictEqual(ev('IF([金額] >= 10000, 1, 0)', { '金額': 15000 }), 1);
});

test('IF（数値比較・偽）', () => {
  assert.strictEqual(ev('IF([金額] >= 10000, 1, 0)', { '金額': 5000 }), 0);
});

test('IF（文字列比較）', () => {
  assert.strictEqual(ev('IF([状態] = "完了", 100, 0)', { '状態': '完了' }), 100);
  assert.strictEqual(ev('IF([状態] = "完了", 100, 0)', { '状態': '作業中' }), 0);
});

test('IF（<> 比較）', () => {
  assert.strictEqual(ev('IF([状態] <> "完了", 1, 0)', { '状態': '作業中' }), 1);
});

test('IF は選ばれなかった側を評価しない（遅延評価）', () => {
  assert.strictEqual(ev('IF(TRUE, 1, 1 / 0)'), 1);
  assert.strictEqual(ev('IF(FALSE, 1 / 0, 2)'), 2);
});

test('IF の偽のとき省略時は空（null）', () => {
  assert.strictEqual(ev('IF(FALSE, 1)'), null);
});

test('比較演算子の結果を四則演算に使う（真=1）', () => {
  assert.strictEqual(ev('([a] > 5) * 100', { a: 10 }), 100);
});

test('複合式： IF と集計の組み合わせ', () => {
  const map = { '明細.金額': [1000, 2000, 3000], '割引率': 0.1 };
  assert.strictEqual(
    ev('IF(SUM([明細.金額]) >= 5000, SUM([明細.金額]) * [割引率], 0)', map),
    600
  );
});

/* ---------- エラー検出 ---------- */

test('ゼロ除算はエラー', () => {
  assert.throws(() => ev('1 / 0'), /0 で割る/);
  assert.throws(() => ev('1 % 0'), /0 で割る/);
});

test('構文エラー： 括弧が閉じていない', () => {
  assert.throws(() => ev('SUM(1, 2'), /必要/);
});

test('構文エラー： 式が演算子で終わる', () => {
  assert.throws(() => ev('1 +'));
});

test('構文エラー： フィールド参照が閉じていない', () => {
  assert.throws(() => ev('[金額 + 1'), /閉じられていません/);
});

test('構文エラー： 未知の関数', () => {
  assert.throws(() => ev('FOO(1)'), /未知の関数/);
});

test('構文エラー： 式の後ろの余分な内容', () => {
  assert.throws(() => ev('1 2'), /余分/);
});

test('構文エラー： 空の式', () => {
  assert.throws(() => ev(''), /空/);
  assert.throws(() => ExcelFormula.validate('   '), /空/);
});

test('validate は正しい式で true を返す', () => {
  assert.strictEqual(ExcelFormula.validate('AVERAGE([明細.単価]) + IF([a] > 0, 1, 0)'), true);
});

test('存在しないフィールド参照はエラー', () => {
  assert.throws(() => ev('[無いフィールド] + 1', {}), /見つかりません/);
});

/* ---------- 結果 ---------- */

console.log('');
console.log('passed: ' + passed + ', failed: ' + failed);
if (failed > 0) {
  console.log('--- failures ---');
  failures.forEach((f) => console.log(f));
  process.exit(1);
}
