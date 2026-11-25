/*
 * Excel関数プラグイン - 式エンジン (formula.js)
 *
 * Excel風の式をトークナイズ＋再帰下降パースで評価する独立モジュール。
 * eval は使わない。ブラウザ（kintone）と Node（単体テスト）の両方で動く。
 *
 * 参照記法:
 *   [フィールドコード]            通常フィールド（スカラー）
 *   [テーブルコード.列コード]     サブテーブル列（全行の値の配列）
 *
 * 公開API:
 *   ExcelFormula.parse(式)                 -> AST（構文チェック用）
 *   ExcelFormula.validate(式)              -> true（エラー時は例外）
 *   ExcelFormula.evaluate(式, resolve)     -> 計算結果（数値 / 文字列 / 真偽値 / null=空）
 *     resolve(参照文字列) はフィールドの値を返す関数。
 *     スカラーは数値・文字列・null、テーブル列は値の配列を返す。
 */
(function (global) {
  'use strict';

  /* ================= トークナイザ ================= */

  function isDigit(ch) { return ch >= '0' && ch <= '9'; }
  function isIdentStart(ch) { return /[A-Za-z_]/.test(ch); }
  function isIdentChar(ch) { return /[A-Za-z0-9_]/.test(ch); }

  function tokenize(input) {
    var tokens = [];
    var i = 0;
    var n = input.length;
    while (i < n) {
      var ch = input[i];

      // 空白（半角・全角）
      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r' || ch === '　') {
        i++;
        continue;
      }

      // フィールド参照 [ ... ]
      if (ch === '[') {
        var end = input.indexOf(']', i);
        if (end === -1) {
          throw new Error('フィールド参照の「]」が閉じられていません（' + (i + 1) + '文字目）');
        }
        var ref = input.slice(i + 1, end).trim();
        if (!ref) {
          throw new Error('フィールド参照が空です（' + (i + 1) + '文字目）');
        }
        tokens.push({ type: 'field', value: ref, pos: i });
        i = end + 1;
        continue;
      }

      // 数値
      if (isDigit(ch) || (ch === '.' && isDigit(input[i + 1] || ''))) {
        var j = i;
        var num = '';
        var dots = 0;
        while (j < n && (isDigit(input[j]) || input[j] === '.')) {
          if (input[j] === '.') {
            dots++;
            if (dots > 1) {
              throw new Error('数値の形式が正しくありません（' + (i + 1) + '文字目）');
            }
          }
          num += input[j];
          j++;
        }
        // 指数表記 (1.5e3 など)
        if (j < n && (input[j] === 'e' || input[j] === 'E')) {
          var k = j + 1;
          if (k < n && (input[k] === '+' || input[k] === '-')) k++;
          if (k < n && isDigit(input[k])) {
            while (k < n && isDigit(input[k])) { num += input[k]; k++; }
            j = k;
          }
        }
        tokens.push({ type: 'number', value: parseFloat(num), pos: i });
        i = j;
        continue;
      }

      // 文字列リテラル " ... "
      if (ch === '"') {
        var s = i + 1;
        var str = '';
        while (s < n && input[s] !== '"') {
          str += input[s];
          s++;
        }
        if (s >= n) {
          throw new Error('文字列が「"」で閉じられていません（' + (i + 1) + '文字目）');
        }
        tokens.push({ type: 'string', value: str, pos: i });
        i = s + 1;
        continue;
      }

      // 識別子（関数名・TRUE / FALSE）
      if (isIdentStart(ch)) {
        var m = i;
        var name = '';
        while (m < n && isIdentChar(input[m])) {
          name += input[m];
          m++;
        }
        tokens.push({ type: 'ident', value: name.toUpperCase(), pos: i });
        i = m;
        continue;
      }

      // 演算子（2文字を先に判定）
      var two = input.substr(i, 2);
      if (two === '<>' || two === '>=' || two === '<=') {
        tokens.push({ type: 'op', value: two, pos: i });
        i += 2;
        continue;
      }
      if ('+-*/%^(),=<>'.indexOf(ch) !== -1) {
        tokens.push({ type: 'op', value: ch, pos: i });
        i++;
        continue;
      }

      throw new Error('使用できない文字「' + ch + '」があります（' + (i + 1) + '文字目）');
    }
    tokens.push({ type: 'eof', value: null, pos: n });
    return tokens;
  }

  /* ================= パーサ（再帰下降） ================= */

  var COMPARISON_OPS = ['=', '<>', '>', '<', '>=', '<='];
  var KNOWN_FUNCTIONS = ['SUM', 'AVERAGE', 'MAX', 'MIN', 'COUNT', 'STDEV', 'STDEVP', 'ROUND', 'IF'];

  function parse(input) {
    if (typeof input !== 'string' || input.trim() === '') {
      throw new Error('式が空です');
    }
    var tokens = tokenize(input);
    var pos = 0;

    function peek() { return tokens[pos]; }
    function next() { return tokens[pos++]; }
    function isOp(tok, v) { return tok.type === 'op' && tok.value === v; }
    function expectOp(v) {
      var tok = next();
      if (!isOp(tok, v)) {
        throw new Error('「' + v + '」が必要です（' + (tok.pos + 1) + '文字目）');
      }
    }

    function parseExpression() { return parseComparison(); }

    function parseComparison() {
      var left = parseAdditive();
      var tok = peek();
      if (tok.type === 'op' && COMPARISON_OPS.indexOf(tok.value) !== -1) {
        next();
        var right = parseAdditive();
        return { kind: 'binary', op: tok.value, left: left, right: right };
      }
      return left;
    }

    function parseAdditive() {
      var left = parseMultiplicative();
      while (isOp(peek(), '+') || isOp(peek(), '-')) {
        var op = next().value;
        var right = parseMultiplicative();
        left = { kind: 'binary', op: op, left: left, right: right };
      }
      return left;
    }

    function parseMultiplicative() {
      var left = parseUnary();
      while (isOp(peek(), '*') || isOp(peek(), '/') || isOp(peek(), '%')) {
        var op = next().value;
        var right = parseUnary();
        left = { kind: 'binary', op: op, left: left, right: right };
      }
      return left;
    }

    function parseUnary() {
      if (isOp(peek(), '-')) {
        next();
        return { kind: 'unary', op: '-', operand: parseUnary() };
      }
      if (isOp(peek(), '+')) {
        next();
        return parseUnary();
      }
      return parsePower();
    }

    function parsePower() {
      var base = parsePrimary();
      if (isOp(peek(), '^')) {
        next();
        // 右結合。指数側は単項マイナスを許す（2^-3 など）
        var exp = parseUnary();
        return { kind: 'binary', op: '^', left: base, right: exp };
      }
      return base;
    }

    function parsePrimary() {
      var tok = next();
      if (tok.type === 'number') return { kind: 'number', value: tok.value };
      if (tok.type === 'string') return { kind: 'string', value: tok.value };
      if (tok.type === 'field') return { kind: 'field', ref: tok.value };
      if (tok.type === 'ident') {
        if (tok.value === 'TRUE') return { kind: 'boolean', value: true };
        if (tok.value === 'FALSE') return { kind: 'boolean', value: false };
        // 関数呼び出し
        if (!isOp(peek(), '(')) {
          throw new Error('「' + tok.value + '」は関数名です。後ろに「( )」が必要です（' + (tok.pos + 1) + '文字目）');
        }
        if (KNOWN_FUNCTIONS.indexOf(tok.value) === -1) {
          throw new Error('未知の関数「' + tok.value + '」です');
        }
        next(); // '(' を消費
        var args = [];
        if (!isOp(peek(), ')')) {
          args.push(parseExpression());
          while (isOp(peek(), ',')) {
            next();
            args.push(parseExpression());
          }
        }
        expectOp(')');
        return { kind: 'call', name: tok.value, args: args };
      }
      if (isOp(tok, '(')) {
        var inner = parseExpression();
        expectOp(')');
        return inner;
      }
      if (tok.type === 'eof') {
        throw new Error('式が途中で終わっています');
      }
      throw new Error('予期しないトークン「' + (tok.value || tok.type) + '」があります（' + (tok.pos + 1) + '文字目）');
    }

    var ast = parseExpression();
    if (peek().type !== 'eof') {
      var extra = peek();
      throw new Error('式の後ろに余分な内容があります（' + (extra.pos + 1) + '文字目）');
    }
    return ast;
  }

  /* ================= 評価 ================= */

  function isEmptyValue(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }

  // 厳密な数値化：空は null、数値でない文字列は例外
  function toNumber(v, what) {
    if (isEmptyValue(v)) return null;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
      var num = Number(v.trim());
      if (!isNaN(num)) return num;
      throw new Error('数値でない値「' + v + '」が計算に使われました' + (what ? '（' + what + '）' : ''));
    }
    if (Array.isArray(v)) {
      throw new Error('テーブル列（複数の値）を単独の値として使うことはできません。SUM などの集計関数の中で使ってください');
    }
    throw new Error('計算できない値が使われました');
  }

  // 集計用：空・非数値は「無視」を表す undefined を返す
  function toNumberLenient(v) {
    if (isEmptyValue(v)) return undefined;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
      var num = Number(v.trim());
      return isNaN(num) ? undefined : num;
    }
    return undefined;
  }

  function collectNumbers(argNodes, resolve) {
    var values = [];
    argNodes.forEach(function (node) {
      var v = evaluateNode(node, resolve);
      if (Array.isArray(v)) {
        v.forEach(function (item) {
          var num = toNumberLenient(item);
          if (num !== undefined) values.push(num);
        });
      } else {
        var num = toNumberLenient(v);
        if (num !== undefined) values.push(num);
      }
    });
    return values;
  }

  function requireValues(values, funcName) {
    if (values.length === 0) {
      throw new Error(funcName + ' の対象となる数値がありません（空・非数値は除かれます）');
    }
    return values;
  }

  // Excel の ROUND と同じ「四捨五入（0.5 は絶対値方向に切り上げ）」
  function roundHalfAwayFromZero(x, digits) {
    var factor = Math.pow(10, digits);
    var scaled = x * factor;
    var corrected = scaled + (scaled >= 0 ? 1e-9 : -1e-9);
    var rounded = corrected >= 0 ? Math.floor(corrected + 0.5) : Math.ceil(corrected - 0.5);
    return rounded / factor;
  }

  function evaluateCall(node, resolve) {
    var name = node.name;

    if (name === 'IF') {
      if (node.args.length < 2 || node.args.length > 3) {
        throw new Error('IF は IF(条件, 真のとき, 偽のとき) の形で書いてください');
      }
      var cond = evaluateNode(node.args[0], resolve);
      var truthy;
      if (cond === null || cond === undefined) truthy = false;
      else if (typeof cond === 'boolean') truthy = cond;
      else {
        var condNum = toNumber(cond, 'IFの条件');
        truthy = condNum !== null && condNum !== 0;
      }
      if (truthy) return evaluateNode(node.args[1], resolve);
      if (node.args.length === 3) return evaluateNode(node.args[2], resolve);
      return null; // 偽のときを省略した場合は空
    }

    if (name === 'ROUND') {
      if (node.args.length < 1 || node.args.length > 2) {
        throw new Error('ROUND は ROUND(値, 桁数) の形で書いてください');
      }
      var target = toNumber(evaluateNode(node.args[0], resolve), 'ROUNDの値');
      if (target === null) return null;
      var digits = 0;
      if (node.args.length === 2) {
        var d = toNumber(evaluateNode(node.args[1], resolve), 'ROUNDの桁数');
        digits = d === null ? 0 : Math.trunc(d);
      }
      return roundHalfAwayFromZero(target, digits);
    }

    // 集計関数
    var values = collectNumbers(node.args, resolve);
    var i, sum;
    switch (name) {
      case 'SUM':
        sum = 0;
        values.forEach(function (v) { sum += v; });
        return sum;
      case 'COUNT':
        return values.length;
      case 'AVERAGE':
        requireValues(values, 'AVERAGE');
        sum = 0;
        values.forEach(function (v) { sum += v; });
        return sum / values.length;
      case 'MAX':
        requireValues(values, 'MAX');
        return Math.max.apply(null, values);
      case 'MIN':
        requireValues(values, 'MIN');
        return Math.min.apply(null, values);
      case 'STDEV':
      case 'STDEVP': {
        if (name === 'STDEV' && values.length < 2) {
          throw new Error('STDEV（標本標準偏差）には数値が2つ以上必要です');
        }
        requireValues(values, name);
        var mean = 0;
        values.forEach(function (v) { mean += v; });
        mean /= values.length;
        var sq = 0;
        values.forEach(function (v) { sq += (v - mean) * (v - mean); });
        var divisor = name === 'STDEV' ? values.length - 1 : values.length;
        return Math.sqrt(sq / divisor);
      }
      default:
        throw new Error('未知の関数「' + name + '」です');
    }
  }

  function evaluateBinary(node, resolve) {
    var op = node.op;

    if (COMPARISON_OPS.indexOf(op) !== -1) {
      var l = evaluateNode(node.left, resolve);
      var r = evaluateNode(node.right, resolve);
      if (Array.isArray(l) || Array.isArray(r)) {
        throw new Error('テーブル列（複数の値）を比較に使うことはできません');
      }
      // 空値との比較：= は false、<> は true、それ以外は false（v1仕様）
      if (isEmptyValue(l) || isEmptyValue(r)) {
        return op === '<>';
      }
      // 両方とも文字列なら文字列として比較
      if (typeof l === 'string' && typeof r === 'string' && isNaN(Number(l)) && isNaN(Number(r))) {
        switch (op) {
          case '=': return l === r;
          case '<>': return l !== r;
          case '>': return l > r;
          case '<': return l < r;
          case '>=': return l >= r;
          case '<=': return l <= r;
        }
      }
      var ln = toNumber(l);
      var rn = toNumber(r);
      switch (op) {
        case '=': return ln === rn;
        case '<>': return ln !== rn;
        case '>': return ln > rn;
        case '<': return ln < rn;
        case '>=': return ln >= rn;
        case '<=': return ln <= rn;
      }
    }

    var a = toNumber(evaluateNode(node.left, resolve));
    var b = toNumber(evaluateNode(node.right, resolve));
    // 空値を含む四則演算は「空」を伝播させる（v1仕様）
    if (a === null || b === null) return null;

    switch (op) {
      case '+': return a + b;
      case '-': return a - b;
      case '*': return a * b;
      case '/':
        if (b === 0) throw new Error('0 で割ることはできません');
        return a / b;
      case '%':
        if (b === 0) throw new Error('0 で割ることはできません（剰余）');
        return a % b;
      case '^': return Math.pow(a, b);
      default:
        throw new Error('未知の演算子「' + op + '」です');
    }
  }

  function evaluateNode(node, resolve) {
    switch (node.kind) {
      case 'number': return node.value;
      case 'string': return node.value;
      case 'boolean': return node.value;
      case 'field': return resolve(node.ref);
      case 'unary': {
        var v = toNumber(evaluateNode(node.operand, resolve));
        if (v === null) return null;
        return node.op === '-' ? -v : v;
      }
      case 'binary': return evaluateBinary(node, resolve);
      case 'call': return evaluateCall(node, resolve);
      default:
        throw new Error('内部エラー：不正な式ノードです');
    }
  }

  // 最終結果の浮動小数点の誤差を整える（0.1 + 0.2 = 0.30000000000000004 対策）
  function cleanResult(v) {
    if (typeof v === 'number' && isFinite(v)) {
      return parseFloat(v.toPrecision(12));
    }
    return v;
  }

  /* ================= 公開API ================= */

  var ExcelFormula = {
    parse: parse,
    validate: function (src) {
      parse(src);
      return true;
    },
    evaluate: function (src, resolve) {
      if (typeof resolve !== 'function') {
        throw new Error('内部エラー：フィールド解決関数がありません');
      }
      var result = evaluateNode(parse(src), resolve);
      // 単独フィールド参照などの結果は数値化できるものは数値に、空は null にそろえる
      if (typeof result === 'string') {
        if (result.trim() === '') return null;
        var num = Number(result);
        if (!isNaN(num)) return cleanResult(num);
        return result;
      }
      return cleanResult(result);
    },
    KNOWN_FUNCTIONS: KNOWN_FUNCTIONS.slice()
  };

  global.ExcelFormula = ExcelFormula;
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = ExcelFormula;
  }
})(typeof window !== 'undefined' ? window : globalThis);
