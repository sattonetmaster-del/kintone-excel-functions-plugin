/*
 * Excel関数プラグイン - desktop.js
 *
 * レコードの新規作成・編集の保存時（submit）に、設定された全ルールを計算し、
 * 結果フィールドへ値をセットする。計算エラー時は event.error で保存を止める。
 * 作成・編集画面の表示時（show）にも初期計算を行う（エラーは無視）。
 */
(function () {
  'use strict';

  var PLUGIN_ID = kintone.$PLUGIN_ID;

  function loadRules() {
    var config = kintone.plugin.app.getConfig(PLUGIN_ID);
    if (!config || !config.rules) return [];
    try {
      var rules = JSON.parse(config.rules);
      return Array.isArray(rules) ? rules : [];
    } catch (e) {
      return [];
    }
  }

  // フィールドオブジェクトから生の値を取り出す
  function rawValue(field) {
    if (!field) return null;
    var v = field.value;
    if (v === undefined || v === null || v === '') return null;
    return v;
  }

  // レコードから式エンジン用の参照解決関数を作る
  function makeResolver(record) {
    return function (ref) {
      // 通常フィールド
      if (Object.prototype.hasOwnProperty.call(record, ref)) {
        return rawValue(record[ref]);
      }
      // サブテーブル列 [テーブルコード.列コード]
      var idx = ref.indexOf('.');
      if (idx > 0) {
        var tableCode = ref.slice(0, idx);
        var columnCode = ref.slice(idx + 1);
        var table = record[tableCode];
        if (table && table.type === 'SUBTABLE' && Array.isArray(table.value)) {
          return table.value.map(function (row) {
            var cell = row.value ? row.value[columnCode] : null;
            return rawValue(cell);
          });
        }
      }
      throw new Error('フィールド「' + ref + '」が見つかりません');
    };
  }

  function applyRules(record, rules) {
    var resolve = makeResolver(record);
    rules.forEach(function (rule) {
      var result = ExcelFormula.evaluate(rule.formula, resolve);
      var target = record[rule.target];
      if (!target) {
        throw new Error('結果を入れるフィールド「' + rule.target + '」が見つかりません');
      }
      if (result === null || result === undefined) {
        target.value = ''; // 空の結果はフィールドを空にする
      } else if (typeof result === 'boolean') {
        target.value = result ? 1 : 0;
      } else {
        target.value = result;
      }
    });
  }

  function handleSubmit(event) {
    var rules = loadRules();
    if (rules.length === 0) return event;
    try {
      applyRules(event.record, rules);
    } catch (e) {
      event.error = 'Excel関数プラグイン： ' + e.message;
    }
    return event;
  }

  function handleShow(event) {
    var rules = loadRules();
    if (rules.length === 0) return event;
    try {
      applyRules(event.record, rules);
    } catch (e) {
      // 表示時の初期計算ではエラーで画面を止めない（保存時に判定する）
    }
    return event;
  }

  kintone.events.on(
    ['app.record.create.submit', 'app.record.edit.submit'],
    handleSubmit
  );
  kintone.events.on(
    ['app.record.create.show', 'app.record.edit.show'],
    handleShow
  );
})();
