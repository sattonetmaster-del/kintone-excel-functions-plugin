/*
 * Excel関数プラグイン - config.js（設定画面）
 *
 * 計算ルール（結果フィールドコード＋式）を複数登録する。
 * 保存時は全ルールの式を構文チェックし、エラーがあれば保存しない。
 * 設定は kintone.plugin.app.setConfig に JSON 文字列で保存する。
 */
(function () {
  'use strict';

  var PLUGIN_ID = kintone.$PLUGIN_ID;

  function el(id) { return document.getElementById(id); }

  // 設定画面のURL（/k/admin/app/<アプリID>/plugin/config?...）から
  // アプリ設定のプラグイン一覧画面のURLを組み立てる
  function appSettingsUrl() {
    var m = window.location.pathname.match(/^\/k\/admin\/app\/(\d+)/);
    if (m) {
      return '/k/admin/app/' + m[1] + '/plugin/';
    }
    return '/k/admin/app/';
  }

  function addRuleRow(target, formula) {
    var tbody = el('rules-body');
    var tr = document.createElement('tr');

    var tdTarget = document.createElement('td');
    tdTarget.style.padding = '4px';
    var inputTarget = document.createElement('input');
    inputTarget.type = 'text';
    inputTarget.className = 'rule-target';
    inputTarget.style.width = '95%';
    inputTarget.placeholder = '例：合計金額';
    inputTarget.value = target || '';
    tdTarget.appendChild(inputTarget);

    var tdFormula = document.createElement('td');
    tdFormula.style.padding = '4px';
    var inputFormula = document.createElement('input');
    inputFormula.type = 'text';
    inputFormula.className = 'rule-formula';
    inputFormula.style.width = '95%';
    inputFormula.placeholder = '例：AVERAGE([明細.単価])';
    inputFormula.value = formula || '';
    tdFormula.appendChild(inputFormula);

    var tdDel = document.createElement('td');
    tdDel.style.padding = '4px';
    tdDel.style.textAlign = 'center';
    var btnDel = document.createElement('button');
    btnDel.type = 'button';
    btnDel.textContent = '削除';
    btnDel.addEventListener('click', function () {
      tbody.removeChild(tr);
    });
    tdDel.appendChild(btnDel);

    tr.appendChild(tdTarget);
    tr.appendChild(tdFormula);
    tr.appendChild(tdDel);
    tbody.appendChild(tr);
  }

  function collectRules() {
    var rules = [];
    var rows = el('rules-body').querySelectorAll('tr');
    rows.forEach(function (row) {
      var target = row.querySelector('.rule-target').value.trim();
      var formula = row.querySelector('.rule-formula').value.trim();
      // 完全に空の行は無視する
      if (target === '' && formula === '') return;
      rules.push({ target: target, formula: formula });
    });
    return rules;
  }

  function validateRules(rules) {
    var errors = [];
    rules.forEach(function (rule, i) {
      var label = 'ルール' + (i + 1);
      if (rule.target === '') {
        errors.push(label + '：結果を入れるフィールドコードが空です');
        return;
      }
      if (rule.formula === '') {
        errors.push(label + '：式が空です');
        return;
      }
      try {
        ExcelFormula.validate(rule.formula);
      } catch (e) {
        errors.push(label + '：' + e.message);
      }
    });
    return errors;
  }

  function init() {
    var config = kintone.plugin.app.getConfig(PLUGIN_ID);
    var rules = [];
    if (config && config.rules) {
      try {
        rules = JSON.parse(config.rules);
      } catch (e) {
        rules = [];
      }
    }
    if (rules.length === 0) {
      addRuleRow('', '');
    } else {
      rules.forEach(function (rule) { addRuleRow(rule.target, rule.formula); });
    }

    el('add-rule').addEventListener('click', function () {
      addRuleRow('', '');
    });

    el('save-config').addEventListener('click', function () {
      var collected = collectRules();
      var errors = validateRules(collected);
      if (errors.length > 0) {
        el('config-error').textContent = errors.join('\n');
        return;
      }
      el('config-error').textContent = '';
      kintone.plugin.app.setConfig(
        { rules: JSON.stringify(collected) },
        function () {
          alert('設定を保存しました。アプリ設定画面に戻り、「アプリを更新」してください。');
          window.location.href = appSettingsUrl();
        }
      );
    });

    el('cancel-config').addEventListener('click', function () {
      window.location.href = appSettingsUrl();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
