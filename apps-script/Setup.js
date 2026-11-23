/**
 * 試算表選單、建立 Google 表單、接收表單送出。
 * 這個專案要「綁定在一份 Google 試算表」上執行（擴充功能 → Apps Script）。
 */

var SHEET = {
  OVERVIEW: '總覽',
  ITEMS: '項目分析',
  PEOPLE: '人員分析',
  TREND: '趨勢分析',
  TIME: '時段分析',
  RECORDS: '每次紀錄',
  DETAIL: '明細',
  ROSTER: '人員名單'
};

var PROP = {
  FORM_ID: 'FORM_ID',              // 目前使用中的表單
  FORM_HISTORY: 'FORM_HISTORY',    // 所有用過的表單 ID（JSON 陣列），改版後舊資料仍可重算
  MAP_PREFIX: 'ITEM_MAP_'          // ITEM_MAP_<formId> → { 表單題目ID: 'meta:date' | 'check:D01' ... }
};

var META = {
  date: '洗冰日期',
  time: '洗冰時間（整點洗冰請選該整點，例如 14:00）',
  filler: '填表人',
  notes: '備註／異常狀況說明（選填）'
};

// ───────────────────────── 選單 ─────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('洗冰檢點')
    .addItem('① 初始化／建立表單', 'setup')
    .addItem('更新表單的人員選單', 'refreshRoster')
    .addItem('顯示表單連結', 'showFormLinks')
    .addSeparator()
    .addItem('重新計算全部資料', 'rebuildAll')
    .addItem('重建分析分頁', 'buildAnalysisSheets')
    .addSeparator()
    .addItem('改版：建立新版表單', 'createNewVersionForm')
    .addToUi();
}

// ───────────────────────── 初始化 ─────────────────────────

function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  ensureDataSheets_(ss);

  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP.FORM_ID)) {
    buildAnalysisSheets();
    showFormLinks();
    return;
  }

  var names = readRoster_();
  if (names.length === 0) {
    ss.setActiveSheet(ss.getSheetByName(SHEET.ROSTER));
    SpreadsheetApp.getUi().alert('請先在「人員名單」分頁 A 欄填入所有同仁姓名（一格一人），再執行一次「① 初始化／建立表單」。');
    return;
  }

  createForm_(names);
  buildAnalysisSheets();
  showFormLinks();
}

function ensureDataSheets_(ss) {
  var roster = getOrCreateSheet_(ss, SHEET.ROSTER);
  if (roster.getLastRow() === 0) {
    roster.getRange('A1:B1').setValues([['姓名', '說明']]).setFontWeight('bold');
    roster.getRange('B2').setValue('A 欄一格填一位同仁姓名。增減人員後，按選單「洗冰檢點 → 更新表單的人員選單」。');
    roster.setColumnWidth(1, 160).setColumnWidth(2, 520).setFrozenRows(1);
  }

  var detail = getOrCreateSheet_(ss, SHEET.DETAIL);
  if (detail.getLastRow() === 0) {
    detail.getRange(1, 1, 1, DETAIL_HEADERS.length).setValues([DETAIL_HEADERS]).setFontWeight('bold');
    detail.setFrozenRows(1);
  }
  // 文字欄：避免「2026-11」「14:00」被自動轉成日期時間
  detail.getRange('D:D').setNumberFormat('@');
  detail.getRange('H:H').setNumberFormat('@');
  detail.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  detail.getRange('C:C').setNumberFormat('yyyy/mm/dd');
  detail.getRange('G:G').setNumberFormat('yyyy/mm/dd');

  var rec = getOrCreateSheet_(ss, SHEET.RECORDS);
  if (rec.getLastRow() === 0) {
    rec.getRange(1, 1, 1, RECORD_HEADERS.length).setValues([RECORD_HEADERS]).setFontWeight('bold');
    rec.setFrozenRows(1);
  }
  rec.getRange('D:D').setNumberFormat('@');
  rec.getRange('G:G').setNumberFormat('@');
  rec.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  rec.getRange('C:C').setNumberFormat('yyyy/mm/dd');
  rec.getRange('F:F').setNumberFormat('yyyy/mm/dd');
  rec.getRange('P:P').setNumberFormat('0.0%');
  rec.getRange('R:R').setWrap(true);
  rec.setColumnWidth(18, 420);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readRoster_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ROSTER);
  if (!sh || sh.getLastRow() < 2) return [];
  var seen = {};
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (n) { if (!n || seen[n]) return false; seen[n] = true; return true; });
}

// ───────────────────────── 建立表單 ─────────────────────────

function createForm_(names) {
  var ss = SpreadsheetApp.getActive();
  var form = FormApp.create('洗冰檢點表 ' + FORM_VERSION + '版');
  form.setTitle('冰上樂園洗冰指揮官(值班)作業安全防護檢點表　' + FORM_VERSION + '版')
    .setDescription('每次洗冰結束後，由值班洗冰指揮官填寫一份，逐項勾選四位當班人員是否做到。\n\n' + SAFETY_NOTE)
    .setProgressBar(true)
    .setShowLinkToRespondAgain(true)
    .setConfirmationMessage('已送出，謝謝！');

  var map = {};
  var put = function (item, key) { map[String(item.getId())] = key; return item; };

  // 第 1 頁：基本資料
  put(form.addDateItem().setTitle(META.date).setRequired(true), 'meta:date');
  put(form.addTimeItem().setTitle(META.time).setRequired(true), 'meta:time');
  put(form.addListItem().setTitle(META.filler).setChoiceValues(names).setRequired(true), 'meta:filler');
  ROLES.forEach(function (role) {
    put(form.addListItem().setTitle(role.name + ' 姓名').setChoiceValues(names).setRequired(true), 'person:' + role.key);
  });

  // 每個角色一頁
  ROLES.forEach(function (role) {
    form.addPageBreakItem().setTitle(role.name).setHelpText('★ 為重點檢核項目。');
    PHASES.forEach(function (phase) {
      var items = ITEMS.filter(function (it) { return !it.retired && it.role === role.key && it.phase === phase.key; });
      if (items.length === 0) return;
      form.addSectionHeaderItem().setTitle('▍' + phase.name);
      items.forEach(function (it) {
        var q = form.addMultipleChoiceItem()
          .setTitle((it.key ? '★ ' : '') + it.text)
          .setChoiceValues([CHECK_OK, CHECK_NG])
          .setRequired(true);
        put(q, 'check:' + it.id);
      });
    });
  });

  form.addPageBreakItem().setTitle('備註');
  put(form.addParagraphTextItem().setTitle(META.notes), 'meta:notes');

  // 原始回覆另存一份在試算表（「表單回應」分頁），當作備份
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  var props = PropertiesService.getScriptProperties();
  props.setProperty(PROP.MAP_PREFIX + form.getId(), JSON.stringify(map));
  props.setProperty(PROP.FORM_ID, form.getId());
  var history = JSON.parse(props.getProperty(PROP.FORM_HISTORY) || '[]');
  history.push(form.getId());
  props.setProperty(PROP.FORM_HISTORY, JSON.stringify(history));

  installTrigger_(form);
  return form;
}

function installTrigger_(form) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onFormSubmitHandler') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmitHandler').forForm(form).onFormSubmit().create();
}

function getForm_() {
  var id = PropertiesService.getScriptProperties().getProperty(PROP.FORM_ID);
  if (!id) throw new Error('尚未建立表單，請先執行「① 初始化／建立表單」。');
  return FormApp.openById(id);
}

function refreshRoster() {
  var names = readRoster_();
  if (names.length === 0) throw new Error('人員名單是空的。');
  var form = getForm_();
  var map = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP.MAP_PREFIX + form.getId()));
  Object.keys(map).forEach(function (itemId) {
    if (map[itemId] === 'meta:filler' || map[itemId].indexOf('person:') === 0) {
      form.getItemById(Number(itemId)).asListItem().setChoiceValues(names);
    }
  });
  SpreadsheetApp.getActive().toast('表單人員選單已更新（' + names.length + ' 人）', '洗冰檢點');
}

function showFormLinks() {
  var form = getForm_();
  var html = HtmlService.createHtmlOutput(
    '<p><b>同仁填寫用（傳到群組、或做成 QR Code 貼在洗冰室）：</b><br>' +
    '<a href="' + form.getPublishedUrl() + '" target="_blank">' + form.getPublishedUrl() + '</a></p>' +
    '<p><b>表單編輯（管理者）：</b><br><a href="' + form.getEditUrl() + '" target="_blank">開啟編輯頁</a></p>'
  ).setWidth(520).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '表單連結');
}

// ───────────────────────── 改版 ─────────────────────────

/**
 * 改版步驟：先修改 Config.js（FORM_VERSION 與 ITEMS），再執行這個功能。
 * 舊表單會停止收件，舊資料保留在「明細」，分析時可用「表單版本」欄區分。
 */
function createNewVersionForm() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('建立 ' + FORM_VERSION + ' 版新表單？',
    '舊表單會停止收件，同仁需要改用新連結。請確認 Config.js 已改成新版題目。', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  var props = PropertiesService.getScriptProperties();
  var oldId = props.getProperty(PROP.FORM_ID);
  if (oldId) FormApp.openById(oldId).setAcceptingResponses(false)
    .setCustomClosedFormMessage('此版本已停用，請改用新版洗冰檢點表。');
  createForm_(readRoster_());
  buildAnalysisSheets();
  showFormLinks();
}

// ───────────────────────── 接收送出 ─────────────────────────

function onFormSubmitHandler(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var formId = e.source.getId();
    var rows = buildRows(parseResponse_(e.response, formId), { ROLES: ROLES, PHASES: PHASES, ITEMS: ITEMS });
    appendRows_([rows]);
  } finally {
    lock.releaseLock();
  }
}

/** 把 FormResponse 轉成 Core.buildRows 需要的格式 */
function parseResponse_(response, formId) {
  var props = PropertiesService.getScriptProperties();
  var map = JSON.parse(props.getProperty(PROP.MAP_PREFIX + formId) || '{}');
  var sub = {
    responseId: response.getId(),
    submittedAt: response.getTimestamp(),
    version: versionOfForm_(formId),
    people: {},
    answers: {}
  };
  response.getItemResponses().forEach(function (ir) {
    var key = map[String(ir.getItem().getId())];
    if (!key) return;
    var val = ir.getResponse();
    var parts = key.split(':');
    if (parts[0] === 'meta') sub[parts[1]] = val;
    else if (parts[0] === 'person') sub.people[parts[1]] = val;
    else if (parts[0] === 'check') sub.answers[parts[1]] = val;
  });
  return sub;
}

var versionCache_ = {};
function versionOfForm_(formId) {
  if (versionCache_[formId]) return versionCache_[formId];
  var title = FormApp.openById(formId).getTitle();
  var m = title.match(/(\d{5}-\d+)版/);
  versionCache_[formId] = m ? m[1] : FORM_VERSION;
  return versionCache_[formId];
}

function appendRows_(list) {
  var ss = SpreadsheetApp.getActive();
  var detail = [], records = [];
  list.forEach(function (r) { detail = detail.concat(r.detail); records.push(r.record); });
  if (detail.length) {
    var d = ss.getSheetByName(SHEET.DETAIL);
    d.getRange(d.getLastRow() + 1, 1, detail.length, DETAIL_HEADERS.length).setValues(detail);
  }
  if (records.length) {
    var s = ss.getSheetByName(SHEET.RECORDS);
    s.getRange(s.getLastRow() + 1, 1, records.length, RECORD_HEADERS.length).setValues(records);
  }
}

/** 清空「明細」「每次紀錄」，從所有版本的表單回覆重新計算（觸發器漏跑或資料被誤改時用） */
function rebuildAll() {
  var ss = SpreadsheetApp.getActive();
  ensureDataSheets_(ss);
  [SHEET.DETAIL, SHEET.RECORDS].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  });
  // 題目文字以目前 Config.js 為準；停用的舊題目要留在 Config 並標 retired: true，舊回覆才算得到。
  var history = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP.FORM_HISTORY) || '[]');
  var all = [];
  history.forEach(function (formId) {
    FormApp.openById(formId).getResponses().forEach(function (resp) {
      all.push(buildRows(parseResponse_(resp, formId), { ROLES: ROLES, PHASES: PHASES, ITEMS: ITEMS }));
    });
  });
  all.sort(function (a, b) { return a.record[1] - b.record[1]; });
  appendRows_(all);
  ss.toast('已重新計算 ' + all.length + ' 筆紀錄', '洗冰檢點');
}
