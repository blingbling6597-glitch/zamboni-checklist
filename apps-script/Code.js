/**
 * 試算表選單、手機填寫網頁（doGet）與網頁呼叫的後端 API。
 * 這個專案要綁定在一份 Google 試算表上（擴充功能 → Apps Script），並「部署為網頁應用程式」。
 */

var SHEET = {
  OVERVIEW: '總覽',
  ITEMS: '項目分析',
  PEOPLE: '人員分析',
  REASONS: '❌原因紀錄',
  TREND: '趨勢分析',
  TIME: '時段分析',
  SESSIONS: '場次',
  DETAIL: '明細',
  ROSTER: '人員名單'
};

var CFG = { ROLES: ROLES, PHASES: PHASES, ITEMS: ITEMS };

// ───────────────────────── 試算表選單 ─────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('洗冰檢點')
    .addItem('① 初始化', 'setup')
    .addItem('顯示填寫網址', 'showAppUrl')
    .addSeparator()
    .addItem('重算場次彙總', 'rebuildSessions')
    .addItem('重建分析分頁', 'buildAnalysisSheets')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Taipei');
  ensureDataSheets_(ss);
  buildAnalysisSheets();
  if (readRoster_().length === 0) {
    ss.setActiveSheet(ss.getSheetByName(SHEET.ROSTER));
    SpreadsheetApp.getUi().alert('初始化完成。\n\n下一步：\n1. 在「人員名單」A 欄填入同仁姓名（一格一人）\n2. Apps Script 編輯器右上「部署 → 新增部署作業 → 網頁應用程式」\n3. 選單「洗冰檢點 → 顯示填寫網址」');
  } else {
    showAppUrl();
  }
}

function showAppUrl() {
  var url = ScriptApp.getService().getUrl();
  var body = url
    ? '<p><b>同仁填寫網址</b>（傳到群組，或做成 QR Code 貼在洗冰室）：</p><p><a href="' + url + '" target="_blank">' + url + '</a></p>'
    : '<p>還沒部署網頁應用程式。</p><p>請到 Apps Script 編輯器右上「部署 → 新增部署作業」，類型選「網頁應用程式」，執行身分選「我」，存取權選「任何人」或「機構內任何人」。</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(body).setWidth(520).setHeight(200), '填寫網址');
}

function ensureDataSheets_(ss) {
  var roster = getOrCreateSheet_(ss, SHEET.ROSTER);
  if (roster.getLastRow() === 0) {
    roster.getRange('A1:B1').setValues([['姓名', '說明']]).setFontWeight('bold');
    roster.getRange('B2').setValue('A 欄一格填一位同仁姓名，填寫網頁的姓名選單會即時讀取這裡。');
    roster.setColumnWidth(1, 160).setColumnWidth(2, 480).setFrozenRows(1);
  }

  var d = getOrCreateSheet_(ss, SHEET.DETAIL);
  if (d.getLastRow() === 0) {
    d.getRange(1, 1, 1, DETAIL_HEADERS.length).setValues([DETAIL_HEADERS]).setFontWeight('bold');
    d.setFrozenRows(1);
  }
  // 文字欄：避免「2026-11-24 14:00」「14:00」「2026-11」被自動轉成日期
  ['場次', '洗冰時間', '月份'].forEach(function (h) { colRange_(d, COL[h]).setNumberFormat('@'); });
  ['洗冰日期', '週起日'].forEach(function (h) { colRange_(d, COL[h]).setNumberFormat('yyyy/mm/dd'); });
  ['自評時間', '覆核時間'].forEach(function (h) { colRange_(d, COL[h]).setNumberFormat('yyyy/mm/dd hh:mm'); });

  var s = getOrCreateSheet_(ss, SHEET.SESSIONS);
  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 1, SESSION_HEADERS.length).setValues([SESSION_HEADERS]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  ['場次', '洗冰時間', '月份'].forEach(function (h) { colRange_(s, SCOL[h]).setNumberFormat('@'); });
  ['洗冰日期', '週起日'].forEach(function (h) { colRange_(s, SCOL[h]).setNumberFormat('yyyy/mm/dd'); });
  ['覆核時間', '最後更新'].forEach(function (h) { colRange_(s, SCOL[h]).setNumberFormat('yyyy/mm/dd hh:mm'); });
  colRange_(s, SCOL['正確率']).setNumberFormat('0.0%');
  colRange_(s, SCOL['❌項目與原因']).setWrap(true);
  s.setColumnWidth(SCOL['❌項目與原因'] + 1, 420);
}

function colRange_(sh, idx0) { return sh.getRange(1, idx0 + 1, sh.getMaxRows(), 1); }

function getOrCreateSheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function readRoster_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ROSTER);
  if (!sh || sh.getLastRow() < 2) return [];
  var seen = {};
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (n) { if (!n || seen[n]) return false; seen[n] = true; return true; });
}

// ───────────────────────── 網頁 ─────────────────────────

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('洗冰檢點')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function roleName_(key) {
  for (var i = 0; i < ROLES.length; i++) if (ROLES[i].key === key) return ROLES[i].name;
  return key;
}

function roleKey_(name) {
  for (var i = 0; i < ROLES.length; i++) if (ROLES[i].name === name) return ROLES[i].key;
  return name;
}

/** 網頁開啟時取得題目與名單 */
function apiBootstrap() {
  var now = new Date();
  var hour = now.getHours() + (now.getMinutes() >= 30 ? 1 : 0); // 洗冰在整點，就近帶入
  return {
    version: FORM_VERSION,
    roles: ROLES,
    phases: PHASES,
    items: ITEMS.filter(function (it) { return !it.retired; }),
    roster: readRoster_(),
    safetyNote: SAFETY_NOTE,
    today: Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd'),
    defaultHour: Math.min(hour, 23)
  };
}

/**
 * 角色送出自評
 * payload: { date:'2026-11-24', time:'14:00', role:'driver', person:'王小明', answers:{ D01:{result,reason} }, overwrite:false }
 */
function apiSubmitSelf(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActive();
    var rows = buildSelfRows({
      date: payload.date, time: payload.time, role: payload.role, person: payload.person,
      answers: payload.answers || {}, submittedAt: new Date(), version: FORM_VERSION
    }, CFG);
    var sid = rows[0][COL['場次']];
    var existing = findSessionRows_(ss, sid);

    if (existing.some(function (e) { return e.values[COL['覆核結果']]; })) {
      throw new Error('場次 ' + sid + ' 已由指揮官覆核簽認，無法再修改。如需更正請聯絡當班指揮官。');
    }
    var mine = existing.filter(function (e) { return e.values[COL['角色']] === roleName_(payload.role); });
    if (mine.length && !payload.overwrite) {
      return { needConfirm: true, message: roleName_(payload.role) + ' 在場次 ' + sid + ' 已由「' + mine[0].values[COL['人員']] + '」填過，要用這次的內容覆蓋嗎？' };
    }

    var d = ss.getSheetByName(SHEET.DETAIL);
    if (mine.length) {
      // 同一角色重填：先清掉舊列再附加（舊列不一定連續）
      deleteRows_(d, mine.map(function (e) { return e.row; }));
    }
    d.getRange(d.getLastRow() + 1, 1, rows.length, DETAIL_HEADERS.length).setValues(rows);
    updateSession_(ss, sid, null);
    return { ok: true, sessionId: sid };
  } finally {
    lock.releaseLock();
  }
}

/** 覆核用：近幾天的場次清單 */
function apiListSessions() {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(SHEET.SESSIONS);
  var n = s.getLastRow() - 1;
  if (n < 1) return [];
  var start = Math.max(2, s.getLastRow() - 199);
  var values = s.getRange(start, 1, s.getLastRow() - start + 1, SESSION_HEADERS.length).getValues();
  var since = new Date(); since.setDate(since.getDate() - 7);
  return values
    .filter(function (v) { return v[SCOL['洗冰日期']] instanceof Date && v[SCOL['洗冰日期']] >= new Date(since.getFullYear(), since.getMonth(), since.getDate()); })
    .map(function (v) {
      return {
        id: v[SCOL['場次']], status: v[SCOL['狀態']], filled: v[SCOL['已填角色數']],
        ng: v[SCOL['❌']], remind: v[SCOL['📐']], reviewer: v[SCOL['覆核者']],
        people: ROLES.map(function (r) { return { role: r.name, person: v[SCOL[r.name]] }; })
      };
    })
    .sort(function (a, b) { return a.id < b.id ? 1 : -1; });
}

/** 覆核用：一個場次的所有角色結果 */
function apiGetSession(sid) {
  var ss = SpreadsheetApp.getActive();
  var rows = findSessionRows_(ss, sid);
  var roles = ROLES.map(function (r) {
    var mine = rows.filter(function (e) { return e.values[COL['角色']] === r.name; });
    return {
      key: r.key, name: r.name,
      person: mine.length ? mine[0].values[COL['人員']] : '',
      submittedAt: mine.length ? Utilities.formatDate(mine[0].values[COL['自評時間']], 'Asia/Taipei', 'MM/dd HH:mm') : '',
      items: mine.map(function (e) {
        var v = e.values;
        return {
          id: v[COL['項目編號']], text: v[COL['檢核項目']], key: v[COL['重點']] === '是', phase: v[COL['階段']],
          self: v[COL['自評結果']], selfReason: v[COL['自評原因']],
          review: v[COL['覆核結果']], reviewReason: v[COL['覆核原因']]
        };
      })
    };
  });
  var srow = findSessionSummaryRow_(ss, sid);
  var reviewed = rows.length > 0 && rows.every(function (e) { return e.values[COL['覆核結果']]; });
  return {
    id: sid, roles: roles, reviewed: reviewed,
    reviewer: reviewed ? rows[0].values[COL['覆核者']] : '',
    comment: srow ? srow.values[SCOL['覆核意見']] : ''
  };
}

/**
 * 指揮官覆核簽認
 * payload: { sessionId, reviewer, answers:{ id:{result,reason} }, comment, confirmMissing }
 */
function apiSubmitReview(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActive();
    var found = findSessionRows_(ss, payload.sessionId);
    if (!found.length) throw new Error('找不到場次 ' + payload.sessionId);

    var filled = {};
    found.forEach(function (e) { filled[e.values[COL['角色']]] = true; });
    var missing = ROLES.filter(function (r) { return !filled[r.name]; }).map(function (r) { return r.name; });
    if (missing.length && !payload.confirmMissing) {
      return { needConfirm: true, message: '以下角色尚未填寫：' + missing.join('、') + '。\n仍要先覆核簽認嗎？（簽認後未填的角色就不能再填）' };
    }

    var rows = found.map(function (e) { return e.values; });
    applyReview(rows, { reviewer: payload.reviewer, reviewedAt: new Date(), answers: payload.answers || {} }, CFG);

    var d = ss.getSheetByName(SHEET.DETAIL);
    contiguousBlocks_(found).forEach(function (block) {
      d.getRange(block[0].row, 1, block.length, DETAIL_HEADERS.length)
        .setValues(block.map(function (e) { return e.values; }));
    });
    updateSession_(ss, payload.sessionId, String(payload.comment || '').trim());
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────── 資料存取 ─────────────────────────

/** 找出某場次在「明細」的所有列：[{ row: 列號, values: [...] }] */
function findSessionRows_(ss, sid) {
  var d = ss.getSheetByName(SHEET.DETAIL);
  if (d.getLastRow() < 2) return [];
  var cells = d.getRange(2, COL['場次'] + 1, d.getLastRow() - 1, 1)
    .createTextFinder(sid).matchEntireCell(true).findAll();
  var rowNums = cells.map(function (c) { return c.getRow(); }).sort(function (a, b) { return a - b; });
  var out = [];
  contiguousBlocks_(rowNums.map(function (r) { return { row: r }; })).forEach(function (block) {
    var vals = d.getRange(block[0].row, 1, block.length, DETAIL_HEADERS.length).getValues();
    block.forEach(function (b, i) { out.push({ row: b.row, values: vals[i] }); });
  });
  return out;
}

/** 依列號把連續的列分成一組，減少讀寫次數 */
function contiguousBlocks_(entries) {
  var blocks = [];
  entries.slice().sort(function (a, b) { return a.row - b.row; }).forEach(function (e) {
    var last = blocks[blocks.length - 1];
    if (last && last[last.length - 1].row === e.row - 1) last.push(e); else blocks.push([e]);
  });
  return blocks;
}

function deleteRows_(sheet, rowNums) {
  var blocks = contiguousBlocks_(rowNums.map(function (r) { return { row: r }; })).reverse();
  blocks.forEach(function (b) { sheet.deleteRows(b[0].row, b.length); });
}

function findSessionSummaryRow_(ss, sid) {
  var s = ss.getSheetByName(SHEET.SESSIONS);
  if (s.getLastRow() < 2) return null;
  var cell = s.getRange(2, 1, s.getLastRow() - 1, 1).createTextFinder(sid).matchEntireCell(true).findNext();
  if (!cell) return null;
  return { row: cell.getRow(), values: s.getRange(cell.getRow(), 1, 1, SESSION_HEADERS.length).getValues()[0] };
}

/** 重算一個場次在「場次」分頁的彙總列；comment 為 null 時保留原本的覆核意見 */
function updateSession_(ss, sid, comment) {
  var rows = findSessionRows_(ss, sid).map(function (e) { return e.values; });
  var s = ss.getSheetByName(SHEET.SESSIONS);
  var existing = findSessionSummaryRow_(ss, sid);
  if (!rows.length) {
    if (existing) s.deleteRow(existing.row);
    return;
  }
  var keepComment = comment === null ? (existing ? existing.values[SCOL['覆核意見']] : '') : comment;
  var summary = summarizeSession(rows, { comment: keepComment, updatedAt: new Date() }, CFG);
  if (existing) s.getRange(existing.row, 1, 1, SESSION_HEADERS.length).setValues([summary]);
  else s.getRange(s.getLastRow() + 1, 1, 1, SESSION_HEADERS.length).setValues([summary]);
}

/** 從「明細」重算整張「場次」分頁（資料被手動修改後用） */
function rebuildSessions() {
  var ss = SpreadsheetApp.getActive();
  ensureDataSheets_(ss);
  var d = ss.getSheetByName(SHEET.DETAIL);
  var s = ss.getSheetByName(SHEET.SESSIONS);
  var comments = {};
  if (s.getLastRow() > 1) {
    s.getRange(2, 1, s.getLastRow() - 1, SESSION_HEADERS.length).getValues().forEach(function (v) {
      comments[v[SCOL['場次']]] = v[SCOL['覆核意見']];
    });
    s.getRange(2, 1, s.getLastRow() - 1, SESSION_HEADERS.length).clearContent();
  }
  if (d.getLastRow() < 2) return;
  var groups = {}, order = [];
  d.getRange(2, 1, d.getLastRow() - 1, DETAIL_HEADERS.length).getValues().forEach(function (v) {
    var id = v[COL['場次']];
    if (!id) return;
    if (!groups[id]) { groups[id] = []; order.push(id); }
    groups[id].push(v);
  });
  order.sort();
  var out = order.map(function (id) { return summarizeSession(groups[id], { comment: comments[id] || '', updatedAt: new Date() }, CFG); });
  if (out.length) s.getRange(2, 1, out.length, SESSION_HEADERS.length).setValues(out);
  ss.toast('已重算 ' + out.length + ' 個場次', '洗冰檢點');
}
