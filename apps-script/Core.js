/**
 * 純邏輯（不呼叫 Google 服務），可在 Node 測試：node --test tests/core.test.js
 *
 * 結果三級：
 *   ✅ 單項有正確執行
 *   📐 單項不熟，但現場提醒後有執行
 *   ❌ 單項沒執行，或溝通提醒後仍無正確執行（必須填原因）
 */

var RESULT = { OK: '✅', REMIND: '📐', NG: '❌' };
var RESULT_LIST = [RESULT.OK, RESULT.REMIND, RESULT.NG];

var DETAIL_HEADERS = [
  '場次', '洗冰日期', '洗冰時間', '時段(時)', '星期', '週起日', '月份', '表單版本',
  '角色', '人員', '階段', '項目編號', '檢核項目', '重點',
  '自評結果', '自評原因', '自評時間',
  '覆核結果', '覆核原因', '覆核者', '覆核時間',
  '最終結果', '最終原因', '正確', '提醒後執行', '未執行', '覆核修正'
];
// 欄位索引（0 起算），程式裡用名字取欄位，避免數錯
var COL = {};
DETAIL_HEADERS.forEach(function (h, i) { COL[h] = i; });

var SESSION_HEADERS = [
  '場次', '洗冰日期', '洗冰時間', '時段(時)', '週起日', '月份', '表單版本',
  '洗冰車駕駛', '冰上安全員1', '冰上安全員2', '值班洗冰指揮官',
  '已填角色數', '狀態', '覆核者', '覆核時間',
  '檢核項數', '✅', '📐', '❌', '正確率', '覆核修正項數', '❌項目與原因', '覆核意見', '最後更新'
];
var SCOL = {};
SESSION_HEADERS.forEach(function (h, i) { SCOL[h] = i; });

var STATUS = { WAIT_FILL: '待填寫', WAIT_REVIEW: '待覆核', REVIEWED: '已覆核' };

var WEEKDAYS = ['日', '一', '二', '三', '四', '五', '六'];

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function parseDate_(s) {
  if (s instanceof Date) return new Date(s.getFullYear(), s.getMonth(), s.getDate());
  var m = String(s).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) throw new Error('日期格式錯誤：' + s);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function formatDate_(d) {
  return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
}

function parseHour_(s) {
  var m = String(s).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) throw new Error('時間格式錯誤：' + s);
  return Number(m[1]);
}

function weekStart_(d) {
  var day = d.getDay();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + (day === 0 ? -6 : 1 - day));
}

/** 場次代碼：同一天同一整點的洗冰，四個角色的填報靠它湊在一起 */
function sessionId(date, time) {
  return formatDate_(parseDate_(date)) + ' ' + pad2_(parseHour_(time)) + ':00';
}

function sessionMeta(date, time) {
  var d = parseDate_(date);
  return {
    id: sessionId(date, time),
    date: d,
    time: pad2_(parseHour_(time)) + ':00',
    hour: parseHour_(time),
    weekday: WEEKDAYS[d.getDay()],
    weekStart: weekStart_(d),
    month: d.getFullYear() + '-' + pad2_(d.getMonth() + 1)
  };
}

function nameOf_(list, key) {
  for (var i = 0; i < list.length; i++) if (list[i].key === key) return list[i].name;
  return key;
}

function activeItemsOf(role, cfg) {
  return cfg.ITEMS.filter(function (it) { return it.role === role && !it.retired; });
}

/**
 * 檢查一份結果是否完整：每題都要有 ✅📐❌，❌ 一定要有原因。
 * @param {Array} items 要檢查的題目
 * @param {Object} answers { 項目編號: { result: '✅', reason: '' } }
 * @return {string[]} 錯誤訊息（空陣列代表通過）
 */
function validateAnswers(items, answers) {
  var errors = [];
  items.forEach(function (it) {
    var a = answers[it.id];
    if (!a || RESULT_LIST.indexOf(a.result) < 0) {
      errors.push(it.id + ' 尚未勾選');
    } else if (a.result === RESULT.NG && !String(a.reason || '').trim()) {
      errors.push(it.id + ' 勾 ❌ 必須填原因');
    }
  });
  return errors;
}

/**
 * 角色自評 → 明細列
 * @param {Object} sub { date, time, role, person, answers, submittedAt: Date, version }
 */
function buildSelfRows(sub, cfg) {
  if (!cfg.ROLES.some(function (r) { return r.key === sub.role; })) throw new Error('角色錯誤：' + sub.role);
  if (!String(sub.person || '').trim()) throw new Error('請選擇姓名');
  var items = activeItemsOf(sub.role, cfg);
  var errors = validateAnswers(items, sub.answers);
  if (errors.length) throw new Error('資料不完整：\n' + errors.join('\n'));

  var s = sessionMeta(sub.date, sub.time);
  return items.map(function (it) {
    var a = sub.answers[it.id];
    var row = new Array(DETAIL_HEADERS.length);
    row[COL['場次']] = s.id;
    row[COL['洗冰日期']] = s.date;
    row[COL['洗冰時間']] = s.time;
    row[COL['時段(時)']] = s.hour;
    row[COL['星期']] = s.weekday;
    row[COL['週起日']] = s.weekStart;
    row[COL['月份']] = s.month;
    row[COL['表單版本']] = sub.version;
    row[COL['角色']] = nameOf_(cfg.ROLES, it.role);
    row[COL['人員']] = sub.person;
    row[COL['階段']] = nameOf_(cfg.PHASES, it.phase);
    row[COL['項目編號']] = it.id;
    row[COL['檢核項目']] = it.text;
    row[COL['重點']] = it.key ? '是' : '否';
    row[COL['自評結果']] = a.result;
    row[COL['自評原因']] = String(a.reason || '').trim();
    row[COL['自評時間']] = sub.submittedAt;
    row[COL['覆核結果']] = '';
    row[COL['覆核原因']] = '';
    row[COL['覆核者']] = '';
    row[COL['覆核時間']] = '';
    return finalize_(row);
  });
}

/** 依自評／覆核欄位算出最終結果與 0/1 指標欄 */
function finalize_(row) {
  var reviewed = row[COL['覆核結果']];
  var final = reviewed || row[COL['自評結果']];
  row[COL['最終結果']] = final;
  row[COL['最終原因']] = reviewed ? row[COL['覆核原因']] : row[COL['自評原因']];
  row[COL['正確']] = final === RESULT.OK ? 1 : 0;
  row[COL['提醒後執行']] = final === RESULT.REMIND ? 1 : 0;
  row[COL['未執行']] = final === RESULT.NG ? 1 : 0;
  row[COL['覆核修正']] = reviewed && reviewed !== row[COL['自評結果']] ? 1 : 0;
  return row;
}

/**
 * 指揮官覆核：把覆核結果套進該場次的明細列（就地修改並回傳）
 * @param {Array[]} rows 該場次所有明細列
 * @param {Object} review { reviewer, reviewedAt: Date, answers: { 項目編號: { result, reason } } }
 *   answers 要包含該場次每一題（畫面上預設帶自評結果，指揮官只改需要改的）
 */
function applyReview(rows, review, cfg) {
  if (!String(review.reviewer || '').trim()) throw new Error('請選擇覆核簽認的指揮官');
  var byId = {};
  cfg.ITEMS.forEach(function (it) { byId[it.id] = it; });
  var items = rows.map(function (r) { return byId[r[COL['項目編號']]] || { id: r[COL['項目編號']] }; });
  var errors = validateAnswers(items, review.answers);
  if (errors.length) throw new Error('覆核資料不完整：\n' + errors.join('\n'));

  rows.forEach(function (row) {
    var a = review.answers[row[COL['項目編號']]];
    row[COL['覆核結果']] = a.result;
    row[COL['覆核原因']] = String(a.reason || '').trim();
    row[COL['覆核者']] = review.reviewer;
    row[COL['覆核時間']] = review.reviewedAt;
    finalize_(row);
  });
  return rows;
}

/** 由該場次明細列算出「場次」分頁的一列 */
function summarizeSession(rows, extra, cfg) {
  var first = rows[0];
  var people = {};
  rows.forEach(function (r) { people[r[COL['角色']]] = r[COL['人員']]; });
  var roleCount = Object.keys(people).length;
  var reviewed = rows.length > 0 && rows.every(function (r) { return r[COL['覆核結果']]; });
  var count = function (k) { return rows.reduce(function (s, r) { return s + r[COL[k]]; }, 0); };
  var ngList = rows.filter(function (r) { return r[COL['最終結果']] === RESULT.NG; }).map(function (r) {
    return r[COL['項目編號']] + '（' + r[COL['人員']] + '）' + r[COL['檢核項目']] + '｜原因：' + r[COL['最終原因']];
  });

  var row = new Array(SESSION_HEADERS.length);
  row[SCOL['場次']] = first[COL['場次']];
  row[SCOL['洗冰日期']] = first[COL['洗冰日期']];
  row[SCOL['洗冰時間']] = first[COL['洗冰時間']];
  row[SCOL['時段(時)']] = first[COL['時段(時)']];
  row[SCOL['週起日']] = first[COL['週起日']];
  row[SCOL['月份']] = first[COL['月份']];
  row[SCOL['表單版本']] = first[COL['表單版本']];
  cfg.ROLES.forEach(function (role) { row[SCOL[role.name]] = people[role.name] || ''; });
  row[SCOL['已填角色數']] = roleCount;
  row[SCOL['狀態']] = reviewed ? STATUS.REVIEWED : (roleCount < cfg.ROLES.length ? STATUS.WAIT_FILL : STATUS.WAIT_REVIEW);
  row[SCOL['覆核者']] = reviewed ? first[COL['覆核者']] : '';
  row[SCOL['覆核時間']] = reviewed ? first[COL['覆核時間']] : '';
  row[SCOL['檢核項數']] = rows.length;
  row[SCOL['✅']] = count('正確');
  row[SCOL['📐']] = count('提醒後執行');
  row[SCOL['❌']] = count('未執行');
  row[SCOL['正確率']] = rows.length ? count('正確') / rows.length : '';
  row[SCOL['覆核修正項數']] = count('覆核修正');
  row[SCOL['❌項目與原因']] = ngList.join('\n');
  row[SCOL['覆核意見']] = (extra && extra.comment) || '';
  row[SCOL['最後更新']] = (extra && extra.updatedAt) || '';
  return row;
}

if (typeof module !== 'undefined') {
  module.exports = {
    RESULT: RESULT, STATUS: STATUS, DETAIL_HEADERS: DETAIL_HEADERS, SESSION_HEADERS: SESSION_HEADERS, COL: COL, SCOL: SCOL,
    sessionId: sessionId, validateAnswers: validateAnswers, buildSelfRows: buildSelfRows,
    applyReview: applyReview, summarizeSession: summarizeSession, weekStart_: weekStart_
  };
}
