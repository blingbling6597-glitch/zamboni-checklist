/**
 * 純邏輯：把一筆表單回覆轉成「明細」與「每次紀錄」的資料列。
 * 不呼叫任何 Google 服務，可以在 Node 直接測試（見 tests/core.test.js）。
 */

var CHECK_OK = '✔ 有做到';
var CHECK_NG = '✘ 未做到';

var DETAIL_HEADERS = [
  '填報編號', '送出時間', '洗冰日期', '洗冰時間', '時段(時)', '星期', '週起日', '月份',
  '表單版本', '角色', '人員', '階段', '項目編號', '檢核項目', '重點', '結果', '有做到', '未做到', '填表人'
];

var RECORD_HEADERS = [
  '填報編號', '送出時間', '洗冰日期', '洗冰時間', '時段(時)', '週起日', '月份', '表單版本',
  '洗冰車駕駛', '冰上安全員1', '冰上安全員2', '值班洗冰指揮官', '填表人',
  '檢核項數', '未做到項數', '達成率', '重點未做到項數', '未做到項目', '備註'
];

var WEEKDAYS = ['日', '一', '二', '三', '四', '五', '六'];

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

/** "2026-11-23" → Date（當地時間 00:00） */
function parseDate_(s) {
  if (s instanceof Date) return new Date(s.getFullYear(), s.getMonth(), s.getDate());
  var m = String(s).match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) throw new Error('無法辨識日期：' + s);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/** "14:00" → 14 */
function parseHour_(s) {
  var m = String(s).match(/(\d{1,2}):(\d{2})/);
  if (!m) throw new Error('無法辨識時間：' + s);
  return Number(m[1]);
}

/** 該週週一 */
function weekStart_(d) {
  var day = d.getDay();
  var diff = day === 0 ? -6 : 1 - day;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + diff);
}

function itemsById_(items) {
  var map = {};
  items.forEach(function (it) { map[it.id] = it; });
  return map;
}

function nameOf_(list, key) {
  for (var i = 0; i < list.length; i++) if (list[i].key === key) return list[i].name;
  return key;
}

/**
 * @param {Object} sub 一筆回覆（已依題目對應表解析）
 *   { responseId, submittedAt: Date, version, date, time, filler,
 *     people: { driver, safety1, safety2, commander },
 *     answers: { D01: '✔ 有做到', ... }, notes: '...' }
 * @param {Object} cfg { ROLES, PHASES, ITEMS }
 * @return {{detail: Array[], record: Array}}
 */
function buildRows(sub, cfg) {
  var date = parseDate_(sub.date);
  var hour = parseHour_(sub.time);
  var ws = weekStart_(date);
  var month = date.getFullYear() + '-' + pad2_(date.getMonth() + 1);
  var byId = itemsById_(cfg.ITEMS);

  var detail = [];
  var ng = [];
  var keyNg = 0;

  cfg.ITEMS.forEach(function (it) {
    var ans = sub.answers[it.id];
    if (ans === undefined || ans === null || ans === '') return; // 該題未作答（理論上必填）
    var ok = String(ans).indexOf('✔') === 0;
    if (!ok && String(ans).indexOf('✘') !== 0) throw new Error(it.id + ' 的答案無法辨識：' + ans);
    if (!ok) {
      ng.push(it.id);
      if (it.key) keyNg++;
    }
    detail.push([
      sub.responseId, sub.submittedAt, date, sub.time, hour, WEEKDAYS[date.getDay()], ws, month,
      sub.version, nameOf_(cfg.ROLES, it.role), sub.people[it.role] || '', nameOf_(cfg.PHASES, it.phase),
      it.id, it.text, it.key ? '是' : '否', ok ? '✔' : '✘', ok ? 1 : 0, ok ? 0 : 1, sub.filler || ''
    ]);
  });

  var total = detail.length;
  var record = [
    sub.responseId, sub.submittedAt, date, sub.time, hour, ws, month, sub.version,
    sub.people.driver || '', sub.people.safety1 || '', sub.people.safety2 || '', sub.people.commander || '',
    sub.filler || '', total, ng.length, total ? (total - ng.length) / total : '', keyNg,
    ng.map(function (id) { return id + ' ' + byId[id].text; }).join('\n'),
    sub.notes || ''
  ];
  return { detail: detail, record: record };
}

if (typeof module !== 'undefined') {
  module.exports = { buildRows: buildRows, DETAIL_HEADERS: DETAIL_HEADERS, RECORD_HEADERS: RECORD_HEADERS, CHECK_OK: CHECK_OK, CHECK_NG: CHECK_NG, weekStart_: weekStart_ };
}
