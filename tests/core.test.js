// 用 Node 內建測試跑：node --test tests/
const test = require('node:test');
const assert = require('node:assert');
const cfg = require('../apps-script/Config.js');
const core = require('../apps-script/Core.js');

function allOk() {
  const a = {};
  cfg.ITEMS.forEach((it) => { a[it.id] = core.CHECK_OK; });
  return a;
}
const base = {
  responseId: 'r1', submittedAt: new Date(2026, 10, 23, 14, 20), version: '11509-2',
  date: '2026-11-23', time: '14:00', filler: '王小明',
  people: { driver: '甲', safety1: '乙', safety2: '丙', commander: '王小明' }, notes: ''
};

test('題目數量與原表一致：48 題、18 個重點', () => {
  assert.strictEqual(cfg.ITEMS.length, 48);
  assert.strictEqual(cfg.ITEMS.filter((i) => i.key).length, 18);
  const ids = cfg.ITEMS.map((i) => i.id);
  assert.strictEqual(new Set(ids).size, ids.length);
});

test('全部做到：48 列明細、達成率 100%', () => {
  const { detail, record } = core.buildRows({ ...base, answers: allOk() }, cfg);
  assert.strictEqual(detail.length, 48);
  assert.strictEqual(detail[0].length, core.DETAIL_HEADERS.length);
  assert.strictEqual(record.length, core.RECORD_HEADERS.length);
  const h = (name) => record[core.RECORD_HEADERS.indexOf(name)];
  assert.strictEqual(h('檢核項數'), 48);
  assert.strictEqual(h('未做到項數'), 0);
  assert.strictEqual(h('達成率'), 1);
  assert.strictEqual(h('時段(時)'), 14);
  assert.strictEqual(h('月份'), '2026-11');
});

test('未做到：人員對應正確、重點計數正確', () => {
  const answers = allOk();
  answers.D01 = core.CHECK_NG; // 重點
  answers.S203 = core.CHECK_NG; // 安全員2第3題（非重點）
  answers.S203 = undefined;
  answers.S204 = core.CHECK_NG;
  const { detail, record } = core.buildRows({ ...base, answers }, cfg);
  const col = (n) => core.DETAIL_HEADERS.indexOf(n);
  const ng = detail.filter((r) => r[col('未做到')] === 1);
  assert.deepStrictEqual(ng.map((r) => [r[col('項目編號')], r[col('人員')], r[col('角色')]]),
    [['D01', '甲', '洗冰車駕駛'], ['S204', '丙', '冰上安全員2']]);
  const h = (name) => record[core.RECORD_HEADERS.indexOf(name)];
  assert.strictEqual(h('檢核項數'), 47); // S203 未作答被略過
  assert.strictEqual(h('未做到項數'), 2);
  assert.strictEqual(h('重點未做到項數'), cfg.ITEMS.find((i) => i.id === 'S204').key ? 2 : 1);
  assert.match(h('未做到項目'), /^D01 經過教練室/);
});

test('週起日為週一', () => {
  const sun = core.weekStart_(new Date(2026, 10, 29)); // 週日
  assert.strictEqual(sun.getDay(), 1);
  assert.strictEqual(sun.getDate(), 23);
});

test('無法辨識的答案會報錯', () => {
  const answers = allOk(); answers.C01 = '亂填';
  assert.throws(() => core.buildRows({ ...base, answers }, cfg), /C01/);
});
