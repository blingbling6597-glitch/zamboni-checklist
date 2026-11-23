// 執行：node --test tests/core.test.js
const test = require('node:test');
const assert = require('node:assert');
const cfg = require('../apps-script/Config.js');
const core = require('../apps-script/Core.js');
const { RESULT, COL, SCOL } = core;

const itemsOf = (role) => cfg.ITEMS.filter((i) => i.role === role);
function answersFor(role, result = RESULT.OK) {
  const a = {};
  itemsOf(role).forEach((it) => { a[it.id] = { result, reason: '' }; });
  return a;
}
const people = { driver: '甲', safety1: '乙', safety2: '丙', commander: '丁' };
function self(role, answers, extra = {}) {
  return core.buildSelfRows({
    date: '2026-11-24', time: '14:00', role, person: people[role],
    answers: answers || answersFor(role), submittedAt: new Date(2026, 10, 24, 14, 20), version: '11509-2', ...extra
  }, cfg);
}

test('題目數量與原表一致：48 題、18 個重點、編號不重複', () => {
  assert.strictEqual(cfg.ITEMS.length, 48);
  assert.strictEqual(cfg.ITEMS.filter((i) => i.key).length, 18);
  assert.strictEqual(new Set(cfg.ITEMS.map((i) => i.id)).size, 48);
});

test('角色只產生自己的題目，場次代碼依日期＋整點', () => {
  const rows = self('safety2');
  assert.strictEqual(rows.length, itemsOf('safety2').length);
  assert.ok(rows.every((r) => r[COL['角色']] === '冰上安全員2' && r[COL['人員']] === '丙'));
  assert.strictEqual(rows[0][COL['場次']], '2026-11-24 14:00');
  assert.strictEqual(core.sessionId('2026-11-24', '9:00'), '2026-11-24 09:00');
});

test('不可空白：漏勾任何一題就拒絕', () => {
  const a = answersFor('driver');
  delete a.D05;
  assert.throws(() => self('driver', a), /D05 尚未勾選/);
});

test('勾 ❌ 必須填原因；📐 不用', () => {
  const a = answersFor('driver');
  a.D03 = { result: RESULT.NG, reason: '  ' };
  a.D04 = { result: RESULT.REMIND, reason: '' };
  assert.throws(() => self('driver', a), /D03 勾 ❌ 必須填原因/);
  a.D03.reason = '急著倒車沒停';
  const rows = self('driver', a);
  const r3 = rows.find((r) => r[COL['項目編號']] === 'D03');
  const r4 = rows.find((r) => r[COL['項目編號']] === 'D04');
  assert.deepStrictEqual([r3[COL['最終結果']], r3[COL['未執行']], r3[COL['最終原因']]], ['❌', 1, '急著倒車沒停']);
  assert.deepStrictEqual([r4[COL['最終結果']], r4[COL['提醒後執行']], r4[COL['正確']]], ['📐', 1, 0]);
});

test('不認得的結果符號會被拒絕', () => {
  const a = answersFor('commander');
  a.C01 = { result: '✔' };
  assert.throws(() => self('commander', a), /C01 尚未勾選/);
});

test('指揮官覆核：可改結果，修正會被記錄，改成 ❌ 必須有原因', () => {
  const rows = [...self('driver'), ...self('safety1')];
  const ans = {};
  rows.forEach((r) => { ans[r[COL['項目編號']]] = { result: r[COL['自評結果']], reason: r[COL['自評原因']] }; });
  ans.S102 = { result: RESULT.NG, reason: '' };
  assert.throws(() => core.applyReview(rows, { reviewer: '丁', reviewedAt: new Date(), answers: ans }, cfg), /S102 勾 ❌ 必須填原因/);
  ans.S102.reason = '吹哨後未立即到位';
  ans.D07 = { result: RESULT.REMIND, reason: '' };
  core.applyReview(rows, { reviewer: '丁', reviewedAt: new Date(), answers: ans }, cfg);
  const get = (id) => rows.find((r) => r[COL['項目編號']] === id);
  assert.deepStrictEqual([get('S102')[COL['自評結果']], get('S102')[COL['最終結果']], get('S102')[COL['覆核修正']], get('S102')[COL['最終原因']]],
    ['✅', '❌', 1, '吹哨後未立即到位']);
  assert.strictEqual(get('D01')[COL['覆核修正']], 0);
  assert.ok(rows.every((r) => r[COL['覆核者']] === '丁'));
});

test('覆核要有簽認者', () => {
  const rows = self('driver');
  assert.throws(() => core.applyReview(rows, { reviewer: '', answers: {} }, cfg), /覆核簽認/);
});

test('場次彙總：狀態從待填寫→待覆核→已覆核', () => {
  let rows = [...self('driver'), ...self('safety1'), ...self('safety2')];
  let s = core.summarizeSession(rows, {}, cfg);
  assert.strictEqual(s[SCOL['狀態']], core.STATUS.WAIT_FILL);
  assert.strictEqual(s[SCOL['冰上安全員2']], '丙');
  assert.strictEqual(s[SCOL['值班洗冰指揮官']], '');

  const ca = answersFor('commander');
  ca.C02 = { result: RESULT.NG, reason: '未吹哨' };
  rows = rows.concat(self('commander', ca));
  s = core.summarizeSession(rows, {}, cfg);
  assert.strictEqual(s[SCOL['狀態']], core.STATUS.WAIT_REVIEW);
  assert.strictEqual(s[SCOL['檢核項數']], 48);
  assert.strictEqual(s[SCOL['❌']], 1);
  assert.match(s[SCOL['❌項目與原因']], /^C02（丁）.*原因：未吹哨$/);

  const ans = {};
  rows.forEach((r) => { ans[r[COL['項目編號']]] = { result: r[COL['自評結果']], reason: r[COL['自評原因']] }; });
  core.applyReview(rows, { reviewer: '丁', reviewedAt: new Date(), answers: ans }, cfg);
  s = core.summarizeSession(rows, { comment: 'ok' }, cfg);
  assert.strictEqual(s[SCOL['狀態']], core.STATUS.REVIEWED);
  assert.strictEqual(s[SCOL['正確率']], 47 / 48);
  assert.strictEqual(s[SCOL['覆核意見']], 'ok');
});

test('週起日為週一', () => {
  const w = core.weekStart_(new Date(2026, 10, 29));
  assert.deepStrictEqual([w.getDay(), w.getDate()], [1, 23]);
});
