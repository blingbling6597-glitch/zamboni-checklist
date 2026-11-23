/**
 * 分析分頁：全部用公式即時計算，同仁送出或指揮官簽認後就會更新。
 * 「最終結果」= 指揮官覆核後的結果；還沒覆核的場次先用自評結果。
 * 選單「重建分析分頁」會清掉並重畫這些分頁（不影響明細與場次）。
 *
 * 明細：A場次 B洗冰日期 C洗冰時間 D時段 E星期 F週起日 G月份 H表單版本 I角色 J人員 K階段 L項目編號
 *       M檢核項目 N重點 O自評結果 P自評原因 Q自評時間 R覆核結果 S覆核原因 T覆核者 U覆核時間
 *       V最終結果 W最終原因 X正確 Y提醒後執行 Z未執行 AA覆核修正
 * 場次：A場次 B洗冰日期 C洗冰時間 D時段 E週起日 F月份 G表單版本 H駕駛 I安全員1 J安全員2 K指揮官
 *       L已填角色數 M狀態 N覆核者 O覆核時間 P檢核項數 Q✅ R📐 S❌ T正確率 U覆核修正項數 V❌項目與原因 W覆核意見 X最後更新
 */

var D = "'明細'!";   // 公式裡引用「明細」分頁
var S = "'場次'!";   // 公式裡引用「場次」分頁
var DQ = D + 'A1:AA';
var SQ = S + 'A1:X';

function buildAnalysisSheets() {
  var ss = SpreadsheetApp.getActive();
  ensureDataSheets_(ss);
  buildItems_(resetSheet_(ss, SHEET.ITEMS));
  buildPeople_(resetSheet_(ss, SHEET.PEOPLE));
  buildReasons_(resetSheet_(ss, SHEET.REASONS));
  buildTrend_(resetSheet_(ss, SHEET.TREND));
  buildTime_(resetSheet_(ss, SHEET.TIME));
  buildOverview_(resetSheet_(ss, SHEET.OVERVIEW));

  [SHEET.OVERVIEW, SHEET.ITEMS, SHEET.PEOPLE, SHEET.REASONS, SHEET.TREND, SHEET.TIME, SHEET.SESSIONS, SHEET.DETAIL, SHEET.ROSTER]
    .forEach(function (name, i) {
      var sh = ss.getSheetByName(name);
      if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
    });
  ss.setActiveSheet(ss.getSheetByName(SHEET.OVERVIEW));
  var empty = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (empty && empty.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(empty);
}

function resetSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) {
    sh.getCharts().forEach(function (c) { sh.removeChart(c); });
    var f = sh.getFilter(); if (f) f.remove();
    sh.clear();
    sh.clearConditionalFormatRules();
  } else {
    sh = ss.insertSheet(name);
  }
  return sh;
}

function title_(sh, a1, text) { sh.getRange(a1).setValue(text).setFontWeight('bold').setFontSize(12); }
function header_(range) { range.setFontWeight('bold').setBackground('#e8eef7'); }
function q_(range, query) { return '=IFERROR(QUERY(' + range + ',"' + query + '",1),"（尚無資料）")'; }

// ───────────────────────── 總覽 ─────────────────────────

function buildOverview_(sh) {
  sh.getRange('A1').setValue('洗冰檢點　總覽').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('✅ 正確執行　📐 不熟但提醒後有執行　❌ 沒執行或提醒後仍不正確').setFontColor('#6b7785');
  var kpi = [
    ['累計場次', '=COUNTA(' + S + 'A2:A)', '0'],
    ['已覆核', '=COUNTIF(' + S + 'M2:M,"已覆核")', '0'],
    ['待覆核', '=COUNTIF(' + S + 'M2:M,"待覆核")', '0'],
    ['待填寫（角色未到齊）', '=COUNTIF(' + S + 'M2:M,"待填寫")', '0'],
    ['整體 ✅ 正確率', '=IFERROR(AVERAGE(' + D + 'X2:X),"")', '0.0%'],
    ['整體 📐 提醒後執行率', '=IFERROR(AVERAGE(' + D + 'Y2:Y),"")', '0.0%'],
    ['整體 ❌ 未執行率', '=IFERROR(AVERAGE(' + D + 'Z2:Z),"")', '0.0%'],
    ['重點項目 ✅ 正確率', '=IFERROR(AVERAGEIF(' + D + 'N2:N,"是",' + D + 'X2:X),"")', '0.0%'],
    ['近 7 天 ✅ 正確率', '=IFERROR(AVERAGEIFS(' + D + 'X2:X,' + D + 'B2:B,">="&(TODAY()-6)),"")', '0.0%'],
    ['本月 ✅ 正確率', '=IFERROR(AVERAGEIF(' + D + 'G2:G,TEXT(TODAY(),"yyyy-mm"),' + D + 'X2:X),"")', '0.0%'],
    ['覆核修正項次（自評與指揮官不同）', '=SUM(' + D + 'AA2:AA)', '0']
  ];
  sh.getRange(4, 1, kpi.length, 2).setValues(kpi.map(function (k) { return [k[0], k[1]]; }));
  kpi.forEach(function (k, i) { sh.getRange(4 + i, 2).setNumberFormat(k[2]); });
  sh.getRange(4, 1, kpi.length, 1).setFontWeight('bold');
  sh.getRange(4, 2, kpi.length, 1).setFontSize(14).setHorizontalAlignment('right');

  title_(sh, 'D3', '尚未完成的場次（待填寫／待覆核）');
  sh.getRange('D4').setFormula(q_(SQ, "select A, M, L, H, I, J, K where A is not null and M <> '已覆核' order by A desc limit 20 " +
    "label A '場次', M '狀態', L '已填角色', H '駕駛', I '安全員1', J '安全員2', K '指揮官'"));

  title_(sh, 'A17', '最需要加強的 10 個項目（❌＋📐 次數）');
  sh.getRange('A18:F18').setValues([['項目編號', '角色', '檢核項目', '📐次數', '❌次數', '📐＋❌']]);
  header_(sh.getRange('A18:F18'));
  var I = "'項目分析'!";
  sh.getRange('A19').setFormula('=IFERROR(ARRAYFORMULA(SORTN(FILTER({' + I + 'A2:A,' + I + 'B2:B,' + I + 'E2:E,' + I + 'H2:H,' + I + 'I2:I,' + I + 'H2:H+' + I + 'I2:I},' +
    I + 'H2:H+' + I + 'I2:I>0),10,0,6,FALSE)),"目前沒有 📐 或 ❌ 紀錄")');

  title_(sh, 'A31', '最近 10 個場次');
  sh.getRange('A32').setFormula(q_(SQ, "select A, M, N, Q, R, S, T, V where A is not null order by A desc limit 10 " +
    "label A '場次', M '狀態', N '覆核者', Q '✅', R '📐', S '❌', T '正確率', V '❌項目與原因' format T '0.0%'"));

  header_(sh.getRange('D4:J4'));
  header_(sh.getRange('A32:H32'));
  sh.setColumnWidth(1, 150).setColumnWidth(2, 90).setColumnWidth(3, 360).setColumnWidth(8, 420);
  sh.getRange('A33:H42').setWrap(true).setVerticalAlignment('top');
}

// ───────────────────────── 項目分析 ─────────────────────────

function buildItems_(sh) {
  var headers = ['項目編號', '角色', '階段', '重點', '檢核項目', '檢核次數', '✅次數', '📐次數', '❌次數',
    '✅正確率', '📐率', '❌率', '覆核修正次數', '狀態'];
  var roleName = {}; ROLES.forEach(function (r) { roleName[r.key] = r.name; });
  var phaseName = {}; PHASES.forEach(function (p) { phaseName[p.key] = p.name; });
  var L = D + 'L2:L';
  var rows = ITEMS.map(function (it, i) {
    var r = i + 2;
    return [
      it.id, roleName[it.role], phaseName[it.phase], it.key ? '是' : '否', it.text,
      '=COUNTIF(' + L + ',A' + r + ')',
      '=SUMIF(' + L + ',A' + r + ',' + D + 'X2:X)',
      '=SUMIF(' + L + ',A' + r + ',' + D + 'Y2:Y)',
      '=SUMIF(' + L + ',A' + r + ',' + D + 'Z2:Z)',
      '=IF(F' + r + '=0,"",G' + r + '/F' + r + ')',
      '=IF(F' + r + '=0,"",H' + r + '/F' + r + ')',
      '=IF(F' + r + '=0,"",I' + r + '/F' + r + ')',
      '=SUMIF(' + L + ',A' + r + ',' + D + 'AA2:AA)',
      it.retired ? '已停用' : '使用中'
    ];
  });
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  header_(sh.getRange(1, 1, 1, headers.length));
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.getRange(2, 10, rows.length, 3).setNumberFormat('0.0%');
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 110).setColumnWidth(3, 140).setColumnWidth(5, 420);
  sh.getRange(2, 5, rows.length, 1).setWrap(true);
  sh.getRange(1, 1, rows.length + 1, headers.length).createFilter();

  var rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMaxpointWithValue('#e06666', SpreadsheetApp.InterpolationType.NUMBER, '0.3')
      .setRanges([sh.getRange(2, 12, rows.length, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMaxpointWithValue('#f6b26b', SpreadsheetApp.InterpolationType.NUMBER, '0.3')
      .setRanges([sh.getRange(2, 11, rows.length, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('是').setBackground('#f8cbad')
      .setRanges([sh.getRange(2, 4, rows.length, 1)]).build()
  ];
  sh.setConditionalFormatRules(rules);

  sh.insertChart(sh.newChart().asBarChart()
    .addRange(sh.getRange(1, 1, rows.length + 1, 1))
    .addRange(sh.getRange(1, 11, rows.length + 1, 2))
    .setStacked()
    .setColors(['#f6b26b', '#e06666'])
    .setOption('title', '各項目 📐 提醒後執行率＋❌ 未執行率')
    .setOption('hAxis', { format: 'percent' })
    .setOption('height', 900).setOption('width', 560)
    .setPosition(1, 16, 0, 0).build());
}

// ───────────────────────── 人員分析 ─────────────────────────

function buildPeople_(sh) {
  title_(sh, 'A1', '人員別（不分角色）');
  sh.getRange('A2').setFormula(q_(DQ, "select J, count(X), sum(X), sum(Y), sum(Z), avg(X), sum(AA) where J is not null and J <> '' group by J order by avg(X) asc " +
    "label J '人員', count(X) '檢核項次', sum(X) '✅', sum(Y) '📐', sum(Z) '❌', avg(X) '✅正確率', sum(AA) '被覆核修正' format avg(X) '0.0%'"));
  header_(sh.getRange('A2:G2'));

  title_(sh, 'I1', '人員 × 角色');
  sh.getRange('I2').setFormula(q_(DQ, "select J, I, count(X), sum(Y), sum(Z), avg(X) where J is not null and J <> '' group by J, I order by J " +
    "label J '人員', I '角色', count(X) '檢核項次', sum(Y) '📐', sum(Z) '❌', avg(X) '✅正確率' format avg(X) '0.0%'"));
  header_(sh.getRange('I2:N2'));

  title_(sh, 'P1', '人員重點項目');
  sh.getRange('P2').setFormula(q_(DQ, "select J, count(X), sum(Y), sum(Z), avg(X) where J is not null and J <> '' and N = '是' group by J order by avg(X) asc " +
    "label J '人員', count(X) '重點項次', sum(Y) '📐', sum(Z) '❌', avg(X) '重點正確率' format avg(X) '0.0%'"));
  header_(sh.getRange('P2:T2'));

  title_(sh, 'V1', '各人員需加強的項目（📐／❌，次數多的在前）');
  sh.getRange('V2').setFormula(q_(DQ, "select J, L, M, sum(Y), sum(Z) where X = 0 group by J, L, M order by J, sum(Z) desc, sum(Y) desc " +
    "label J '人員', L '編號', M '檢核項目', sum(Y) '📐', sum(Z) '❌'"));
  header_(sh.getRange('V2:Z2'));
  sh.setColumnWidth(24, 400);
  sh.setFrozenRows(2);
}

// ───────────────────────── ❌ 原因紀錄 ─────────────────────────

function buildReasons_(sh) {
  title_(sh, 'A1', '所有 ❌ 項目與原因（新的在上，原因以指揮官覆核後為準）');
  sh.getRange('A2').setFormula(q_(DQ, "select A, I, J, L, M, W, O, R, T where V = '❌' order by A desc " +
    "label A '場次', I '角色', J '人員', L '編號', M '檢核項目', W '原因', O '自評', R '覆核', T '覆核者'"));
  header_(sh.getRange('A2:I2'));
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 140).setColumnWidth(5, 360).setColumnWidth(6, 300);
  sh.getRange('E3:F2000').setWrap(true).setVerticalAlignment('top');
}

// ───────────────────────── 趨勢分析 ─────────────────────────

function buildTrend_(sh) {
  var cols = "count(A), sum(Q), sum(R), sum(S), avg(T)";
  var labels = "count(A) '場次數', sum(Q) '✅', sum(R) '📐', sum(S) '❌', avg(T) '✅正確率'";

  title_(sh, 'A1', '每日');
  sh.getRange('A2').setFormula(q_(SQ, "select B, " + cols + " where A is not null group by B order by B label B '日期', " + labels + " format B 'yyyy/mm/dd', avg(T) '0.0%'"));
  header_(sh.getRange('A2:F2'));

  title_(sh, 'H1', '每週（週一起）');
  sh.getRange('H2').setFormula(q_(SQ, "select E, " + cols + " where A is not null group by E order by E label E '週起日', " + labels + " format E 'yyyy/mm/dd', avg(T) '0.0%'"));
  header_(sh.getRange('H2:M2'));

  title_(sh, 'O1', '每月');
  sh.getRange('O2').setFormula(q_(SQ, "select F, " + cols + " where A is not null group by F order by F label F '月份', " + labels + " format avg(T) '0.0%'"));
  header_(sh.getRange('O2:T2'));

  title_(sh, 'V1', '每月 × 角色 ✅正確率');
  sh.getRange('V2').setFormula(q_(DQ, "select G, avg(X) where G is not null and G <> '' group by G pivot I label G '月份'"));
  sh.getRange('W3:Z200').setNumberFormat('0.0%');
  header_(sh.getRange('V2:Z2'));
  sh.setFrozenRows(2);

  sh.insertChart(sh.newChart().asLineChart()
    .addRange(sh.getRange('A2:A400')).addRange(sh.getRange('F2:F400'))
    .setOption('title', '每日 ✅ 正確率').setOption('legend', { position: 'none' })
    .setOption('vAxis', { format: 'percent', viewWindow: { max: 1 } })
    .setPosition(8, 22, 0, 0).setOption('width', 620).setOption('height', 300).build());
  sh.insertChart(sh.newChart().asColumnChart()
    .addRange(sh.getRange('H2:H200')).addRange(sh.getRange('J2:L200'))
    .setStacked().setColors(['#93c47d', '#f6b26b', '#e06666'])
    .setOption('title', '每週 ✅📐❌ 項次')
    .setPosition(24, 22, 0, 0).setOption('width', 620).setOption('height', 300).build());
}

// ───────────────────────── 時段分析 ─────────────────────────

function buildTime_(sh) {
  title_(sh, 'A1', '洗冰時段（整點）');
  sh.getRange('A2').setFormula(q_(SQ, "select D, count(A), sum(R), sum(S), avg(T) where A is not null group by D order by D " +
    "label D '時段(時)', count(A) '場次數', sum(R) '📐', sum(S) '❌', avg(T) '✅正確率' format avg(T) '0.0%'"));
  header_(sh.getRange('A2:E2'));

  title_(sh, 'G1', '時段 × 角色 ✅正確率');
  sh.getRange('G2').setFormula(q_(DQ, "select D, avg(X) where D is not null group by D pivot I label D '時段(時)'"));
  sh.getRange('H3:K30').setNumberFormat('0.0%');
  header_(sh.getRange('G2:K2'));

  title_(sh, 'M1', '星期');
  sh.getRange('M2:P2').setValues([['星期', '檢核項次', '✅正確率', '❌未執行率']]);
  header_(sh.getRange('M2:P2'));
  ['一', '二', '三', '四', '五', '六', '日'].forEach(function (d, i) {
    var r = i + 3;
    sh.getRange(r, 13, 1, 4).setValues([[
      d,
      '=COUNTIF(' + D + 'E2:E,M' + r + ')',
      '=IFERROR(AVERAGEIF(' + D + 'E2:E,M' + r + ',' + D + 'X2:X),"")',
      '=IFERROR(AVERAGEIF(' + D + 'E2:E,M' + r + ',' + D + 'Z2:Z),"")'
    ]]);
  });
  sh.getRange('O3:P9').setNumberFormat('0.0%');
  sh.setFrozenRows(2);

  sh.insertChart(sh.newChart().asColumnChart()
    .addRange(sh.getRange('A2:A30')).addRange(sh.getRange('E2:E30'))
    .setOption('title', '各時段 ✅ 正確率').setOption('legend', { position: 'none' })
    .setOption('vAxis', { format: 'percent', viewWindow: { max: 1 } })
    .setPosition(12, 13, 0, 0).setOption('width', 560).setOption('height', 300).build());
}
