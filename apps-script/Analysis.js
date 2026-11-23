/**
 * 分析分頁：全部用公式即時計算，表單一送出就會更新。
 * 選單「重建分析分頁」會清掉並重畫這些分頁（不影響明細與每次紀錄）。
 *
 * 明細欄位：A填報編號 B送出時間 C洗冰日期 D洗冰時間 E時段 F星期 G週起日 H月份 I表單版本
 *           J角色 K人員 L階段 M項目編號 N檢核項目 O重點 P結果 Q有做到 R未做到 S填表人
 * 每次紀錄：A填報編號 B送出時間 C洗冰日期 D洗冰時間 E時段 F週起日 G月份 H表單版本
 *           I駕駛 J安全員1 K安全員2 L指揮官 M填表人 N檢核項數 O未做到項數 P達成率 Q重點未做到 R未做到項目 S備註
 */

var D = "'明細'!";      // 公式裡引用「明細」分頁的前綴
var R = "'每次紀錄'!";  // 公式裡引用「每次紀錄」分頁的前綴

function buildAnalysisSheets() {
  var ss = SpreadsheetApp.getActive();
  buildItems_(resetSheet_(ss, SHEET.ITEMS));
  buildPeople_(resetSheet_(ss, SHEET.PEOPLE));
  buildTrend_(resetSheet_(ss, SHEET.TREND));
  buildTime_(resetSheet_(ss, SHEET.TIME));
  buildOverview_(resetSheet_(ss, SHEET.OVERVIEW));

  // 分頁順序：總覽在最前面，資料分頁在後面
  [SHEET.OVERVIEW, SHEET.ITEMS, SHEET.PEOPLE, SHEET.TREND, SHEET.TIME, SHEET.RECORDS, SHEET.DETAIL, SHEET.ROSTER]
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
    sh.clear();
    sh.clearConditionalFormatRules();
    var f = sh.getFilter(); if (f) f.remove();
  } else {
    sh = ss.insertSheet(name);
  }
  return sh;
}

function title_(sh, a1, text) {
  sh.getRange(a1).setValue(text).setFontWeight('bold').setFontSize(12);
}

function header_(range) {
  range.setFontWeight('bold').setBackground('#e8eef7');
}

// ───────────────────────── 總覽 ─────────────────────────

function buildOverview_(sh) {
  sh.getRange('A1').setValue('洗冰檢點　總覽').setFontSize(16).setFontWeight('bold');
  var kpi = [
    ['累計洗冰次數', '=COUNTA(' + R + 'A2:A)', '0'],
    ['整體達成率', '=IFERROR(AVERAGE(' + D + 'Q2:Q),"")', '0.0%'],
    ['重點項目達成率', '=IFERROR(AVERAGEIF(' + D + 'O2:O,"是",' + D + 'Q2:Q),"")', '0.0%'],
    ['近 7 天達成率', '=IFERROR(AVERAGEIFS(' + D + 'Q2:Q,' + D + 'C2:C,">="&(TODAY()-6)),"")', '0.0%'],
    ['本月達成率', '=IFERROR(AVERAGEIF(' + D + 'H2:H,TEXT(TODAY(),"yyyy-mm"),' + D + 'Q2:Q),"")', '0.0%'],
    ['本月洗冰次數', '=COUNTIF(' + R + 'G2:G,TEXT(TODAY(),"yyyy-mm"))', '0']
  ];
  sh.getRange(3, 1, kpi.length, 2).setValues(kpi.map(function (k) { return [k[0], k[1]]; }));
  kpi.forEach(function (k, i) { sh.getRange(3 + i, 2).setNumberFormat(k[2]); });
  sh.getRange(3, 1, kpi.length, 1).setFontWeight('bold');
  sh.getRange(3, 2, kpi.length, 1).setFontSize(14).setHorizontalAlignment('right');

  title_(sh, 'A11', '最常未做到的 10 個項目');
  sh.getRange('A12:E12').setValues([['項目編號', '角色', '檢核項目', '未做到次數', '未做到率']]);
  header_(sh.getRange('A12:E12'));
  sh.getRange('A13').setFormula(
    '=IFERROR(SORTN(FILTER({\'項目分析\'!A2:A,\'項目分析\'!B2:B,\'項目分析\'!E2:E,\'項目分析\'!G2:G,\'項目分析\'!H2:H},' +
    '\'項目分析\'!G2:G>0),10,0,4,FALSE),"目前沒有未做到紀錄")');
  sh.getRange('E13:E22').setNumberFormat('0.0%');

  title_(sh, 'A25', '最近 10 次洗冰');
  sh.getRange('A26').setFormula('=QUERY(' + R + 'A1:S,"select C, D, L, N, O, P, R where A is not null order by B desc limit 10 ' +
    'label C \'日期\', D \'時間\', L \'指揮官\', N \'項數\', O \'未做到\', P \'達成率\', R \'未做到項目\' format C \'yyyy/mm/dd\', P \'0.0%\'",1)');
  header_(sh.getRange('A26:G26'));

  sh.setColumnWidth(1, 130).setColumnWidth(2, 110).setColumnWidth(3, 420).setColumnWidth(7, 420);
  sh.getRange('A27:G36').setWrap(true).setVerticalAlignment('top');
}

// ───────────────────────── 項目分析 ─────────────────────────

function buildItems_(sh) {
  var headers = ['項目編號', '角色', '階段', '重點', '檢核項目', '檢核次數', '未做到次數', '未做到率', '狀態'];
  var roleName = {}; ROLES.forEach(function (r) { roleName[r.key] = r.name; });
  var phaseName = {}; PHASES.forEach(function (p) { phaseName[p.key] = p.name; });
  var rows = ITEMS.map(function (it, i) {
    var r = i + 2;
    return [
      it.id, roleName[it.role], phaseName[it.phase], it.key ? '是' : '否', it.text,
      '=COUNTIF(' + D + 'M2:M,A' + r + ')',
      '=SUMIF(' + D + 'M2:M,A' + r + ',' + D + 'R2:R)',
      '=IF(F' + r + '=0,"",G' + r + '/F' + r + ')',
      it.retired ? '已停用' : '使用中'
    ];
  });
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  header_(sh.getRange(1, 1, 1, headers.length));
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.getRange(2, 8, rows.length, 1).setNumberFormat('0.0%');
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 110).setColumnWidth(3, 140).setColumnWidth(5, 460);
  sh.getRange(2, 5, rows.length, 1).setWrap(true);
  sh.getRange(1, 1, rows.length + 1, headers.length).createFilter();

  var rateRange = sh.getRange(2, 8, rows.length, 1);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue('#e06666', SpreadsheetApp.InterpolationType.NUMBER, '0.3')
    .setRanges([rateRange]).build();
  var keyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('是').setBackground('#f8cbad')
    .setRanges([sh.getRange(2, 4, rows.length, 1)]).build();
  sh.setConditionalFormatRules([rule, keyRule]);

  var chart = sh.newChart().asBarChart()
    .addRange(sh.getRange(1, 1, rows.length + 1, 1))
    .addRange(sh.getRange(1, 8, rows.length + 1, 1))
    .setOption('title', '各項目未做到率')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { format: 'percent' })
    .setOption('height', 900).setOption('width', 520)
    .setPosition(1, 11, 0, 0)
    .build();
  sh.insertChart(chart);
}

// ───────────────────────── 人員分析 ─────────────────────────

function buildPeople_(sh) {
  title_(sh, 'A1', '人員別達成率（不分角色）');
  sh.getRange('A2').setFormula('=QUERY(' + D + 'A1:S,"select K, count(Q), sum(R), avg(Q) where K is not null and K <> \'\' group by K order by avg(Q) asc ' +
    'label K \'人員\', count(Q) \'檢核項次\', sum(R) \'未做到次數\', avg(Q) \'達成率\' format avg(Q) \'0.0%\'",1)');
  header_(sh.getRange('A2:D2'));

  title_(sh, 'F1', '人員 × 角色');
  sh.getRange('F2').setFormula('=QUERY(' + D + 'A1:S,"select K, J, count(Q), sum(R), avg(Q) where K is not null and K <> \'\' group by K, J order by K ' +
    'label K \'人員\', J \'角色\', count(Q) \'檢核項次\', sum(R) \'未做到次數\', avg(Q) \'達成率\' format avg(Q) \'0.0%\'",1)');
  header_(sh.getRange('F2:J2'));

  title_(sh, 'L1', '人員重點項目達成率');
  sh.getRange('L2').setFormula('=QUERY(' + D + 'A1:S,"select K, count(Q), sum(R), avg(Q) where K is not null and K <> \'\' and O = \'是\' group by K order by avg(Q) asc ' +
    'label K \'人員\', count(Q) \'重點檢核項次\', sum(R) \'重點未做到\', avg(Q) \'重點達成率\' format avg(Q) \'0.0%\'",1)');
  header_(sh.getRange('L2:O2'));

  title_(sh, 'Q1', '各人員未做到的項目（次數多的在前）');
  sh.getRange('Q2').setFormula('=IFERROR(QUERY(' + D + 'A1:S,"select K, M, N, sum(R) where R = 1 group by K, M, N order by K, sum(R) desc ' +
    'label K \'人員\', M \'編號\', N \'檢核項目\', sum(R) \'未做到次數\'",1),"目前沒有未做到紀錄")');
  header_(sh.getRange('Q2:T2'));
  sh.setColumnWidth(19, 420);
  sh.setFrozenRows(2);
}

// ───────────────────────── 趨勢分析 ─────────────────────────

function buildTrend_(sh) {
  title_(sh, 'A1', '每日');
  sh.getRange('A2').setFormula('=QUERY(' + R + 'A1:S,"select C, count(A), sum(O), avg(P) where A is not null group by C order by C ' +
    'label C \'日期\', count(A) \'洗冰次數\', sum(O) \'未做到項數\', avg(P) \'達成率\' format C \'yyyy/mm/dd\', avg(P) \'0.0%\'",1)');
  header_(sh.getRange('A2:D2'));

  title_(sh, 'F1', '每週（週一起）');
  sh.getRange('F2').setFormula('=QUERY(' + R + 'A1:S,"select F, count(A), sum(O), avg(P) where A is not null group by F order by F ' +
    'label F \'週起日\', count(A) \'洗冰次數\', sum(O) \'未做到項數\', avg(P) \'達成率\' format F \'yyyy/mm/dd\', avg(P) \'0.0%\'",1)');
  header_(sh.getRange('F2:I2'));

  title_(sh, 'K1', '每月');
  sh.getRange('K2').setFormula('=QUERY(' + R + 'A1:S,"select G, count(A), sum(O), avg(P) where A is not null group by G order by G ' +
    'label G \'月份\', count(A) \'洗冰次數\', sum(O) \'未做到項數\', avg(P) \'達成率\' format avg(P) \'0.0%\'",1)');
  header_(sh.getRange('K2:N2'));

  title_(sh, 'P1', '每月 × 角色達成率');
  sh.getRange('P2').setFormula('=IFERROR(QUERY(' + D + 'A1:S,"select H, avg(Q) where H is not null and H <> \'\' group by H pivot J ' +
    'label H \'月份\'",1),"")');
  sh.getRange('Q3:T200').setNumberFormat('0.0%');
  header_(sh.getRange('P2:T2'));

  sh.setFrozenRows(2);
  sh.insertChart(sh.newChart().asLineChart()
    .addRange(sh.getRange('A2:A400')).addRange(sh.getRange('D2:D400'))
    .setOption('title', '每日達成率').setOption('legend', { position: 'none' })
    .setOption('vAxis', { format: 'percent', viewWindow: { max: 1 } })
    .setPosition(4, 16, 0, 160).setOption('width', 620).setOption('height', 300).build());
  sh.insertChart(sh.newChart().asColumnChart()
    .addRange(sh.getRange('F2:F200')).addRange(sh.getRange('I2:I200'))
    .setOption('title', '每週達成率').setOption('legend', { position: 'none' })
    .setOption('vAxis', { format: 'percent', viewWindow: { max: 1 } })
    .setPosition(20, 16, 0, 160).setOption('width', 620).setOption('height', 300).build());
}

// ───────────────────────── 時段分析 ─────────────────────────

function buildTime_(sh) {
  title_(sh, 'A1', '洗冰時段（整點）');
  sh.getRange('A2').setFormula('=QUERY(' + R + 'A1:S,"select E, count(A), sum(O), avg(P) where A is not null group by E order by E ' +
    'label E \'時段(時)\', count(A) \'洗冰次數\', sum(O) \'未做到項數\', avg(P) \'達成率\' format avg(P) \'0.0%\'",1)');
  header_(sh.getRange('A2:D2'));

  title_(sh, 'F1', '時段 × 角色達成率');
  sh.getRange('F2').setFormula('=IFERROR(QUERY(' + D + 'A1:S,"select E, avg(Q) where E is not null group by E pivot J label E \'時段(時)\'",1),"")');
  sh.getRange('G3:J30').setNumberFormat('0.0%');
  header_(sh.getRange('F2:J2'));

  title_(sh, 'L1', '星期');
  var days = ['一', '二', '三', '四', '五', '六', '日'];
  sh.getRange('L2:N2').setValues([['星期', '檢核項次', '達成率']]);
  header_(sh.getRange('L2:N2'));
  days.forEach(function (d, i) {
    var r = i + 3;
    sh.getRange(r, 12, 1, 3).setValues([[
      d,
      '=COUNTIF(' + D + 'F2:F,L' + r + ')',
      '=IFERROR(AVERAGEIF(' + D + 'F2:F,L' + r + ',' + D + 'Q2:Q),"")'
    ]]);
  });
  sh.getRange('N3:N9').setNumberFormat('0.0%');

  sh.setFrozenRows(2);
  sh.insertChart(sh.newChart().asColumnChart()
    .addRange(sh.getRange('A2:A30')).addRange(sh.getRange('D2:D30'))
    .setOption('title', '各時段達成率').setOption('legend', { position: 'none' })
    .setOption('vAxis', { format: 'percent', viewWindow: { max: 1 } })
    .setPosition(12, 12, 0, 0).setOption('width', 560).setOption('height', 300).build());
}
