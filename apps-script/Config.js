/**
 * 題目設定（由原始表單 11509-2版.xlsx 自動產生）。
 * 改版規則：
 *   - 題目文字不變（或只修錯字）→ 保留原編號，跨版本統計才接得起來。
 *   - 新增題目 → 給新編號（同角色往後編）。
 *   - 刪除題目 → 不要刪掉這一行，改加 retired: true（新表單不會出現，舊回覆仍可重算）。編號不要重複使用。
 */
var FORM_VERSION = '11509-2';

var ROLES = [
  { key: 'driver', name: '洗冰車駕駛' },
  { key: 'safety1', name: '冰上安全員1' },
  { key: 'safety2', name: '冰上安全員2' },
  { key: 'commander', name: '值班洗冰指揮官' },
];

var PHASES = [
  { key: 'P1', name: '廣播至倒車定位期間' },
  { key: 'P2', name: '洗外圍第一圈' },
  { key: 'P3', name: '洗中圈至回冰室' },
];

// key: true = 重點檢核項目（原表橘底）
var ITEMS = [
  { id: 'D01', role: 'driver', phase: 'P1', key: true, text: "經過教練室前走道是否有將無玻璃2區圍板牆上水壺,小音箱,衣物等物品移至圍板外側中間框台" },
  { id: 'D02', role: 'driver', phase: 'P1', key: false, text: "倒車前是否有先按壓喇叭警示通道及冰面" },
  { id: 'D03', role: 'driver', phase: 'P1', key: false, text: "倒車到走道時是否先停車等候" },
  { id: 'D04', role: 'driver', phase: 'P1', key: false, text: "是否有等候安全員高舉單手示意倒車路徑淨空才開始倒車入冰面" },
  { id: 'D05', role: 'driver', phase: 'P1', key: true, text: "入冰面門時是否有將頭先側彎,避免頭碰觸到天花板" },
  { id: 'D06', role: 'driver', phase: 'P1', key: false, text: "入冰面後是否有於待命區先停車等候" },
  { id: 'D07', role: 'driver', phase: 'P1', key: false, text: "是否有等候安全員1&2高舉單手示意全淨空才開始行進洗冰" },
  { id: 'D08', role: 'driver', phase: 'P2', key: true, text: "洗外圍圈時頭是否有側彎觀看左側與左側前方冰面有無異物" },
  { id: 'D09', role: 'driver', phase: 'P2', key: false, text: "轉彎處是否有減速" },
  { id: 'D10', role: 'driver', phase: 'P2', key: false, text: "車頭前緣行經立柱轉彎螢光貼時是否有開始轉彎" },
  { id: 'D11', role: 'driver', phase: 'P3', key: false, text: "洗冰車要開進洗冰車室方向時是否有按壓喇叭警示通道及冰面" },
  { id: 'D12', role: 'driver', phase: 'P3', key: true, text: "到出口前是否有先停車+收起刮刀+同仁頭側彎+再前進(防撞頭)" },
  { id: 'D13', role: 'driver', phase: 'P3', key: false, text: "進洗冰室是否行進到前輪碰木樁才停車+舉升冰槽(防撞天花板)" },
  { id: 'D14', role: 'driver', phase: 'P3', key: false, text: "是否有在現場待洗冰室鐵捲門完全關閉無再彈起後再離開" },
  { id: 'S101', role: 'safety1', phase: 'P1', key: true, text: "移動角錐,球門,擋板時是否有左右張望做角錐周邊檢視異物動作" },
  { id: 'S102', role: 'safety1', phase: 'P1', key: false, text: "值班吹哨時是否即到洗冰室前進行人物淨空作業" },
  { id: 'S103', role: 'safety1', phase: 'P1', key: true, text: "進行洗冰車倒車路徑淨空時是否有向冰面做出指差確認動作" },
  { id: 'S104', role: 'safety1', phase: 'P1', key: true, text: "倒車路徑淨空後是否有向駕駛高舉單手示意倒車路徑淨空" },
  { id: 'S105', role: 'safety1', phase: 'P1', key: false, text: "是否有確認教練室區西側出入口門皆關閉" },
  { id: 'S106', role: 'safety1', phase: 'P1', key: true, text: "是否有在洗冰室側南區環繞一圈並向四個方向做出淨空指差" },
  { id: 'S107', role: 'safety1', phase: 'P1', key: true, text: "是否有向駕駛高舉單手示意洗冰室側南區已淨空完成" },
  { id: 'S108', role: 'safety1', phase: 'P2', key: true, text: "是否到洗冰室側南區球門區警戒待命確認無人或物品再侵入" },
  { id: 'S109', role: 'safety1', phase: 'P2', key: false, text: "洗冰車洗過教練室2後是否移位到洗冰室前協助將球門推靠圍板" },
  { id: 'S110', role: 'safety1', phase: 'P2', key: false, text: "洗冰車洗過觀眾席3區後是否移位到鞋部前將角錐推靠圍板" },
  { id: 'S111', role: 'safety1', phase: 'P3', key: true, text: "是否有至洗冰室前圍板門口冰面警戒" },
  { id: 'S112', role: 'safety1', phase: 'P3', key: false, text: "若有緊急狀況是否有先作險阻手勢示意停車後,隨即前往排除" },
  { id: 'S113', role: 'safety1', phase: 'P3', key: false, text: "洗冰完成是否協助恢復角錐,球門,擋牆定位" },
  { id: 'S201', role: 'safety2', phase: 'P1', key: true, text: "移動角錐,球門,擋板時是否有左右張望做角錐周邊檢視異物動作" },
  { id: 'S202', role: 'safety2', phase: 'P1', key: false, text: "是否有確認觀眾席區東側所有出入口門皆關閉" },
  { id: 'S203', role: 'safety2', phase: 'P1', key: true, text: "是否有在鞋部側北區環繞一圈並向四個方向做出淨空指差" },
  { id: 'S204', role: 'safety2', phase: 'P1', key: true, text: "是否有向駕駛高舉單手示意鞋部側北區已淨空完成" },
  { id: 'S205', role: 'safety2', phase: 'P2', key: true, text: "是否有到鞋部側北區角錐處警戒待命確認無人或物品再進入" },
  { id: 'S206', role: 'safety2', phase: 'P2', key: false, text: "洗冰車洗過教練室2後是否移位到洗冰室前將球門推靠圍板" },
  { id: 'S207', role: 'safety2', phase: 'P2', key: false, text: "協助將球門推靠圍板後是否有回鞋部側北區角錐放置處警戒" },
  { id: 'S208', role: 'safety2', phase: 'P3', key: false, text: "是否回鞋部協助收護具冰鞋並警戒北側區冰面" },
  { id: 'S209', role: 'safety2', phase: 'P3', key: false, text: "洗冰完成是否協助恢復角錐,球門,擋牆定位" },
  { id: 'C01', role: 'commander', phase: 'P1', key: false, text: "確認54分洗冰預告廣播是否持續播放 56分是否有協助上冰引導人員離開冰面" },
  { id: 'C02', role: 'commander', phase: 'P1', key: false, text: "00分是否有吹哨提醒冰面人員即將洗冰,未離開人員迅速引導離開" },
  { id: 'C03', role: 'commander', phase: 'P1', key: true, text: "確認洗冰車倒車到走道時是否先停車等候" },
  { id: 'C04', role: 'commander', phase: 'P1', key: false, text: "安全員1在完成洗冰室前倒車路徑淨空後,是否有向冰面做出指差確認動作" },
  { id: 'C05', role: 'commander', phase: 'P1', key: false, text: "冰上安全員1在倒車路徑淨空後是否有向駕駛高舉單手示意倒車路徑淨空" },
  { id: 'C06', role: 'commander', phase: 'P1', key: true, text: "確認洗冰車是否在看到安全員1&2高舉單手示意倒車路徑已淨空完成才開始前進" },
  { id: 'C07', role: 'commander', phase: 'P1', key: false, text: "安全員1&2完成南北側淨空後,是否有環繞一圈並向四個方向冰面做出指差動作" },
  { id: 'C08', role: 'commander', phase: 'P2', key: false, text: "安全員1&2完成南北側淨空後,是否有向駕駛高舉單手示意南北區已淨空完成" },
  { id: 'C09', role: 'commander', phase: 'P2', key: true, text: "確認洗冰車是否在看到安全員1&2分別高舉單手示意已淨空完成才開始前進" },
  { id: 'C10', role: 'commander', phase: 'P3', key: false, text: "是否回鞋部後有持續警戒北側區" },
  { id: 'C11', role: 'commander', phase: 'P3', key: false, text: "洗冰完成鞋部是否有播放小心冰面濕滑廣播" },
  { id: 'C12', role: 'commander', phase: 'P3', key: false, text: "櫃台冰面人潮計數是否有歸零" },
];

var SAFETY_NOTE = '※注意→洗冰期間有事向駕駛轉達一律「先停車、再溝通」／預警碰撞警音鳴叫或LED亮起需先停車／有緊急停車要作險阻手勢（雙手高舉交叉揮動）＋拍圍板出聲';

if (typeof module !== 'undefined') module.exports = { FORM_VERSION: FORM_VERSION, ROLES: ROLES, PHASES: PHASES, ITEMS: ITEMS, SAFETY_NOTE: SAFETY_NOTE };

