// 筒賀水質管理センター週報アプリ GASバックエンド（v2 - Phase 2: subtab + transcription対応）
// 注: シート参照は index 番号で統一（H yamane ルール: エンコーディング問題回避）
//   テンプレ Excel: SHEETS[0]=日常水質, [1]=運転管理月報, [2]=水質管理報告,
//                   [3]=日常機械設備, [4]=日常電気設備, [5]=主要機器運転時間
//   SS-D（蓄積用）: シート名で参照（書き込み専用なので名前OK）

// テンプレート Spreadsheet ID（exportExcel で月次コピー元として使用）
var SPREADSHEET_ID = '1XFPH90_XXvyGQfewQXZORcdvc-jH2jAzP_yxV2pcrpQ';

// テンプレ内のシート index（シート名ではなく番号で参照する）
var SHEET_IDX = {
  daily_water: 0,         // 日常水質
  operation_monthly: 1,   // 運転管理月報
  water_report: 2,        // 水質管理報告
  daily_mech: 3,          // 日常機械設備
  daily_elec: 4,          // 日常電気設備
  equipment_hours: 5      // 主要機器運転時間
};

var DAY_NAMES = { mon: '月', tue: '火', wed: '水', thu: '木', fri: '金' };
var DAY_REVERSE = { '月': 'mon', '火': 'tue', '水': 'wed', '木': 'thu', '金': 'fri' };
var DAYS = ['mon','tue','wed','thu','fri'];
var DAY_OFFSETS = { mon: 0, tue: 1, wed: 2, thu: 3, fri: 4 };

// ===== 日常水質シート（テンプレ idx=0） =====
// 各週ブロック開始行: r1, r53, r105, r157, r209（ピッチ52行）
var WATER_BLOCK_ROW = { 1: 1, 2: 53, 3: 105, 4: 157, 5: 209 };
var WATER_WEEK_STARTS = [1, 53, 105, 157, 209]; // 0-based index 互換 (W1=0, W5=4)
var WATER_WEEK_OFFSET = 52;          // 全週統一ピッチ
var WATER_INSPECTOR_ROW_OFFSET = 14; // r15 (=block+14)
var WATER_BIKOU_ROW_OFFSET = 43;     // r44

// 上段使用量 11項目 (各週単位、6セル: mon-fri+nextMon)
// アプリ→SS-D保存時のキー
var USAGE_KEYS = [
  'powerAllDay','powerMeasure','powerReactive','demand','water','diesel',
  'returnSludge1','returnSludge2','discharge','excessSludge1','excessSludge2'
];
// 上段Excel書き込み: ブロック内のオフセット (r4 = offset 3)
// 順: 全日電力量(r4), 力測(r5), 無効(r6), デマンド(r7), 水道(r8), 軽油(r9),
//     返送1(r10), 返送2(r11), 放流(r12)
//     ※ 余剰汚泥1/2 は日常水質シートには書き込まない（日常電気設備の余剰汚泥流量計セルに転記）
var USAGE_ROW_OFFSET = {
  powerAllDay: 3, powerMeasure: 4, powerReactive: 5, demand: 6,
  water: 7, diesel: 8,
  returnSludge1: 9, returnSludge2: 10, discharge: 11
  // excessSludge1, excessSludge2 → 日常水質には書き込まない
};
// 上段曜日カラム: I=9, J=10, K=11, L=12, M=13, N=14（mon..nextMon）
var USAGE_DAY_COL = { mon: 9, tue: 10, wed: 11, thu: 12, fri: 13, nextMon: 14 };
var USAGE_CELLS = ['mon','tue','wed','thu','fri','nextMon'];

// 下段水質測定 34項目/日 (per-day field) — Excel 行オフセット
var MEASURE_ROW_OFFSET = {
  // 水温（4 + 終沈2）
  tempIn: 15, tempD1: 16, tempD2: 17, tempOut: 18,
  // 外観（4）
  appearIn: 19, appearD1: 20, appearD2: 21, appearOut: 22,
  // 透視度（2）+ 終沈2は水質管理報告のみ
  transIn: 23, transOut: 24,
  // 臭気（4）
  odorIn: 25, odorD1: 26, odorD2: 27, odorOut: 28,
  // PH（4）+ 終沈2は水質管理報告のみ
  phIn: 29, phD1: 30, phD2: 31, phOut: 32,
  // SV 1系・2系（SV30/SV24h）
  sv1_30: 33, sv1_24h: 34, sv2_30: 35, sv2_24h: 36,
  // MLSS（1系/2系）
  mlss1: 37, mlss2: 38,
  // 汚泥界面（1系/2系終沈）
  sludge1: 39, sludge2: 40,
  // 残留塩素 / 塩素投入量
  chlorine: 41, chlorineDose: 42
  // 終沈水温/透視度/PH（No.1/No.2）は水質管理報告に直接転記
};
// 下段曜日カラム: J=10, K=11, L=12, M=13, N=14（mon..fri）
var MEASURE_DAY_COL = { mon: 10, tue: 11, wed: 12, thu: 13, fri: 14 };

// 下段で取り扱う全フィールド（フォームと同期）
var MEASURE_FIELDS = [
  'tempIn','tempD1','tempD2','tempOut','tempFinal1','tempFinal2',
  'appearIn','appearD1','appearD2','appearOut',
  'transIn','transOut','transFinal1','transFinal2',
  'odorIn','odorD1','odorD2','odorOut',
  'phIn','phD1','phD2','phOut','phFinal1','phFinal2',
  'sv1_30','sv1_24h','sv2_30','sv2_24h',
  'mlss1','mlss2',
  'sludge1','sludge2',
  'chlorine','chlorineDose',
  'inspector','bikou'
];

// ===== 水質管理報告シート（テンプレ idx=2） =====
// 前半 r8=日付A=A6 (月初), r8〜r38 (月初+0〜+30日)
// 後半 r49=A8, r79=A38 (同じ日付列を参照)
// 行算出: row前半 = 7 + day_of_month, row後半 = 48 + day_of_month
function reportRowFront(date) { return 7 + date.getDate(); }
function reportRowBack(date)  { return 48 + date.getDate(); }

// 列マッピング（前半: r8-r38）
var REPORT_FRONT_COL = {
  // 水温
  tempIn: 3,    // C
  tempD1: 4,    // D（No.1ディッチ）
  tempD2: 5,    // E（No.2ディッチ）
  tempFinal1: 6,// F（No.1終沈）
  tempFinal2: 7,// G（No.2終沈）
  tempOut: 8,   // H（放流水）
  // 外観 (I-N) — アプリ入力なし、テンプレ固定値（黄灰色/茶色/透明等）
  // 透視度
  transIn: 15,    // O（流入）
  transFinal1: 16,// P（No.1終沈）
  transFinal2: 17,// Q（No.2終沈）
  transOut: 18    // R（放流）
};
// 列マッピング（後半: r49-r79）
var REPORT_BACK_COL = {
  // 臭気 (C-H) — テンプレ固定
  // DO値
  doNo1: 9,     // I
  doNo2: 10,    // J
  // PH
  phIn: 11,    // K
  phD1: 12,    // L
  phD2: 13,    // M
  phFinal1: 14,// N（No.1終沈）
  phFinal2: 15,// O（No.2終沈）
  phOut: 16,   // P
  // 残留塩素
  chlorine: 17, // Q
  // SV30
  sv1_30: 18,   // R（1系）
  sv2_30: 19    // S（2系）
};

// ─── 仕様マッピング（Japanese-keyed, 参照用） ───
// 筒賀_水質シート行マッピング.js / .md の表現を GAS にそのまま転記。
// 既存の {USAGE,MEASURE}_ROW_OFFSET（English 名）と等価で、用途は (1) 仕様との
// 1:1 検証 (2) 将来 getWaterRow() 経由で Japanese キーから絶対行を引くため。
var WATER_USAGE_ROWS_W1 = {
  "全日電力量": 4,        // Kwh   (English: powerAllDay)
  "力測電力量": 5,        // Kwh   (powerMeasure)
  "無効電力量": 6,        // Kvar  (powerReactive)
  "デマンド": 7,          // Kw    (demand)
  "水道使用量": 8,        // ㎥    (water)
  "軽油（残量）": 9,      // ℓ     (diesel)
  "返送汚泥流量（№1）": 10, // ㎥    (returnSludge1)
  "返送汚泥流量（№2）": 11, // ㎥    (returnSludge2)
  "放流流量": 12          // ㎥    (discharge)
  // 余剰汚泥流量 No.1/No.2 はテンプレに行が無いので 日常電気設備 sub2 r56/r57 へ書き込み (ELEC_EXCESS_KEY_NO{1,2})
};

var WATER_MEASURE_ROWS_W1 = {
  "水温_流入": 16,        "水温_ディッチ1": 17,    "水温_ディッチ2": 18,    "水温_放流": 19,
  "外観_流入": 20,        "外観_ディッチ1": 21,    "外観_ディッチ2": 22,    "外観_放流": 23,
  "透視度_流入水": 24,    "透視度_放流水": 25,
  "臭気_流入": 26,        "臭気_ディッチ1": 27,    "臭気_ディッチ2": 28,    "臭気_放流": 29,
  "PH_流入": 30,          "PH_ディッチ1": 31,      "PH_ディッチ2": 32,      "PH_放流": 33,
  "SV1系_SV30": 34,       "SV1系_24h": 35,         "SV2系_SV30": 36,        "SV2系_24h": 37,
  "MLSS_1系": 38,         "MLSS_2系": 39,
  "汚泥界面_1系終沈": 40, "汚泥界面_2系終沈": 41,
  "残留塩素_放流水": 42,
  "塩素投入量_塩素混和池": 43
  // 終沈3項目（水温/透視度/PH × No.1/No.2）はテンプレに行が無いので
  // 水質管理報告に直接書き込み: REPORT_FRONT_COL.tempFinal{1,2}/transFinal{1,2}, REPORT_BACK_COL.phFinal{1,2}
};

// W{n} の絶対行 = baseRow + (n-1) * 52
//   week: 1〜5
//   itemKey: WATER_USAGE_ROWS_W1 / WATER_MEASURE_ROWS_W1 のいずれかのキー
function getWaterRow(week, itemKey) {
  var w = parseInt(week, 10);
  if (isNaN(w) || w < 1 || w > 5) throw new Error('Invalid week: ' + week);
  var baseRow = (WATER_USAGE_ROWS_W1[itemKey] != null) ? WATER_USAGE_ROWS_W1[itemKey]
              : (WATER_MEASURE_ROWS_W1[itemKey] != null) ? WATER_MEASURE_ROWS_W1[itemKey]
              : null;
  if (baseRow == null) return null;
  return baseRow + WATER_WEEK_OFFSET * (w - 1);
}

// 水質管理報告の前半・後半行を計算
//   dayOfMonth: 1〜31
//   返り値: { frontRow, backRow }
//   ※ 平日のみ書き込み (土日・対象月外は呼び出し側でフィルタ)
function getWQRRows(dayOfMonth) {
  var d = parseInt(dayOfMonth, 10);
  if (isNaN(d) || d < 1 || d > 31) throw new Error('Invalid dayOfMonth: ' + dayOfMonth);
  return { frontRow: 7 + d, backRow: 48 + d };
}

// ===== 主要機器運転時間 =====
var EQUIP_NAMES = [
  'mainPump1Hr','mainPump2Hr',
  'aeration1Hr','aeration2Hr','aeration3Hr','aeration4Hr',
  'finalScraper1Hr','finalScraper2Hr',
  'returnPump1Hr','returnPump2Hr','returnPump3Hr','returnPump4Hr',
  'excessPump1Hr','excessPump2Hr',
  'finalDrainPumpHr',
  'thickPump1Hr','thickPump2Hr',
  'returnWaterPump1Hr','returnWaterPump2Hr',
  'sludgeRoomDrainPumpHr',
  'deodorFanHr','generatorHr'
];
var EQUIP_LABELS = [
  'NO.1主ポンプ','NO.2主ポンプ',
  'NO.1エアレーション装置','NO.2エアレーション装置','NO.3エアレーション装置','NO.4エアレーション装置',
  '№1終沈汚泥かき寄せ機','№2終沈汚泥かき寄せ機',
  'NO.1返送汚泥ポンプ','NO.2返送汚泥ポンプ','NO.3返送汚泥ポンプ','NO.4返送汚泥ポンプ',
  'NO.1余剰汚泥ポンプ','NO.2余剰汚泥ポンプ',
  '終沈室床排ポンプ',
  '№1濃縮汚泥引抜ポンプ','№2濃縮汚泥引抜ポンプ',
  'NO.1返流水ポンプ','NO.2返流水ポンプ',
  '汚泥ポンプ室床排ポンプ',
  '臭気ガス吸引装置','発電機'
];
var EQUIP_WEEK_START = { 1: 8, 2: 46, 3: 84, 4: 122, 5: 160 };
var EQUIP_WEEK_DATE_CELL = { 1: 'A5', 2: 'A43', 3: 'A81', 4: 'A119', 5: 'A157' };
var EQUIP_HOUR_DAY_COL = { mon: 4, tue: 5, wed: 6, thu: 7, fri: 8 };

// ===== 日常電気設備（テンプレ idx=4） =====
// 入力列 G-K (7-11) for mon-fri
var ELEC_DAY_COL = { mon: 7, tue: 8, wed: 9, thu: 10, fri: 11 };

// W1基準の絶対行 + サブタブごとに異なる週オフセット
//   sub1: W1→W2 が +103 (W1の備考領域が大きい)
//   sub2: W1→W2 が +101
//   sub3: W1→W2 が +100
//   W2 以降は全サブタブで +100 統一 (機械設備のような W4→W5 ピッチ違いはなし)
var WEEK_OFFSETS_ELEC = {
  sub1: [0, 103, 203, 303, 403],
  sub2: [0, 101, 201, 301, 401],
  sub3: [0, 100, 200, 300, 400]
};
// 旧名互換 (sub1 視点 — 単一の電気週オフセットを期待する古いコード向け)
var ELEC_WEEK_BASE = { 1: 0, 2: 103, 3: 203, 4: 303, 5: 403 };

// ─── 点検表 1/3 (電気室・配電系) ─── W1基準の絶対行（29項目）
var ELEC1_ROW_W1 = {
  // 引込・受電
  "引込・受電__電圧__KV": 3,
  "引込・受電__電圧__モード": 4,
  "引込・受電__力率__ψ": 5,
  "引込・受電__電流__A": 6,
  "引込・受電__電流__モード": 7,
  "引込・受電__普通電力量(×10kw)__読み": 8,
  "引込・受電__普通電力量(×10kw)__使用量": 9,
  "引込・受電__引込・受電設定__手動/自動": 10,
  "引込・受電__引込・受電設定__切/入": 11,
  // 変圧器
  "変圧器__主変圧器温度__℃": 12,
  "変圧器__異音、脱臭の有無__有・無": 13,
  // 低圧分岐
  "低圧分岐__動力(210V)電圧__V": 14,
  "低圧分岐__動力(210V)電圧__モード": 15,
  "低圧分岐__動力(210V)電流__A": 16,
  "低圧分岐__動力(210V)電流__モード": 17,
  "低圧分岐__照明(210-105V)電流__A": 18,
  "低圧分岐__照明(210-105V)電流__モード": 19,
  "低圧分岐__低圧分岐設備__買電/自家発": 20,
  "低圧分岐__沈砂池ポンプ設備__読み": 21,
  "低圧分岐__沈砂池ポンプ設備__使用量": 22,
  "低圧分岐__水処理設備(1)__読み": 23,
  "低圧分岐__水処理設備(1)__使用量": 24,
  "低圧分岐__水処理設備(2)__読み": 25,
  "低圧分岐__水処理設備(2)__使用量": 26,
  "低圧分岐__汚泥処理設備電力量__読み": 27,
  "低圧分岐__汚泥処理設備電力量__使用量": 28,
  "低圧分岐__No.1コンデンサー__切/入": 29,
  "低圧分岐__No.2コンデンサー__切/入": 30,
  "低圧分岐__No.3コンデンサー__切/入": 31
};

// ─── 点検表 2/3 (計装設備：流量計系・DO計含む) ─── W1基準（26項目）
var ELEC2_ROW_W1 = {
  // No.1 オキシデーションディッチ DO計
  "Nｏ．１オキシデーションディッチＤＯ計__外観の確認": 40,
  "Nｏ．１オキシデーションディッチＤＯ計__指示状況の確認": 41,
  "Nｏ．１オキシデーションディッチＤＯ計__ＤＯ値__㎎/ℓ": 42,
  // No.2 オキシデーションディッチ DO計
  "Nｏ．２オキシデーションディッチＤＯ計__外観の確認": 43,
  "Nｏ．２オキシデーションディッチＤＯ計__指示状況の確認": 44,
  "Nｏ．２オキシデーションディッチＤＯ計__ＤＯ値__㎎/ℓ": 45,
  // No.1 返送汚泥流量計
  "Nｏ．１返送汚泥流量計__外観の確認": 46,
  "Nｏ．１返送汚泥流量計__指示状況の確認": 47,
  "Nｏ．１返送汚泥流量計__流量計の読み__㎥/ｈ": 48,
  "Nｏ．１返送汚泥流量計__流量計積算計読み__㎥": 49,
  // No.2 返送汚泥流量計
  "Nｏ．２返送汚泥流量計__外観の確認": 50,
  "Nｏ．２返送汚泥流量計__指示状況の確認": 51,
  "Nｏ．２返送汚泥流量計__流量計の読み__㎥/ｈ": 52,
  "Nｏ．２返送汚泥流量計__流量計積算計読み__㎥": 53,
  // 余剰汚泥流量計
  "余剰汚泥流量計__外観の確認": 54,
  "余剰汚泥流量計__指示状況の確認": 55,
  "余剰汚泥流量計__No.１余剰汚泥ポンプ流量計読み__㎥/ｈ": 56,
  "余剰汚泥流量計__No.２余剰汚泥ポンプ流量計読み__㎥/ｈ": 57,
  // 放流流量計
  "放流流量計__外観の確認": 58,
  "放流流量計__指示状況の確認": 59,
  "放流流量計__流量計の読み__㎥/ｈ": 60,
  "放流流量計__流量計積算の読み__㎥": 61,
  // 汚濁負荷量計
  "汚濁負荷量計__外観の確認": 62,
  "汚濁負荷量計__指示状況の確認": 63,
  // (放流水UV)
  "（放流水ＵＶ）__ＵＶ値__Abs": 64,
  "（放流水ＵＶ）__ＣＯＤ値__㎎/ℓ": 65
};

// ─── 点検表 3/3 (計装設備：水質計系・発電機) ─── W1基準（12項目）
var ELEC3_ROW_W1 = {
  // 全窒素・全燐計
  "全窒素・全燐計__外観の確認": 72,
  "全窒素・全燐計__指示状況の確認": 73,
  "全窒素・全燐計__計器読み（全窒素）__㎎/ℓ": 74,
  "全窒素・全燐計__計器読み（全りん）__㎎/ℓ": 75,
  // 圧力式液位計
  "圧力式液位計__外観の確認": 76,
  "圧力式液位計__指示状況の確認": 77,
  "圧力式液位計__機器指示地の記録__ｍ": 78,
  // 発電機・エンジン・燃料タンク
  "発電機盤__外観の確認": 79,
  "エンジン__外観の確認": 80,
  "燃料タンク__外観の確認": 81,
  "燃料タンク__残量の値__ℓ": 82,
  // 通報装置
  "通報装置__通報状態の確認": 83
};

// 旧名互換エイリアス (コード上の他参照向け)
var ELEC1_ROW = ELEC1_ROW_W1;
var ELEC2_ROW = ELEC2_ROW_W1;
var ELEC3_ROW = ELEC3_ROW_W1;

// ─── 特殊書き込み先キー（GAS 複合転記用） ───
// DO値: 日常水質下段の do1/do2 → 日常電気設備 sub2 + 水質管理報告
var ELEC_DO_KEY_NO1 = "Nｏ．１オキシデーションディッチＤＯ計__ＤＯ値__㎎/ℓ"; // sub2 r42
var ELEC_DO_KEY_NO2 = "Nｏ．２オキシデーションディッチＤＯ計__ＤＯ値__㎎/ℓ"; // sub2 r45
// 余剰汚泥流量: 上段 excessSludge1/2 → 日常電気設備 sub2
var ELEC_EXCESS_KEY_NO1 = "余剰汚泥流量計__No.１余剰汚泥ポンプ流量計読み__㎥/ｈ"; // sub2 r56
var ELEC_EXCESS_KEY_NO2 = "余剰汚泥流量計__No.２余剰汚泥ポンプ流量計読み__㎥/ｈ"; // sub2 r57
// 燃料タンク残量: 上段 diesel → 日常電気設備 sub3
var ELEC_FUEL_TANK_KEY = "燃料タンク__残量の値__ℓ"; // sub3 r82

// W1基準行 + サブタブ別週オフセット で絶対行を返す
//   week: 1〜5
//   subId: 'sub1' | 'sub2' | 'sub3'  (互換: 1|2|3)
//   itemKey: マッピングのキー (例: "引込・受電__電圧__KV")
function getElecRow(week, subId, itemKey) {
  var w = parseInt(week, 10);
  if (isNaN(w) || w < 1 || w > 5) throw new Error('Invalid week: ' + week);
  var s = String(subId);
  var key = (s === 'sub1' || s === '1') ? 'sub1'
          : (s === 'sub2' || s === '2') ? 'sub2'
          : (s === 'sub3' || s === '3') ? 'sub3'
          : null;
  if (!key) throw new Error('Unknown elec subId: ' + subId);
  var map = (key === 'sub1') ? ELEC1_ROW_W1 : (key === 'sub2') ? ELEC2_ROW_W1 : ELEC3_ROW_W1;
  var baseRow = map[itemKey];
  if (baseRow == null) return null; // フォーム側との同期問題は no-op
  return baseRow + WEEK_OFFSETS_ELEC[key][w - 1];
}

// 電気フィールド全集合 (フォーム + GAS転記用)
var ELEC_FIELDS = []
  .concat(Object.keys(ELEC1_ROW_W1))
  .concat(Object.keys(ELEC2_ROW_W1))
  .concat(Object.keys(ELEC3_ROW_W1));

// ===== 日常機械設備（テンプレ idx=3） =====
// 入力列 E-I (5-9) for mon-fri
var MECH_DAY_COL = { mon: 5, tue: 6, wed: 7, thu: 8, fri: 9 };
// 5週分の絶対行オフセット（W1基準, 0-based index）
//  - W1→W2→W3→W4 は 177行ピッチ
//  - W4→W5 は 176行ピッチ（注意: 他と異なる）
var WEEK_OFFSETS_MECH = [0, 177, 354, 531, 707];
// 旧名互換（1-based week => offset）
var MECH_WEEK_BASE = { 1: 0, 2: 177, 3: 354, 4: 531, 5: 707 };

// ─── 点検表 1/3 (処理系前段) ─── W1基準の絶対行（56項目）
var MECH1_ROW_W1 = {
  // ポンプ井
  "ポンプ井__common1__内部状況確認": 4,
  // 汚水ポンプ
  "汚水ポンプ__1__電流値": 5,
  "汚水ポンプ__1__吐出圧力": 6,
  "汚水ポンプ__1__異音、振動": 7,
  "汚水ポンプ__2__電流値": 8,
  "汚水ポンプ__2__吐出圧力": 9,
  "汚水ポンプ__2__異音、振動": 10,
  "汚水ポンプ__common1__モード確認": 11,
  "汚水ポンプ__common1__機種選択": 12,
  // 自動除塵機
  "自動除塵機__油圧圧力": 13,
  "自動除塵機__異音、振動": 14,
  "自動除塵機__し渣量の確認": 15,
  "自動除塵機__運転動作": 16,
  // エアレーション装置
  "エアレーション装置__No1__電流値": 17,
  "エアレーション装置__No1__異音、振動": 18,
  "エアレーション装置__No1__モード確認": 19,
  "エアレーション装置__No2__電流値": 20,
  "エアレーション装置__No2__異音、振動": 21,
  "エアレーション装置__No2__モード確認": 22,
  "エアレーション装置__No3__電流値": 23,
  "エアレーション装置__No3__異音、振動": 24,
  "エアレーション装置__No3__モード確認": 25,
  "エアレーション装置__No4__電流値": 26,
  "エアレーション装置__No4__異音、振動": 27,
  "エアレーション装置__No4__モード確認": 28,
  // 消泡装置
  "消泡装置__common1__消泡ノズル": 29,
  "消泡装置__common1__消泡状態": 30,
  // 流出可動堰
  "流出可動堰__1__開度": 31,
  "流出可動堰__2__開度": 32,
  // 流出ゲート
  "流出ゲート__1__開度": 33,
  "流出ゲート__2__開度": 34,
  // 連絡ゲート
  "連絡ゲート__common1__開度": 35,
  // 終沈汚泥掻寄機
  "終沈汚泥掻寄機__1__電流値": 36,
  "終沈汚泥掻寄機__1__異音、振動等": 37,
  "終沈汚泥掻寄機__2__電流値": 38,
  "終沈汚泥掻寄機__2__異音、振動等": 39,
  // スカムスキマー
  "スカムスキマー__1__動作状況": 40,
  "スカムスキマー__1__スカムの状況": 41,
  "スカムスキマー__2__動作状況": 42,
  "スカムスキマー__2__スカムの状況": 43,
  // 返送汚泥ポンプ (No1-2 / common1, No3-4 / common2)
  "返送汚泥ポンプ__No1__電流値": 44,
  "返送汚泥ポンプ__No1__吐出圧力": 45,
  "返送汚泥ポンプ__No1__異音、振動": 46,
  "返送汚泥ポンプ__No2__電流値": 47,
  "返送汚泥ポンプ__No2__吐出圧力": 48,
  "返送汚泥ポンプ__No2__異音、振動": 49,
  "返送汚泥ポンプ__common1__モード確認": 50,
  "返送汚泥ポンプ__common1__常用機選択": 51,
  "返送汚泥ポンプ__No3__電流値": 52,
  "返送汚泥ポンプ__No3__吐出圧力": 53,
  "返送汚泥ポンプ__No3__異音、振動": 54,
  "返送汚泥ポンプ__No4__電流値": 55,
  "返送汚泥ポンプ__No4__吐出圧力": 56,
  "返送汚泥ポンプ__No4__異音、振動": 57,
  "返送汚泥ポンプ__common2__モード確認": 58,
  "返送汚泥ポンプ__common2__常用機選択": 59
};

// ─── 点検表 2/3 (汚泥処理系) ─── W1基準の絶対行（50項目）
var MECH2_ROW_W1 = {
  // 余剰汚泥ポンプ
  "余剰汚泥ポンプ__1__電流値": 64,
  "余剰汚泥ポンプ__1__吐出圧力": 65,
  "余剰汚泥ポンプ__1__異音、振動": 66,
  "余剰汚泥ポンプ__1__液漏れ": 67,
  "余剰汚泥ポンプ__1__モード確認": 68,
  "余剰汚泥ポンプ__2__電流値": 69,
  "余剰汚泥ポンプ__2__吐出圧力": 70,
  "余剰汚泥ポンプ__2__異音、振動": 71,
  "余剰汚泥ポンプ__2__液漏れ": 72,
  "余剰汚泥ポンプ__2__モード確認": 73,
  // 終沈床排水ポンプ
  "終沈床排水ポンプ__common1__電流値": 74,
  "終沈床排水ポンプ__common1__吐出圧力": 75,
  "終沈床排水ポンプ__common1__異音、振動": 76,
  "終沈床排水ポンプ__common1__モード確認": 77,
  // 返流水ポンプ
  "返流水ポンプ__No1__吐出圧力": 78,
  "返流水ポンプ__No1__異音、振動": 79,
  "返流水ポンプ__No2__吐出圧力": 80,
  "返流水ポンプ__No2__異音、振動": 81,
  "返流水ポンプ__common1__モード確認": 82,
  "返流水ポンプ__common1__機種選択": 83,
  // 濃縮汚泥引抜ポンプ
  "濃縮汚泥引抜ポンプ__1__電流値": 84,
  "濃縮汚泥引抜ポンプ__1__吐出圧力": 85,
  "濃縮汚泥引抜ポンプ__1__異音、振動": 86,
  "濃縮汚泥引抜ポンプ__2__電流値": 87,
  "濃縮汚泥引抜ポンプ__2__吐出圧力": 88,
  "濃縮汚泥引抜ポンプ__2__異音、振動": 89,
  "濃縮汚泥引抜ポンプ__common1__モード確認": 90,
  "濃縮汚泥引抜ポンプ__common1__機種選択": 91,
  // 濃縮汚泥引抜弁
  "濃縮汚泥引抜弁__common1__電流値": 92,
  "濃縮汚泥引抜弁__common1__異音、振動": 93,
  "濃縮汚泥引抜弁__common1__動作状況": 94,
  "濃縮汚泥引抜弁__common1__モード確認": 95,
  // 汚泥ポンプ室床排水ポンプ
  "汚泥ポンプ室床排水ポンプ__common1__電流値": 96,
  "汚泥ポンプ室床排水ポンプ__common1__吐出圧力": 97,
  "汚泥ポンプ室床排水ポンプ__common1__異音、振動": 98,
  "汚泥ポンプ室床排水ポンプ__common1__モード確認": 99,
  // 濃縮汚泥掻寄機
  "濃縮汚泥掻寄機__電流値": 100,
  "濃縮汚泥掻寄機__異音、振動": 101,
  // スカムスキマー (汚泥側)
  "スカムスキマー__動作状況": 102,
  // 汚泥供給ポンプ
  "汚泥供給ポンプ__1__吐出圧力": 103,
  "汚泥供給ポンプ__1__異音、振動": 104,
  "汚泥供給ポンプ__1__液漏れ": 105,
  "汚泥供給ポンプ__2__吐出圧力": 106,
  "汚泥供給ポンプ__2__異音、振動": 107,
  "汚泥供給ポンプ__2__液漏れ": 108,
  // 汚泥貯留槽攪拌機
  "汚泥貯留槽攪拌機__common1__電流値": 109,
  "汚泥貯留槽攪拌機__common1__異音、振動": 110,
  "汚泥貯留槽攪拌機__common1__モード確認": 111,
  // 空気圧縮機
  "空気圧縮機__common1__異音、振動等": 112,
  "空気圧縮機__common1__安全弁等の動作確認": 113
};

// ─── 点検表 3/3 (補機・換気系) ─── W1基準の絶対行（49項目）
var MECH3_ROW_W1 = {
  "井水揚水ポンプ__common1__異音、振動等": 120,
  "自動給水装置__1__電流値": 121,
  "自動給水装置__1__異音、振動等": 122,
  "自動給水装置__2__電流値": 123,
  "自動給水装置__2__異音、振動等": 124,
  "自動給水装置__common1__モード確認": 125,
  "井水貯留槽__common1__水漏れ": 126,
  "混和池ＢＧ__common1__開閉": 127,
  "塩素接触装置__common1__薬剤残量確認": 128,
  "臭気ガス吸引装置__common1__ボルト・ナットの緩み": 129,
  "臭気ガス吸引装置__common1__モード確認": 130,
  "吸着脱装置__common1__マノメータ等の作動確認": 131,
  "吸着脱装置__common1__ボルト・ナットの緩み": 132,
  "電気室給気ファン__common1__電流値": 133,
  "電気室給気ファン__common1__異音、振動等": 134,
  "電気室給気ファン__common1__モード確認": 135,
  "電気室排気ファン__common1__電流値": 136,
  "電気室排気ファン__common1__異音、振動等": 137,
  "電気室排気ファン__common1__モード確認": 138,
  "自家発室給気ファン__common1__異音、振動": 139,
  "自家発室給気ファン__common1__モード確認": 140,
  "汚泥ポンプ室排気ファン__common1__電流値": 141,
  "汚泥ポンプ室排気ファン__common1__異音、振動": 142,
  "汚泥ポンプ室排気ファン__common1__モード確認": 143,
  "汚泥ポンプ室吸気ファン__common1__電流値": 144,
  "汚泥ポンプ室吸気ファン__common1__異音、振動": 145,
  "汚泥ポンプ室吸気ファン__common1__モード確認": 146,
  "脱水作業室排気ファン__common1__電流値": 147,
  "脱水作業室排気ファン__common1__異音、振動": 148,
  "脱水作業室排気ファン__common1__モード確認": 149,
  "脱水作業室吸気ファン__common1__電流値": 150,
  "脱水作業室吸気ファン__common1__異音、振動": 151,
  "脱水作業室吸気ファン__common1__モード確認": 152,
  "脱臭機械室排気ファン__common1__電流値": 153,
  "脱臭機械室排気ファン__common1__異音、振動": 154,
  "脱臭機械室排気ファン__common1__モード確認": 155,
  "脱臭機械室吸気ファン__common1__電流値": 156,
  "脱臭機械室吸気ファン__common1__異音、振動": 157,
  "脱臭機械室吸気ファン__common1__モード確認": 158,
  "終沈ポンプ室排気ファン__common1__電流値": 159,
  "終沈ポンプ室排気ファン__common1__異音、振動": 160,
  "終沈ポンプ室排気ファン__common1__モード確認": 161,
  "終沈ポンプ室吸気ファン__common1__電流値": 162,
  "終沈ポンプ室吸気ファン__common1__異音、振動": 163,
  "終沈ポンプ室吸気ファン__common1__モード確認": 164,
  "自動巻取型エアフィルター__common1__巻取完了表示点灯確認": 165,
  "自動巻取型エアフィルター__common1__差圧": 166,
  "自動巻取型エアフィルター__common1__モード確認": 167,
  "自動巻取型エアフィルター__common1__設定": 168
};

// 旧名互換エイリアス（コード上の他参照向け）
var MECH1_ROW = MECH1_ROW_W1;
var MECH2_ROW = MECH2_ROW_W1;
var MECH3_ROW = MECH3_ROW_W1;

// W1基準行 + 週オフセット で絶対行を返す
//   week: 1〜5
//   subId: 'sub1' | 'sub2' | 'sub3'  (互換: 1|2|3, 'mech1'..)
//   itemKey: マッピングのキー (例: "汚水ポンプ__1__電流値")
function getMechRow(week, subId, itemKey) {
  var w = parseInt(week, 10);
  if (isNaN(w) || w < 1 || w > 5) throw new Error('Invalid week: ' + week);
  var map;
  var s = String(subId);
  if (s === 'sub1' || s === '1' || s === 'mech1') map = MECH1_ROW_W1;
  else if (s === 'sub2' || s === '2' || s === 'mech2') map = MECH2_ROW_W1;
  else if (s === 'sub3' || s === '3' || s === 'mech3') map = MECH3_ROW_W1;
  else throw new Error('Unknown mech subId: ' + subId);
  var baseRow = map[itemKey];
  if (baseRow == null) return null; // キー未登録は no-op（フォーム側との同期問題を握りつぶす）
  return baseRow + WEEK_OFFSETS_MECH[w - 1];
}

// ===== 公開アクション（認証不要） =====
var PUBLIC_ACTIONS = { 'ping':1, 'login':1, 'verifyToken':1 };


// ===================================================================
// ===== ユーティリティ =====
// ===================================================================

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function normalizeMonth(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2);
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  m = s.match(/^(\d{4})年(\d{1,2})月$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  m = s.match(/^(\d{1,2})月$/);
  if (m) {
    var y = new Date().getFullYear();
    return y + '-' + ('0' + m[1]).slice(-2);
  }
  return s;
}

function toDateStr(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2);
  }
  return String(v);
}

function matchMonth(cellValue, month) {
  if (cellValue === month) return true;
  if (cellValue instanceof Date) { var y=cellValue.getFullYear(),m=('0'+(cellValue.getMonth()+1)).slice(-2); return (y+'-'+m)===month; }
  var norm = normalizeMonth(cellValue);
  return norm === month || String(cellValue) === month;
}

// 月内の各週月曜日を計算
function computeWeekMondays(year, monthNum) {
  var first = new Date(year, monthNum - 1, 1);
  var dow = first.getDay(); // 0=日, 1=月, ..., 6=土
  var diff = (dow === 0) ? -6 : (1 - dow);
  var w1Mon = new Date(year, monthNum - 1, 1 + diff);
  var result = {};
  for(var w = 1; w <= 5; w++){
    var d = new Date(w1Mon.getFullYear(), w1Mon.getMonth(), w1Mon.getDate() + (w - 1) * 7);
    if (d <= new Date(year, monthNum, 0)) result[w] = d;
  }
  return result;
}

function formatReiwaDateRange(startDate, endDate) {
  var ry1 = startDate.getFullYear() - 2018;
  var ry2 = endDate.getFullYear() - 2018;
  return '令和' + ry1 + '年' + (startDate.getMonth() + 1) + '月' + startDate.getDate() + '日 ～ ' +
         '令和' + ry2 + '年' + (endDate.getMonth() + 1) + '月' + endDate.getDate() + '日';
}

// (week, day) → 当該日の Date を返す
function dateOfWeekDay(weekMondays, week, day) {
  var w = parseInt(week);
  if (!weekMondays[w]) return null;
  var off = DAY_OFFSETS[day];
  if (off === undefined) return null;
  var base = weekMondays[w];
  return new Date(base.getFullYear(), base.getMonth(), base.getDate() + off);
}

// 前月最終金曜の返送汚泥流量読み値を取得
function getPreviousMonthLastFridayValues(year, monthNum) {
  var prevDate = new Date(year, monthNum - 2, 1);
  var py = prevDate.getFullYear();
  var pm = ('0' + (prevDate.getMonth() + 1)).slice(-2);
  var prevMonth = py + '-' + pm;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 上段使用量 SS-D シートから前月の最終週金曜の値を取得
  var sheet = ss.getSheetByName('日常水質_上段');
  if (!sheet || sheet.getLastRow() <= 1) return { no1: null, no2: null };
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  // (month, week, json, updatedAt)
  var byWeek = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!matchMonth(row[0], prevMonth)) continue;
    var w = parseInt(row[1]);
    if (!w) continue;
    try {
      var json = JSON.parse(row[2] || '{}');
      var rs1 = json.returnSludge1 && json.returnSludge1.fri;
      var rs2 = json.returnSludge2 && json.returnSludge2.fri;
      if (rs1 !== undefined && rs1 !== null && rs1 !== '') {
        byWeek[w] = byWeek[w] || {}; byWeek[w].no1 = Number(rs1);
      }
      if (rs2 !== undefined && rs2 !== null && rs2 !== '') {
        byWeek[w] = byWeek[w] || {}; byWeek[w].no2 = Number(rs2);
      }
    } catch (e) {}
  }
  for (var w2 = 5; w2 >= 1; w2--) {
    if (byWeek[w2] && (byWeek[w2].no1 != null || byWeek[w2].no2 != null)) {
      return { no1: byWeek[w2].no1 != null ? byWeek[w2].no1 : null, no2: byWeek[w2].no2 != null ? byWeek[w2].no2 : null };
    }
  }
  return { no1: null, no2: null };
}

// 前月の電気設備データから、5項目の使用量計算用「読み」値を最終日付順に取得
// 戻り値: { [readKey]: lastReadingValue, ... }
function getPreviousMonthLastElecReadings(year, monthNum) {
  var prevDate = new Date(year, monthNum - 2, 1, 12, 0, 0);
  var py = prevDate.getFullYear();
  var pm = ('0' + (prevDate.getMonth() + 1)).slice(-2);
  var prevMonth = py + '-' + pm;
  var prevMonthIdx = prevDate.getMonth();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('日常電気設備');
  if (!sheet || sheet.getLastRow() <= 1) return {};

  var prevWeekMondays = computeWeekMondays(py, prevDate.getMonth() + 1);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  var entries = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!matchMonth(row[0], prevMonth)) continue;
    var w = parseInt(row[1]);
    var dayKey = DAY_REVERSE[row[2]];
    if (!dayKey) continue;
    var date = dateOfWeekDay(prevWeekMondays, w, dayKey);
    if (!date || date.getMonth() !== prevMonthIdx) continue;
    var json = {};
    try { json = JSON.parse(row[3] || '{}'); } catch(e) { continue; }
    entries.push({ date: date, json: json });
  }
  entries.sort(function(a, b){ return b.date.getTime() - a.date.getTime(); });

  var KEYS = [
    "引込・受電__普通電力量(×10kw)__読み",
    "低圧分岐__沈砂池ポンプ設備__読み",
    "低圧分岐__水処理設備(1)__読み",
    "低圧分岐__水処理設備(2)__読み",
    "低圧分岐__汚泥処理設備電力量__読み"
  ];
  var result = {};
  for (var k = 0; k < KEYS.length; k++) {
    for (var j = 0; j < entries.length; j++) {
      var v = entries[j].json[KEYS[k]];
      if (v != null && v !== '' && !isNaN(parseFloat(v))) {
        result[KEYS[k]] = parseFloat(v);
        break;
      }
    }
  }
  return result;
}


// ===================================================================
// ===== 認証システム =====
// ===================================================================

function setupAuth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ユーザー管理');
  if (sheet) return 'すでに「ユーザー管理」シートは存在します。';
  sheet = ss.insertSheet('ユーザー管理');
  sheet.getRange(1, 1, 1, 6).setValues([['ID', 'パスワードハッシュ', '氏名', '有効', '最終ログイン', '備考']]);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e0f2e9');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 200);
  sheet.getRange(2, 1, 1, 6).setValues([['cpkanri', '', '管理者', true, '', '初期管理者（setAdminPassword()でパスワード設定）']]);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('AUTH_SALT')) {
    var salt = Utilities.getUuid() + '-' + new Date().getTime();
    props.setProperty('AUTH_SALT', salt);
  }
  return '「ユーザー管理」シートを作成しました。次に setAdminPassword("任意のパスワード") を実行してください。';
}

function setAdminPassword(password) {
  if (!password || password.length < 6) return 'パスワードは6文字以上を指定してください。';
  return setUserPassword('cpkanri', password);
}

function setUserPassword(userId, password) {
  if (!userId || !password || password.length < 6) return 'IDとパスワード（6文字以上）を指定してください。';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ユーザー管理');
  if (!sheet) return 'ユーザー管理シートがありません。setupAuth() を先に実行してください。';
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === userId) {
      var hash = hashPassword(password);
      sheet.getRange(i + 1, 2).setValue(hash);
      return 'ユーザー「' + userId + '」のパスワードを設定しました。';
    }
  }
  return 'ユーザー「' + userId + '」が見つかりません。';
}

function hashPassword(password) {
  var salt = PropertiesService.getScriptProperties().getProperty('AUTH_SALT') || 'default-salt';
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < raw.length; i++) {
    var b = raw[i] & 0xff;
    hex += ('0' + b.toString(16)).slice(-2);
  }
  return hex;
}

function login(data) {
  var userId = data.userId;
  var password = data.password;
  if (!userId || !password) return jsonResponse({error:'IDとパスワードを入力してください'});
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ユーザー管理');
  if (!sheet) return jsonResponse({error:'認証システム未初期化（管理者に連絡してください）'});
  var rows = sheet.getDataRange().getValues();
  var hash = hashPassword(password);
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][0]);
    var storedHash = String(rows[i][1]);
    var name = String(rows[i][2] || '');
    var enabled = rows[i][3];
    if (id !== userId) continue;
    if (!enabled || enabled === false || String(enabled).toLowerCase() === 'false') {
      return jsonResponse({error:'このアカウントは無効化されています'});
    }
    if (!storedHash || storedHash !== hash) {
      return jsonResponse({error:'IDまたはパスワードが正しくありません'});
    }
    sheet.getRange(i + 1, 5).setValue(new Date());
    var token = issueToken(userId);
    return jsonResponse({status: 'ok', token: token.token, expires: token.expires, userId: userId, name: name});
  }
  return jsonResponse({error:'IDまたはパスワードが正しくありません'});
}

function issueToken(userId) {
  var props = PropertiesService.getScriptProperties();
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var now = new Date().getTime();
  var expires = now + 30 * 24 * 60 * 60 * 1000;
  var tokenKey = 'TK_' + token;
  props.setProperty(tokenKey, JSON.stringify({userId: userId, expires: expires}));
  try { cleanupExpiredTokens(); } catch (e) {}
  return { token: token, expires: expires };
}

function cleanupExpiredTokens() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var now = new Date().getTime();
  for (var key in all) {
    if (key.indexOf('TK_') !== 0) continue;
    try {
      var obj = JSON.parse(all[key]);
      if (obj.expires < now) props.deleteProperty(key);
    } catch (e) { props.deleteProperty(key); }
  }
}

function verifyToken(token) {
  if (!token) return null;
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('TK_' + token);
  if (!raw) return null;
  try {
    var obj = JSON.parse(raw);
    if (obj.expires < new Date().getTime()) {
      props.deleteProperty('TK_' + token);
      return null;
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ユーザー管理');
    if (!sheet) return null;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === obj.userId) {
        var enabled = rows[i][3];
        if (!enabled || enabled === false || String(enabled).toLowerCase() === 'false') {
          props.deleteProperty('TK_' + token);
          return null;
        }
        return obj.userId;
      }
    }
    props.deleteProperty('TK_' + token);
    return null;
  } catch (e) { return null; }
}

function logout(data) {
  var token = data && data.token;
  if (token) PropertiesService.getScriptProperties().deleteProperty('TK_' + token);
  return jsonResponse({status:'ok'});
}

function verifyTokenApi(token) {
  var userId = verifyToken(token);
  if (userId) return jsonResponse({valid:true, userId:userId});
  return jsonResponse({valid:false});
}

function requireAuth(action, token) {
  if (PUBLIC_ACTIONS[action]) return null;
  var userId = verifyToken(token);
  if (!userId) return jsonResponse({error:'認証が必要です', authRequired:true});
  return null;
}


// ===================================================================
// ===== 祝日リスト =====
// ===================================================================

function setupHolidays() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('祝日リスト');
  if (sheet) return '既に「祝日リスト」シートが存在します。';
  sheet = ss.insertSheet('祝日リスト');
  sheet.getRange(1, 1, 2, 2).setValues([['祝日一覧', ''], ['日付', '名称']]);
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.getRange(2, 1, 1, 2).setFontWeight('bold').setBackground('#e0f2e9');
  var holidays = [
    ['2025-04-29','昭和の日'],['2025-05-03','憲法記念日'],['2025-05-04','みどりの日'],['2025-05-05','こどもの日'],['2025-05-06','振替休日'],
    ['2025-07-21','海の日'],['2025-08-11','山の日'],['2025-09-15','敬老の日'],['2025-09-23','秋分の日'],['2025-10-13','スポーツの日'],['2025-11-03','文化の日'],['2025-11-23','勤労感謝の日'],['2025-11-24','振替休日'],
    ['2026-01-01','元日'],['2026-01-02','振替休日'],['2026-01-12','成人の日'],['2026-02-11','建国記念の日'],['2026-02-23','天皇誕生日'],['2026-03-20','春分の日'],
    ['2026-04-29','昭和の日'],['2026-05-03','憲法記念日'],['2026-05-04','みどりの日'],['2026-05-05','こどもの日'],['2026-05-06','振替休日'],['2026-07-20','海の日'],['2026-08-11','山の日'],['2026-09-21','敬老の日'],['2026-09-23','秋分の日'],['2026-10-12','スポーツの日'],['2026-11-03','文化の日'],['2026-11-23','勤労感謝の日'],
    ['2027-01-01','元日'],['2027-01-11','成人の日'],['2027-02-11','建国記念の日'],['2027-02-23','天皇誕生日'],['2027-03-21','春分の日'],['2027-03-22','振替休日'],['2027-04-29','昭和の日'],['2027-05-03','憲法記念日'],['2027-05-04','みどりの日'],['2027-05-05','こどもの日'],['2027-07-19','海の日'],['2027-08-11','山の日'],['2027-09-20','敬老の日'],['2027-09-23','秋分の日'],['2027-10-11','スポーツの日'],['2027-11-03','文化の日'],['2027-11-23','勤労感謝の日'],
    ['2028-01-01','元日'],['2028-01-10','成人の日'],['2028-02-11','建国記念の日'],['2028-02-23','天皇誕生日'],['2028-03-20','春分の日'],['2028-04-29','昭和の日'],['2028-05-03','憲法記念日'],['2028-05-04','みどりの日'],['2028-05-05','こどもの日'],['2028-07-17','海の日'],['2028-08-11','山の日'],['2028-09-18','敬老の日'],['2028-09-22','秋分の日'],['2028-10-09','スポーツの日'],['2028-11-03','文化の日'],['2028-11-23','勤労感謝の日']
  ];
  var rows = holidays.map(function(h) {
    var p = h[0].split('-').map(Number);
    return [new Date(p[0], p[1] - 1, p[2]), h[1]];
  });
  sheet.getRange(3, 1, rows.length, 2).setValues(rows);
  sheet.getRange(3, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  return '「祝日リスト」シートを作成しました（' + rows.length + '件）。';
}

function getHolidays() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('祝日リスト');
    if (!sheet) return jsonResponse({holidays: []});
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return jsonResponse({holidays: []});
    var data = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var d = data[i][0];
      if (!d) continue;
      if (d instanceof Date) {
        var y = d.getFullYear();
        var m = ('0' + (d.getMonth() + 1)).slice(-2);
        var day = ('0' + d.getDate()).slice(-2);
        result.push({date: y + '-' + m + '-' + day, name: data[i][1] || ''});
      }
    }
    return jsonResponse({holidays: result});
  } catch (err) { return jsonResponse({error: err.message}); }
}


// ===================================================================
// ===== HTTP エンドポイント =====
// ===================================================================

function doGet(e) {
  try {
    var a = e.parameter.action;
    var token = e.parameter.token || '';
    var authFail = requireAuth(a, token);
    if (authFail) return authFail;

    if (a==='ping') return jsonResponse({status:'ok',message:'筒賀週報API稼働中 v2'});
    if (a==='verifyToken') return verifyTokenApi(token);
    if (a==='getAllData') return getAllData(e.parameter.month);
    if (a==='exportExcel') return exportExcel(e.parameter.month, e.parameter.templateId);
    if (a==='getHolidays') return getHolidays();
    return jsonResponse({error:'不明なアクション'});
  } catch(err) { return jsonResponse({error:err.message}); }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents), a = data.action;
    var token = data.token || '';
    var authFail = requireAuth(a, token);
    if (authFail) return authFail;

    if (a==='login') return login(data);
    if (a==='logout') return logout(data);
    if (a==='saveWaterMeasure') return saveWaterMeasureData(data);
    if (a==='saveWaterUsage') return saveWaterUsageData(data);
    if (a==='saveEquipment') return saveEquipmentData(data);
    if (a==='saveElectrical') return saveElectricalData(data);
    if (a==='saveMechanical') return saveMechanicalData(data);
    if (a==='saveEverything') return saveEverything(data);
    return jsonResponse({error:'不明なアクション'});
  } catch(err) { return jsonResponse({error:err.message}); }
}


// ===================================================================
// ===== データ取得・保存（SS-D） =====
// ===================================================================

// 保存先シート名（SS-D 上に存在する想定。なければ自動作成）
function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function findRowByKey(sheet, keyCols, keyVals) {
  var lr = sheet.getLastRow(); if (lr <= 1) return -1;
  var d = sheet.getRange(2, 1, lr - 1, Math.max(keyCols.length, sheet.getLastColumn())).getValues();
  for (var i = 0; i < d.length; i++) {
    var ok = true;
    for (var k = 0; k < keyCols.length; k++) {
      var ci = keyCols[k] - 1;
      var rv = d[i][ci];
      var kv = keyVals[k];
      if (k === 0) { // month col
        if (!matchMonth(rv, kv)) { ok = false; break; }
      } else {
        if (typeof rv === 'number' && typeof kv === 'number') {
          if (rv !== kv) { ok = false; break; }
        } else if (String(rv) !== String(kv)) { ok = false; break; }
      }
    }
    if (ok) return i + 2;
  }
  return -1;
}

// ===== 下段水質測定の保存（per-day, JSON） =====
function saveWaterMeasureData(data) {
  var month = data.month, week = data.week, days = data.days;
  if (!month || !week || !days) return jsonResponse({error:'パラメータ不足'});
  var sh = getOrCreateSheet('日常水質_下段', ['月','週','曜日','日付','データJSON','更新日時']);
  var now = new Date(), cnt = 0;
  for (var dc in days) {
    var dd = days[dc]; if (!dd) continue;
    var dn = DAY_NAMES[dc]; if (!dn) continue;
    var ri = findRowByKey(sh, [1,2,3], [month, parseInt(week), dn]);
    var json = JSON.stringify(dd);
    var rv = [month, parseInt(week), dn, toDateStr(dd.date) || '', json, now];
    if (ri > 0) sh.getRange(ri, 1, 1, rv.length).setValues([rv]); else sh.appendRow(rv);
    cnt++;
  }
  return jsonResponse({status:'ok', message:'水質下段'+cnt+'日分保存', savedCount: cnt});
}

// ===== 上段使用量の保存（per-week, JSON） =====
function saveWaterUsageData(data) {
  var month = data.month, weeks = data.weeks;
  if (!month || !weeks) return jsonResponse({error:'パラメータ不足'});
  var sh = getOrCreateSheet('日常水質_上段', ['月','週','データJSON','更新日時']);
  var now = new Date(), cnt = 0;
  for (var wn in weeks) {
    var wd = weeks[wn]; if (!wd) continue;
    var ri = findRowByKey(sh, [1,2], [month, parseInt(wn)]);
    var json = JSON.stringify(wd);
    var rv = [month, parseInt(wn), json, now];
    if (ri > 0) sh.getRange(ri, 1, 1, rv.length).setValues([rv]); else sh.appendRow(rv);
    cnt++;
  }
  return jsonResponse({status:'ok', message:'水質上段'+cnt+'週分保存', savedCount: cnt});
}

// ===== 機器運転時間の保存（22機器、列ベース） =====
function saveEquipmentData(data) {
  var month = data.month, eq = data.equipment;
  if (!month || !eq) return jsonResponse({error:'パラメータ不足'});
  var sh = getOrCreateSheet('機器運転時間', ['月','週','機器名','月','火','水','木','金','更新日時']);
  var now = new Date(), cnt = 0;
  for (var wn in eq) {
    var wd = eq[wn]; if (!wd) continue;
    for (var i = 0; i < EQUIP_NAMES.length; i++) {
      var ek = EQUIP_NAMES[i], el = EQUIP_LABELS[i], ed = wd[ek];
      if (!ed) continue;
      var ri = findRowByKey(sh, [1,2,3], [month, parseInt(wn), el]);
      var rv = [month, parseInt(wn), el,
                ed.mon != null ? ed.mon : '',
                ed.tue != null ? ed.tue : '',
                ed.wed != null ? ed.wed : '',
                ed.thu != null ? ed.thu : '',
                ed.fri != null ? ed.fri : '',
                now];
      if (ri > 0) sh.getRange(ri, 1, 1, rv.length).setValues([rv]); else sh.appendRow(rv);
      cnt++;
    }
  }
  return jsonResponse({status:'ok', message:'機器'+cnt+'件保存', savedCount: cnt});
}

// ===== 電気設備の保存（per-day, JSON） =====
function saveElectricalData(data) {
  var month = data.month, el = data.electrical;
  if (!month || !el) return jsonResponse({error:'パラメータ不足'});
  var sh = getOrCreateSheet('日常電気設備', ['月','週','曜日','データJSON','更新日時']);
  var now = new Date(), cnt = 0;
  for (var wn in el) {
    var wd = el[wn]; if (!wd) continue;
    for (var dc in wd) {
      var dd = wd[dc]; if (!dd) continue;
      var dn = DAY_NAMES[dc]; if (!dn) continue;
      var ri = findRowByKey(sh, [1,2,3], [month, parseInt(wn), dn]);
      var rv = [month, parseInt(wn), dn, JSON.stringify(dd), now];
      if (ri > 0) sh.getRange(ri, 1, 1, rv.length).setValues([rv]); else sh.appendRow(rv);
      cnt++;
    }
  }
  return jsonResponse({status:'ok', message:'電気'+cnt+'日分保存', savedCount: cnt});
}

// ===== 機械設備の保存（per-day, JSON） =====
function saveMechanicalData(data) {
  var month = data.month, mech = data.mechanical;
  if (!month || !mech) return jsonResponse({error:'パラメータ不足'});
  var sh = getOrCreateSheet('日常機械設備', ['月','週','曜日','データJSON','更新日時']);
  var now = new Date(), cnt = 0;
  for (var wn in mech) {
    var wd = mech[wn]; if (!wd) continue;
    for (var dc in wd) {
      var dd = wd[dc]; if (!dd) continue;
      var dn = DAY_NAMES[dc]; if (!dn) continue;
      var ri = findRowByKey(sh, [1,2,3], [month, parseInt(wn), dn]);
      var rv = [month, parseInt(wn), dn, JSON.stringify(dd), now];
      if (ri > 0) sh.getRange(ri, 1, 1, rv.length).setValues([rv]); else sh.appendRow(rv);
      cnt++;
    }
  }
  return jsonResponse({status:'ok', message:'機械'+cnt+'日分保存', savedCount: cnt});
}

function saveEverything(data) {
  var results = [], month = data.month;
  if (data.waterUsage) {
    var r = saveWaterUsageData({month:month, weeks:data.waterUsage});
    results.push(JSON.parse(r.getContent()).message);
  }
  if (data.waterMeasure) {
    // weeks: { 1: { mon:{...}, tue:{...}, ...}, ...}
    for (var wn in data.waterMeasure) {
      var r = saveWaterMeasureData({month:month, week:parseInt(wn), days:data.waterMeasure[wn]});
      // (累積メッセージは省略)
    }
    results.push('水質下段保存OK');
  }
  if (data.equipment) {
    var r = saveEquipmentData({month:month, equipment:data.equipment});
    results.push(JSON.parse(r.getContent()).message);
  }
  if (data.electrical) {
    var r = saveElectricalData({month:month, electrical:data.electrical});
    results.push(JSON.parse(r.getContent()).message);
  }
  if (data.mechanical) {
    var r = saveMechanicalData({month:month, mechanical:data.mechanical});
    results.push(JSON.parse(r.getContent()).message);
  }
  return jsonResponse({status:'ok', message:results.join('、')});
}

// ===== 全データ取得 =====
function getAllData(month) {
  if (!month) return jsonResponse({error:'月が指定されていません'});
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {month:month, waterUsage:{}, waterMeasure:{}, equipment:{}, electrical:{}, mechanical:{}};

  var ws = ss.getSheetByName('日常水質_上段');
  if (ws && ws.getLastRow() > 1) {
    var v = ws.getRange(2,1,ws.getLastRow()-1,4).getValues();
    for (var i = 0; i < v.length; i++) {
      if (matchMonth(v[i][0], month)) {
        try { result.waterUsage[v[i][1]] = JSON.parse(v[i][2] || '{}'); } catch(e) {}
      }
    }
  }
  var wsd = ss.getSheetByName('日常水質_下段');
  if (wsd && wsd.getLastRow() > 1) {
    var v = wsd.getRange(2,1,wsd.getLastRow()-1,6).getValues();
    for (var i = 0; i < v.length; i++) {
      if (matchMonth(v[i][0], month)) {
        var dc = DAY_REVERSE[v[i][2]]; if (!dc) continue;
        var w = v[i][1];
        result.waterMeasure[w] = result.waterMeasure[w] || {};
        try { result.waterMeasure[w][dc] = JSON.parse(v[i][4] || '{}'); } catch(e) {}
        if (v[i][3]) result.waterMeasure[w][dc].date = toDateStr(v[i][3]);
      }
    }
  }
  var es = ss.getSheetByName('機器運転時間');
  if (es && es.getLastRow() > 1) {
    var v = es.getRange(2,1,es.getLastRow()-1,9).getValues();
    for (var i = 0; i < v.length; i++) {
      if (matchMonth(v[i][0], month)) {
        var w = v[i][1], lbl = v[i][2];
        var ek = EQUIP_LABELS.indexOf(lbl);
        if (ek < 0) continue;
        result.equipment[w] = result.equipment[w] || {};
        result.equipment[w][EQUIP_NAMES[ek]] = {mon:v[i][3], tue:v[i][4], wed:v[i][5], thu:v[i][6], fri:v[i][7]};
      }
    }
  }
  var els = ss.getSheetByName('日常電気設備');
  if (els && els.getLastRow() > 1) {
    var v = els.getRange(2,1,els.getLastRow()-1,5).getValues();
    for (var i = 0; i < v.length; i++) {
      if (matchMonth(v[i][0], month)) {
        var dc = DAY_REVERSE[v[i][2]]; if (!dc) continue;
        var w = v[i][1];
        result.electrical[w] = result.electrical[w] || {};
        try { result.electrical[w][dc] = JSON.parse(v[i][3] || '{}'); } catch(e) {}
      }
    }
  }
  var ms = ss.getSheetByName('日常機械設備');
  if (ms && ms.getLastRow() > 1) {
    var v = ms.getRange(2,1,ms.getLastRow()-1,5).getValues();
    for (var i = 0; i < v.length; i++) {
      if (matchMonth(v[i][0], month)) {
        var dc = DAY_REVERSE[v[i][2]]; if (!dc) continue;
        var w = v[i][1];
        result.mechanical[w] = result.mechanical[w] || {};
        try { result.mechanical[w][dc] = JSON.parse(v[i][3] || '{}'); } catch(e) {}
      }
    }
  }
  return jsonResponse(result);
}


// ===================================================================
// ===== Excel出力 =====
// ===================================================================

function exportExcel(month, templateId) {
  if (!month) return jsonResponse({error:'月が指定されていません'});
  if (!templateId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), st = ss.getSheetByName('設定');
    if (st) {
      var sd = st.getDataRange().getValues();
      for (var i = 0; i < sd.length; i++) {
        if (sd[i][0] === 'テンプレートID') { templateId = sd[i][1]; break; }
      }
    }
  }
  if (!templateId) templateId = SPREADSHEET_ID;
  if (!templateId) return jsonResponse({error:'テンプレートIDなし'});

  try {
    var tf = DriveApp.getFileById(templateId);
    var ml = month.replace('-','年') + '月';
    var nf = '筒賀' + ml.substring(5);
    var cf = tf.makeCopy(nf), nid = cf.getId();
    var wb = SpreadsheetApp.openById(nid);
    var sheets = wb.getSheets();

    var ym = month.split('-');
    var year = parseInt(ym[0]), monthNum = parseInt(ym[1]);
    var weekMondays = computeWeekMondays(year, monthNum);

    var waterSheet = sheets[SHEET_IDX.daily_water];
    var opSheet    = sheets[SHEET_IDX.operation_monthly];
    var reportSheet = sheets[SHEET_IDX.water_report];
    var mechSheet = sheets[SHEET_IDX.daily_mech];
    var elecSheet = sheets[SHEET_IDX.daily_elec];
    var eqSheet   = sheets[SHEET_IDX.equipment_hours];

    // ======= データを SS-D から取り込む =======
    var allDataRes = JSON.parse(getAllData(month).getContent());
    var waterUsage   = allDataRes.waterUsage   || {};
    var waterMeasure = allDataRes.waterMeasure || {};
    var equipment    = allDataRes.equipment    || {};
    var electrical   = allDataRes.electrical   || {};
    var mechanical   = allDataRes.mechanical   || {};

    // ======= [1] 日常水質シート（上段使用量 + 下段測定） =======
    for (var w = 1; w <= 5; w++) {
      var blockRow = WATER_BLOCK_ROW[w]; if (!blockRow) continue;
      var monDay = weekMondays[w];

      // 第1週の月曜日付を I3 に書く（全週分の参照元）
      if (w === 1 && monDay) {
        waterSheet.getRange(3, 9).setValue(monDay);
        // 月年タイトルを M1 に書き込み（例: "令和8年5月"）
        waterSheet.getRange('M1').setValue('令和' + (year - 2018) + '年' + monthNum + '月');
      }

      // 上段使用量（per-week）
      var usage = waterUsage[w];
      if (usage) {
        for (var key in USAGE_ROW_OFFSET) {
          var item = usage[key]; if (!item) continue;
          var rowAbs = blockRow + USAGE_ROW_OFFSET[key];
          for (var k = 0; k < USAGE_CELLS.length; k++) {
            var cell = USAGE_CELLS[k];
            var v = item[cell];
            if (v != null && v !== '') {
              waterSheet.getRange(rowAbs, USAGE_DAY_COL[cell]).setValue(v);
            }
          }
        }
      }

      // 下段水質測定（per-day）
      var weekMeas = waterMeasure[w] || {};
      for (var di = 0; di < DAYS.length; di++) {
        var day = DAYS[di], col = MEASURE_DAY_COL[day];
        var dd = weekMeas[day]; if (!dd) continue;
        for (var fk in MEASURE_ROW_OFFSET) {
          var rv = dd[fk];
          if (rv != null && rv !== '') {
            waterSheet.getRange(blockRow + MEASURE_ROW_OFFSET[fk], col).setValue(rv);
          }
        }
        // 測定者
        if (dd.inspector) waterSheet.getRange(blockRow + WATER_INSPECTOR_ROW_OFFSET, col).setValue(dd.inspector);
        // 備考
        if (dd.bikou) waterSheet.getRange(blockRow + WATER_BIKOU_ROW_OFFSET, col).setValue(dd.bikou);
      }
    }

    // ======= [2] 水質管理報告 — 終沈3項目 + 重複項目 + DO =======
    for (var w = 1; w <= 5; w++) {
      var weekMeas = waterMeasure[w] || {};
      var weekElec = electrical[w] || {};
      for (var di = 0; di < DAYS.length; di++) {
        var day = DAYS[di];
        var date = dateOfWeekDay(weekMondays, w, day);
        if (!date || date.getMonth() !== monthNum - 1) continue; // 当月のみ
        var rowF = reportRowFront(date);  // 前半行
        var rowB = reportRowBack(date);   // 後半行

        // 重複（前半: 水温/透視度）+ 終沈（前半: 終沈水温/終沈透視度）
        var meas = weekMeas[day] || {};
        var frontMap = {
          tempIn: meas.tempIn, tempD1: meas.tempD1, tempD2: meas.tempD2, tempOut: meas.tempOut,
          tempFinal1: meas.tempFinal1, tempFinal2: meas.tempFinal2,
          transIn: meas.transIn, transOut: meas.transOut,
          transFinal1: meas.transFinal1, transFinal2: meas.transFinal2
        };
        for (var k in frontMap) {
          var v = frontMap[k];
          if (v != null && v !== '') reportSheet.getRange(rowF, REPORT_FRONT_COL[k]).setValue(v);
        }

        // 重複（後半: PH/残留塩素/SV30）+ 終沈PH No1/No2 + DO No1/No2
        var elec = weekElec[day] || {};
        var backMap = {
          phIn: meas.phIn, phD1: meas.phD1, phD2: meas.phD2, phOut: meas.phOut,
          phFinal1: meas.phFinal1, phFinal2: meas.phFinal2,
          chlorine: meas.chlorine,
          sv1_30: meas.sv1_30, sv2_30: meas.sv2_30,
          doNo1: elec[ELEC_DO_KEY_NO1], doNo2: elec[ELEC_DO_KEY_NO2]
        };
        for (var k in backMap) {
          var v = backMap[k];
          if (v != null && v !== '') reportSheet.getRange(rowB, REPORT_BACK_COL[k]).setValue(v);
        }
      }
    }

    // ======= [3] 運転管理月報 =======
    if (opSheet) {
      // 月初日付を A6 に書き込み → A7以降の日付列とB列(曜日)が数式チェーンで自動更新
      opSheet.getRange('A6').setValue(new Date(year, monthNum - 1, 1, 12, 0, 0));
      // 月年タイトルを J1 に書き込み（例: "令和8年　5月分"）
      opSheet.getRange('J1').setValue('令和' + (year - 2018) + '年　' + monthNum + '月分');
      // 前月最終金曜の返送汚泥（既存）
      try {
        var prev = getPreviousMonthLastFridayValues(year, monthNum);
        if (prev) {
          if (prev.no1 !== null && prev.no1 !== undefined) opSheet.getRange('J6').setValue(prev.no1);
          if (prev.no2 !== null && prev.no2 !== undefined) opSheet.getRange('L6').setValue(prev.no2);
        }
      } catch(prevErr) { Logger.log('prev-month fetch skipped: ' + prevErr.message); }
    }

    // ======= [4] 主要機器運転時間 =======
    if (eqSheet) {
      // 各週ヘッダー：A列に start date を書き込む
      for (var w = 1; w <= 5; w++) {
        var cell = EQUIP_WEEK_DATE_CELL[w];
        if (cell && weekMondays[w]) eqSheet.getRange(cell).setValue(weekMondays[w]);
      }
      for (var w = 1; w <= 5; w++) {
        var weekEq = equipment[w]; if (!weekEq) continue;
        var rowBase = EQUIP_WEEK_START[w]; if (!rowBase) continue;
        for (var i = 0; i < EQUIP_NAMES.length; i++) {
          var ed = weekEq[EQUIP_NAMES[i]]; if (!ed) continue;
          var trgRow = rowBase + i;
          for (var d = 0; d < DAYS.length; d++) {
            var day = DAYS[d];
            var date = dateOfWeekDay(weekMondays, w, day);
            if (!date || date.getMonth() !== monthNum - 1) continue; // 当月外スキップ
            var v = ed[day];
            if (v != null && v !== '') eqSheet.getRange(trgRow, EQUIP_HOUR_DAY_COL[day]).setValue(v);
          }
        }
      }
    }

    // ======= [5] 日常電気設備（点検表 1/3, 2/3, 3/3 + 余剰汚泥/燃料/DO転記） =======
    //   キー = "機器__識別子__項目" (Japanese)
    //   行 = ELEC{1,2,3}_ROW_W1[key] + WEEK_OFFSETS_ELEC[subKey][w-1]
    //     ※ サブタブごとに W1→W2 オフセットが異なる (sub1: +103, sub2: +101, sub3: +100)
    //   フォーム側で未登録のキーは getElecRow が null を返すので no-op
    if (elecSheet) {
      var ELEC_PAGES = [
        { sub: 'sub1', map: ELEC1_ROW_W1 },
        { sub: 'sub2', map: ELEC2_ROW_W1 },
        { sub: 'sub3', map: ELEC3_ROW_W1 }
      ];
      for (var w = 1; w <= 5; w++) {
        var weekEl = electrical[w] || {};
        var weekUs = waterUsage[w]  || {};
        for (var di = 0; di < DAYS.length; di++) {
          var day = DAYS[di];
          var date = dateOfWeekDay(weekMondays, w, day);
          if (!date || date.getMonth() !== monthNum - 1) continue; // 当月外スキップ
          var col = ELEC_DAY_COL[day];
          var dd = weekEl[day] || {};
          for (var pi = 0; pi < ELEC_PAGES.length; pi++) {
            var page = ELEC_PAGES[pi];
            var off = WEEK_OFFSETS_ELEC[page.sub][w - 1];
            for (var k in page.map) {
              var v = dd[k];
              // sub2: 余剰汚泥流量計 No.1/No.2 は上段使用量から自動転記
              if (page.sub === 'sub2') {
                if (k === ELEC_EXCESS_KEY_NO1 && weekUs.excessSludge1 && weekUs.excessSludge1[day] != null) v = weekUs.excessSludge1[day];
                else if (k === ELEC_EXCESS_KEY_NO2 && weekUs.excessSludge2 && weekUs.excessSludge2[day] != null) v = weekUs.excessSludge2[day];
              }
              // sub3: 燃料タンク残量は上段の軽油残量(diesel)から自動転記
              if (page.sub === 'sub3' && k === ELEC_FUEL_TANK_KEY && weekUs.diesel && weekUs.diesel[day] != null) {
                v = weekUs.diesel[day];
              }
              if (v != null && v !== '') elecSheet.getRange(page.map[k] + off, col).setValue(v);
            }
          }
        }
      }

      // === 第1週初日の使用量を前月最終読みから計算（5項目） ===
      try {
        var prevElecReadings = getPreviousMonthLastElecReadings(year, monthNum);
        var ELEC_USAGE_PAIRS = [
          { readKey: "引込・受電__普通電力量(×10kw)__読み",       usageRow: 9  },
          { readKey: "低圧分岐__沈砂池ポンプ設備__読み",           usageRow: 22 },
          { readKey: "低圧分岐__水処理設備(1)__読み",              usageRow: 24 },
          { readKey: "低圧分岐__水処理設備(2)__読み",              usageRow: 26 },
          { readKey: "低圧分岐__汚泥処理設備電力量__読み",         usageRow: 28 }
        ];
        var week1El = electrical[1] || {};
        var sub1Off = WEEK_OFFSETS_ELEC.sub1[0];
        for (var u = 0; u < ELEC_USAGE_PAIRS.length; u++) {
          var pair = ELEC_USAGE_PAIRS[u];
          var prevVal = prevElecReadings[pair.readKey];
          if (prevVal == null) continue;
          // 第1週で最初に当月かつ「読み」が記録されている日を探す
          for (var di2 = 0; di2 < DAYS.length; di2++) {
            var d2 = DAYS[di2];
            var dt = dateOfWeekDay(weekMondays, 1, d2);
            if (!dt || dt.getMonth() !== monthNum - 1) continue;
            var dd2 = week1El[d2];
            if (!dd2) continue;
            var currVal = dd2[pair.readKey];
            if (currVal == null || currVal === '') continue;
            var currNum = parseFloat(currVal);
            if (isNaN(currNum)) continue;
            elecSheet.getRange(pair.usageRow + sub1Off, ELEC_DAY_COL[d2]).setValue(currNum - prevVal);
            break;
          }
        }
      } catch (uErr) {
        Logger.log('elec usage calc skipped: ' + uErr.message);
      }
    }

    // ======= [6] 日常機械設備（点検表 1/3, 2/3, 3/3） =======
    //   キー = "機器__識別子__項目" (Japanese), 行は MECH{1,2,3}_ROW_W1 + WEEK_OFFSETS_MECH[w-1]
    //   フォーム側で未登録のキーは getMechRow が null を返すので no-op
    if (mechSheet) {
      var MECH_PAGE_MAPS = [MECH1_ROW_W1, MECH2_ROW_W1, MECH3_ROW_W1];
      for (var w = 1; w <= 5; w++) {
        var weekMe = mechanical[w] || {};
        var weekOff = WEEK_OFFSETS_MECH[w - 1];
        for (var di = 0; di < DAYS.length; di++) {
          var day = DAYS[di];
          var date = dateOfWeekDay(weekMondays, w, day);
          if (!date || date.getMonth() !== monthNum - 1) continue; // 当月外スキップ
          var col = MECH_DAY_COL[day];
          var dd = weekMe[day] || {};
          for (var mi = 0; mi < MECH_PAGE_MAPS.length; mi++) {
            var m = MECH_PAGE_MAPS[mi];
            for (var k in m) {
              var v = dd[k];
              if (v != null && v !== '') mechSheet.getRange(m[k] + weekOff, col).setValue(v);
            }
          }
        }
      }
    }

    SpreadsheetApp.flush();
    var eu = 'https://docs.google.com/spreadsheets/d/' + nid + '/export?format=xlsx';
    var tk = ScriptApp.getOAuthToken();
    var rsp = UrlFetchApp.fetch(eu, {headers:{'Authorization':'Bearer ' + tk}});
    var bl = rsp.getBlob();
    bl.setName(nf + '.xlsx');
    // テンプレコピーは不要なので削除（Driveに残さない）
    DriveApp.getFileById(nid).setTrashed(true);
    // base64でクライアントに直接返す（Drive経由のダウンロード不要）
    return jsonResponse({
      status:'ok',
      message: ml + 'のExcelファイルを生成しました',
      fileName: nf + '.xlsx',
      fileBase64: Utilities.base64Encode(bl.getBytes())
    });
  } catch(err) {
    return jsonResponse({error:'Excel生成エラー: ' + err.message});
  }
}


// ===================================================================
// ===== 後方互換用エイリアス =====
// ===================================================================

// 既存クライアントが saveWeek / saveAll を叩く可能性に備える
function saveWeekData(data) {
  // 旧 saveWeek = 下段水質測定
  return saveWaterMeasureData(data);
}
function saveAllData(data) {
  return saveWaterMeasureData({month:data.month, week:data.week, days:data.days});
}
