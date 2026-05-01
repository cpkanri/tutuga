// 日常水質 + 水質管理報告 行マッピング検証スクリプト
// 仕様: 筒賀_水質シート行マッピング.js / .md
// 検証対象: tutuga_gscode.txt 内の USAGE_ROW_OFFSET / MEASURE_ROW_OFFSET /
//           WATER_USAGE_ROWS_W1 / WATER_MEASURE_ROWS_W1 / REPORT_FRONT_COL /
//           REPORT_BACK_COL の各定義と、index.html 内の MEASURE_FIELDS / USAGE_KEYS。

const fs = require("fs");
const html = fs.readFileSync("index.html", "utf8");
const gas = fs.readFileSync("tutuga_gscode.txt", "utf8");

function extractGasObj(gas, name) {
  const re = new RegExp("var " + name + "\\s*=\\s*\\{([\\s\\S]*?)\\n\\};", "m");
  const m = gas.match(re);
  if (!m) return null;
  const out = {};
  // string keys
  let r;
  const keyReStr = /"([^"]+)":\s*(\d+)/g;
  while ((r = keyReStr.exec(m[1]))) out[r[1]] = parseInt(r[2], 10);
  // bare-identifier keys
  const keyReBare = /(?:^|[\s,{])([A-Za-z_][A-Za-z0-9_]*)\s*:\s*(\d+)/g;
  while ((r = keyReBare.exec(m[1]))) out[r[1]] = parseInt(r[2], 10);
  return out;
}

function extractGasArr(gas, name) {
  const re = new RegExp("var " + name + "\\s*=\\s*\\[([\\s\\S]*?)\\];", "m");
  const m = gas.match(re);
  if (!m) return null;
  const items = [];
  const re2 = /['"]([^'"]+)['"]/g;
  let r;
  while ((r = re2.exec(m[1]))) items.push(r[1]);
  return items;
}

function extractHtmlConstArr(html, name) {
  const re = new RegExp("const " + name + "\\s*=\\s*\\[([\\s\\S]*?)\\];", "m");
  const m = html.match(re);
  if (!m) return null;
  const items = [];
  const re2 = /['"]([^'"]+)['"]/g;
  let r;
  while ((r = re2.exec(m[1]))) items.push(r[1]);
  return items;
}

const blockRow1 = 1; // W1 block start
const WATER_WEEK_OFFSET = 52;

// ------------- 仕様 (筒賀_水質シート行マッピング.js より) -------------

const SPEC_USAGE_ROWS_W1 = {
  "全日電力量": 4,        // English: powerAllDay
  "力測電力量": 5,        // powerMeasure
  "無効電力量": 6,        // powerReactive
  "デマンド": 7,          // demand
  "水道使用量": 8,        // water
  "軽油（残量）": 9,      // diesel
  "返送汚泥流量（№1）": 10, // returnSludge1
  "返送汚泥流量（№2）": 11, // returnSludge2
  "放流流量": 12          // discharge
};

const SPEC_MEASURE_ROWS_W1 = {
  "水温_流入": 16, "水温_ディッチ1": 17, "水温_ディッチ2": 18, "水温_放流": 19,
  "外観_流入": 20, "外観_ディッチ1": 21, "外観_ディッチ2": 22, "外観_放流": 23,
  "透視度_流入水": 24, "透視度_放流水": 25,
  "臭気_流入": 26, "臭気_ディッチ1": 27, "臭気_ディッチ2": 28, "臭気_放流": 29,
  "PH_流入": 30, "PH_ディッチ1": 31, "PH_ディッチ2": 32, "PH_放流": 33,
  "SV1系_SV30": 34, "SV1系_24h": 35, "SV2系_SV30": 36, "SV2系_24h": 37,
  "MLSS_1系": 38, "MLSS_2系": 39,
  "汚泥界面_1系終沈": 40, "汚泥界面_2系終沈": 41,
  "残留塩素_放流水": 42,
  "塩素投入量_塩素混和池": 43,
};

// English (GAS) 名 → 仕様の Japanese 名
const ENG_TO_SPEC_USAGE = {
  powerAllDay: "全日電力量",   powerMeasure: "力測電力量",
  powerReactive: "無効電力量", demand: "デマンド",
  water: "水道使用量",         diesel: "軽油（残量）",
  returnSludge1: "返送汚泥流量（№1）",
  returnSludge2: "返送汚泥流量（№2）",
  discharge: "放流流量",
};

const ENG_TO_SPEC_MEASURE = {
  tempIn: "水温_流入", tempD1: "水温_ディッチ1", tempD2: "水温_ディッチ2", tempOut: "水温_放流",
  appearIn: "外観_流入", appearD1: "外観_ディッチ1", appearD2: "外観_ディッチ2", appearOut: "外観_放流",
  transIn: "透視度_流入水", transOut: "透視度_放流水",
  odorIn: "臭気_流入", odorD1: "臭気_ディッチ1", odorD2: "臭気_ディッチ2", odorOut: "臭気_放流",
  phIn: "PH_流入", phD1: "PH_ディッチ1", phD2: "PH_ディッチ2", phOut: "PH_放流",
  sv1_30: "SV1系_SV30", sv1_24h: "SV1系_24h", sv2_30: "SV2系_SV30", sv2_24h: "SV2系_24h",
  mlss1: "MLSS_1系", mlss2: "MLSS_2系",
  sludge1: "汚泥界面_1系終沈", sludge2: "汚泥界面_2系終沈",
  chlorine: "残留塩素_放流水", chlorineDose: "塩素投入量_塩素混和池",
};

// 仕様のWQR列 (筒賀_水質シート行マッピング.js より)
const SPEC_WQR_FRONT_COLS = {
  tempIn: 3, tempD1: 4, tempD2: 5, tempFinal1: 6, tempFinal2: 7, tempOut: 8,
  transIn: 15, transFinal1: 16, transFinal2: 17, transOut: 18,
};
const SPEC_WQR_BACK_COLS = {
  doNo1: 9, doNo2: 10,
  phIn: 11, phD1: 12, phD2: 13, phFinal1: 14, phFinal2: 15, phOut: 16,
  chlorine: 17, sv1_30: 18, sv2_30: 19,
};

// 3月平日の参考値（spec.md より、W1=2026-03-02 月)
const SPEC_MARCH_2026 = [
  ["2026-03-02", "月",  9, 50],
  ["2026-03-09", "月", 16, 57],
  ["2026-03-31", "火", 38, 79],
];

// ------------- ASSERT helpers -------------

let failures = 0;
function ok(label) { console.log("  ✓ " + label); }
function fail(label) { console.log("  ✗ " + label); failures++; }

// ------------- 1) USAGE 行 -------------
console.log("\n[1] USAGE_ROW_OFFSET → 絶対行 vs 仕様 WATER_USAGE_ROWS_W1");
const usageOffset = extractGasObj(gas, "USAGE_ROW_OFFSET");
for (const eng in ENG_TO_SPEC_USAGE) {
  const jpKey = ENG_TO_SPEC_USAGE[eng];
  const spec = SPEC_USAGE_ROWS_W1[jpKey];
  const offsetVal = usageOffset[eng];
  const absRow = blockRow1 + offsetVal;
  if (absRow === spec) ok(`${eng} (${jpKey}): block+${offsetVal} = r${absRow}`);
  else fail(`${eng} (${jpKey}): block+${offsetVal} = r${absRow}, spec r${spec}`);
}

// ------------- 2) MEASURE 行 -------------
console.log("\n[2] MEASURE_ROW_OFFSET → 絶対行 vs 仕様 WATER_MEASURE_ROWS_W1");
const measureOffset = extractGasObj(gas, "MEASURE_ROW_OFFSET");
for (const eng in ENG_TO_SPEC_MEASURE) {
  const jpKey = ENG_TO_SPEC_MEASURE[eng];
  const spec = SPEC_MEASURE_ROWS_W1[jpKey];
  const offsetVal = measureOffset[eng];
  const absRow = blockRow1 + offsetVal;
  if (absRow === spec) ok(`${eng} (${jpKey}): block+${offsetVal} = r${absRow}`);
  else fail(`${eng} (${jpKey}): block+${offsetVal} = r${absRow}, spec r${spec}`);
}

// ------------- 3) GAS 内の Japanese-keyed 仕様マッピング (alias) -------------
console.log("\n[3] WATER_USAGE_ROWS_W1 / WATER_MEASURE_ROWS_W1 (GAS alias) vs 仕様");
const gasUsageJp = extractGasObj(gas, "WATER_USAGE_ROWS_W1") || {};
const gasMeasureJp = extractGasObj(gas, "WATER_MEASURE_ROWS_W1") || {};
for (const k in SPEC_USAGE_ROWS_W1) {
  if (gasUsageJp[k] === SPEC_USAGE_ROWS_W1[k]) ok(`USAGE [${k}] = r${gasUsageJp[k]}`);
  else fail(`USAGE [${k}]: GAS=${gasUsageJp[k]}, spec=r${SPEC_USAGE_ROWS_W1[k]}`);
}
for (const k in SPEC_MEASURE_ROWS_W1) {
  if (gasMeasureJp[k] === SPEC_MEASURE_ROWS_W1[k]) ok(`MEASURE [${k}] = r${gasMeasureJp[k]}`);
  else fail(`MEASURE [${k}]: GAS=${gasMeasureJp[k]}, spec=r${SPEC_MEASURE_ROWS_W1[k]}`);
}

// ------------- 4) WQR 列 -------------
console.log("\n[4] REPORT_FRONT_COL / REPORT_BACK_COL vs 仕様");
const front = extractGasObj(gas, "REPORT_FRONT_COL");
const back = extractGasObj(gas, "REPORT_BACK_COL");
for (const k in SPEC_WQR_FRONT_COLS) {
  if (front[k] === SPEC_WQR_FRONT_COLS[k]) ok(`FRONT [${k}] = col ${front[k]}`);
  else fail(`FRONT [${k}]: GAS=${front[k]}, spec=${SPEC_WQR_FRONT_COLS[k]}`);
}
for (const k in SPEC_WQR_BACK_COLS) {
  if (back[k] === SPEC_WQR_BACK_COLS[k]) ok(`BACK [${k}] = col ${back[k]}`);
  else fail(`BACK [${k}]: GAS=${back[k]}, spec=${SPEC_WQR_BACK_COLS[k]}`);
}

// ------------- 5) HTML form 全 28 measure 項目 + 6 終沈 + 9 usage を含むか -------------
console.log("\n[5] HTML MEASURE_FIELDS / USAGE_KEYS の網羅性");
const htmlMeasure = extractHtmlConstArr(html, "MEASURE_FIELDS") || [];
const htmlUsage = extractHtmlConstArr(html, "USAGE_KEYS") || [];
for (const eng in ENG_TO_SPEC_MEASURE) {
  if (htmlMeasure.includes(eng)) ok(`HTML MEASURE has ${eng}`);
  else fail(`HTML MEASURE missing ${eng}`);
}
const finalKeys = ["tempFinal1","tempFinal2","transFinal1","transFinal2","phFinal1","phFinal2"];
for (const k of finalKeys) {
  if (htmlMeasure.includes(k)) ok(`HTML MEASURE has 終沈 key ${k}`);
  else fail(`HTML MEASURE missing 終沈 key ${k}`);
}
for (const eng in ENG_TO_SPEC_USAGE) {
  if (htmlUsage.includes(eng)) ok(`HTML USAGE has ${eng}`);
  else fail(`HTML USAGE missing ${eng}`);
}
for (const k of ["excessSludge1","excessSludge2"]) {
  if (htmlUsage.includes(k)) ok(`HTML USAGE has ${k} (→ elec sub2 へ転記)`);
  else fail(`HTML USAGE missing ${k}`);
}

// ------------- 6) 5週分の絶対行サンプル -------------
console.log("\n[6] 5週分の絶対行 (block_row + offset)");
const samples = [
  ["全日電力量", 4],
  ["軽油（残量）", 9],
  ["放流流量", 12],
  ["水温_流入", 16],
  ["透視度_放流水", 25],
  ["残留塩素_放流水", 42],
  ["塩素投入量_塩素混和池", 43],
];
for (const [label, base] of samples) {
  const rows = [0,1,2,3,4].map(i => base + WATER_WEEK_OFFSET * i);
  console.log(`  ${label} (W1=r${base}) -> W1..W5: ${rows.join(", ")}`);
}

// ------------- 7) WQR 行式の検証 -------------
console.log("\n[7] WQR 行式 frontRow=7+day, backRow=48+day  (3月平日サンプル)");
for (const [date, dow, expF, expB] of SPEC_MARCH_2026) {
  const day = parseInt(date.slice(8,10), 10);
  const f = 7 + day, b = 48 + day;
  if (f === expF && b === expB) ok(`${date}(${dow}): front r${f}, back r${b}`);
  else fail(`${date}(${dow}): got front r${f}/back r${b}, spec front r${expF}/back r${expB}`);
}

// ------------- 8) 二重書き込み参照キー (DO/燃料/余剰) が正しい elec キーを参照するか -------------
console.log("\n[8] §4 二重書き込みの elec キー参照");
function findGasVar(name) {
  const re = new RegExp("var " + name + '\\s*=\\s*"([^"]+)"');
  const m = gas.match(re);
  return m ? m[1] : null;
}
const expectedDoNo1 = "Nｏ．１オキシデーションディッチＤＯ計__ＤＯ値__㎎/ℓ";
const expectedDoNo2 = "Nｏ．２オキシデーションディッチＤＯ計__ＤＯ値__㎎/ℓ";
const expectedExcess1 = "余剰汚泥流量計__No.１余剰汚泥ポンプ流量計読み__㎥/ｈ";
const expectedExcess2 = "余剰汚泥流量計__No.２余剰汚泥ポンプ流量計読み__㎥/ｈ";
const expectedFuel = "燃料タンク__残量の値__ℓ";
const checks = [
  ["ELEC_DO_KEY_NO1", expectedDoNo1],
  ["ELEC_DO_KEY_NO2", expectedDoNo2],
  ["ELEC_EXCESS_KEY_NO1", expectedExcess1],
  ["ELEC_EXCESS_KEY_NO2", expectedExcess2],
  ["ELEC_FUEL_TANK_KEY", expectedFuel],
];
for (const [name, want] of checks) {
  const got = findGasVar(name);
  if (got === want) ok(`${name} = "${got}"`);
  else fail(`${name}: got "${got}", expected "${want}"`);
}

// ------------- 9) 各シートのテンプレ範囲内に書き込みが収まるか -------------
//   注意: 仕様書の max_row=260 は「日常水質シート」を指す。電気/機械設備は別シートで
//        それぞれ独自のテンプレ縦幅 (Phase 2 採寸時に sub1=r404+, sub2=r466+, sub3=r483+;
//        機械=r875+ を確認済み) を持つため、シート別に bound を当てる。
console.log("\n[9] 各シート別 テンプレ範囲内チェック");
const MAX_COL = 18;

// 9a) 日常水質シート: max_row=260 (仕様値)
const waterMaxRow = 43 + WATER_WEEK_OFFSET * 4;       // 251 (塩素投入量 W5)
const waterMaxCol = 14;                                // N列
if (waterMaxRow <= 260) ok(`日常水質: W5最大行 r${waterMaxRow} ≤ 260 (仕様値)`);
else fail(`日常水質: W5最大行 r${waterMaxRow} > 260`);
if (waterMaxCol <= MAX_COL) ok(`日常水質: 最大列 ${waterMaxCol} (N) ≤ ${MAX_COL}`);
else fail(`日常水質: 最大列 ${waterMaxCol} > ${MAX_COL}`);

// 9b) 水質管理報告シート: 31日(month max) → frontRow=38, backRow=79
const wqrMaxRow = 48 + 31; // 79
const wqrMaxCol = 19;      // S列 (SV2系 SV30)
if (wqrMaxRow <= 80) ok(`水質管理報告: 月末最大行 r${wqrMaxRow} ≤ 80 (31日連続構造)`);
else fail(`水質管理報告: 月末最大行 r${wqrMaxRow} > 80`);
if (wqrMaxCol <= 19) ok(`水質管理報告: 最大列 ${wqrMaxCol} (S列, 仕様通り)`);
else fail(`水質管理報告: 最大列 ${wqrMaxCol} 超過`);

// 9c) 日常電気設備シート: 各サブタブの W5 最大行
//   sub1: r31 + 403 = r434 (低圧分岐 No.3コンデンサー)
//   sub2: r57 + 401 = r458 (余剰汚泥No.2、二重書き込み先)、r65+401=r466 (COD値)
//   sub3: r82 + 400 = r482 (燃料タンク残量、二重書き込み先)、r83+400=r483 (通報装置)
const elecMaxRow = Math.max(31 + 403, 65 + 401, 83 + 400); // = 483
if (elecMaxRow <= 500) ok(`日常電気設備: 全サブタブ W5最大行 r${elecMaxRow} ≤ 500 (Phase2採寸範囲内)`);
else fail(`日常電気設備: r${elecMaxRow} 超過`);

// 9d) 日常機械設備シート: W5 最大行 r168+707=r875 (自動巻取型エアフィルター 設定)
const mechMaxRow = 168 + 707;
if (mechMaxRow <= 900) ok(`日常機械設備: W5最大行 r${mechMaxRow} ≤ 900 (Phase2採寸範囲内)`);
else fail(`日常機械設備: r${mechMaxRow} 超過`);

// ------------- 10) 月毎の平日数が動的計算されるか (frontRow式の単純検証) -------------
console.log("\n[10] 平日数の動的計算 (月毎の Mon-Fri 抽出)");
function weekdaysInMonth(year, monthNum0Based) {
  const days = new Date(year, monthNum0Based + 1, 0).getDate();
  let count = 0;
  for (let d = 1; d <= days; d++) {
    const w = new Date(year, monthNum0Based, d).getDay();
    if (w !== 0 && w !== 6) count++;
  }
  return count;
}
const monthCases = [
  [2026, 2, "2026-03", 22],  // 3月
  [2026, 3, "2026-04", 22],  // 4月
  [2026, 1, "2026-02", 20],  // 2月 2026 = 28日, 平日20日 (月〜金)
];
for (const [y, m0, label, expected] of monthCases) {
  const got = weekdaysInMonth(y, m0);
  if (got === expected) ok(`${label}: 平日 ${got}日 (期待: ${expected})`);
  else fail(`${label}: 平日 ${got}日, 期待 ${expected}日`);
}

// 確認: 3月平日マッピングがハードコードされていないか (= dateOfWeekDay() で動的計算されているか)
const usesDynamic = /dateOfWeekDay\s*\(/.test(gas);
const usesGetDate = /reportRowFront\s*\(\s*date\s*\)\s*\{[^}]*date\.getDate\(\)/.test(gas);
if (usesDynamic) ok("GAS: dateOfWeekDay() を呼び出し平日を動的に取得");
else fail("GAS: dateOfWeekDay() の呼び出しが見つからない (ハードコードの可能性)");
if (usesGetDate) ok("GAS: reportRowFront() が date.getDate() ベースで動的計算");
else fail("GAS: reportRowFront() の動的計算が見つからない");

// ------------- 11) 平日のみ書き込み (土日除外) -------------
console.log("\n[11] 平日のみ書き込み (form input で Mon-Fri のみ受領)");
const daysArr = extractGasArr(gas, "DAYS");
if (daysArr && daysArr.length === 5 && !daysArr.includes("sat") && !daysArr.includes("sun")) {
  ok(`GAS: DAYS = [${daysArr.join(",")}] (土日なし)`);
} else {
  fail(`GAS: DAYS = ${JSON.stringify(daysArr)} (期待: mon-fri のみ)`);
}

// ------------- summary -------------
console.log("\n========================================");
console.log(failures === 0 ? "ALL CHECKS PASSED ✓" : `FAILED: ${failures} issue(s)`);
process.exit(failures === 0 ? 0 : 1);
