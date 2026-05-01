const fs = require("fs");
const html = fs.readFileSync("index.html", "utf8");
const gas = fs.readFileSync("tutuga_gscode.txt", "utf8");

function extractHtmlKeys(html, name) {
  const re = new RegExp("const " + name + "\\s*=\\s*\\[([\\s\\S]*?)\\n\\];", "m");
  const m = html.match(re);
  if (!m) return [];
  const body = m[1];
  const keyRe = /key:\s*"([^"]+)"/g;
  const keys = [];
  let r;
  while ((r = keyRe.exec(body))) keys.push(r[1]);
  return keys;
}

function extractGasMap(gas, name) {
  const re = new RegExp("var " + name + "\\s*=\\s*\\{([\\s\\S]*?)\\n\\};", "m");
  const m = gas.match(re);
  if (!m) return {};
  const body = m[1];
  const keyRe = /"([^"]+)":\s*(\d+)/g;
  const out = {};
  let r;
  while ((r = keyRe.exec(body))) out[r[1]] = parseInt(r[2], 10);
  return out;
}

let allOk = true;
const allMaps = {};
for (const i of [1, 2, 3]) {
  const hk = extractHtmlKeys(html, "MECH_SECTIONS_" + i);
  const gMap = extractGasMap(gas, "MECH" + i + "_ROW_W1");
  const gk = Object.keys(gMap);
  const hs = new Set(hk), gs = new Set(gk);
  const onlyH = hk.filter((k) => !gs.has(k));
  const onlyG = gk.filter((k) => !hs.has(k));
  console.log(`MECH${i}: HTML=${hk.length}, GAS=${gk.length}, only-in-HTML=${onlyH.length}, only-in-GAS=${onlyG.length}`);
  if (onlyH.length) { console.log("  HTML-only:", onlyH); allOk = false; }
  if (onlyG.length) { console.log("  GAS-only:", onlyG); allOk = false; }
  const hdup = hk.filter((k, idx) => hk.indexOf(k) !== idx);
  if (hdup.length) { console.log("  HTML-dup:", hdup); allOk = false; }
  allMaps[i] = gMap;
}
console.log("\nKey parity: " + (allOk ? "OK" : "MISMATCH"));

const WEEK_OFFSETS_MECH = [0, 177, 354, 531, 707];
console.log("\n=== Sample row computations (subId, key, W1..W5 absolute rows) ===");
const samples = [
  [1, "ポンプ井__common1__内部状況確認"],
  [1, "返送汚泥ポンプ__common2__常用機選択"],
  [2, "余剰汚泥ポンプ__1__電流値"],
  [2, "空気圧縮機__common1__安全弁等の動作確認"],
  [3, "井水揚水ポンプ__common1__異音、振動等"],
  [3, "自動巻取型エアフィルター__common1__設定"]
];
for (const [sub, key] of samples) {
  const baseRow = allMaps[sub][key];
  const rows = WEEK_OFFSETS_MECH.map((off) => baseRow + off);
  console.log(`  sub${sub} [${key}] base=${baseRow} -> rows: ${rows.join(", ")}`);
}

console.log("\n=== W5 pitch sanity ===");
console.log("  W4 - W3 =", WEEK_OFFSETS_MECH[3] - WEEK_OFFSETS_MECH[2], "(expected 177)");
console.log("  W5 - W4 =", WEEK_OFFSETS_MECH[4] - WEEK_OFFSETS_MECH[3], "(expected 176)");

const counts = [56, 50, 49];
const tot = counts.reduce((a,b)=>a+b,0);
console.log(`\nTotal mapping items: ${tot} (expected 155)`);
