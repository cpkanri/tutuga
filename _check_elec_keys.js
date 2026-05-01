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
  const hk = extractHtmlKeys(html, "ELEC_SECTIONS_" + i);
  const gMap = extractGasMap(gas, "ELEC" + i + "_ROW_W1");
  const gk = Object.keys(gMap);
  const hs = new Set(hk), gs = new Set(gk);
  const onlyH = hk.filter((k) => !gs.has(k));
  const onlyG = gk.filter((k) => !hs.has(k));
  console.log(`ELEC${i}: HTML=${hk.length}, GAS=${gk.length}, only-in-HTML=${onlyH.length}, only-in-GAS=${onlyG.length}`);
  if (onlyH.length) { console.log("  HTML-only:", onlyH); allOk = false; }
  if (onlyG.length) { console.log("  GAS-only:", onlyG); allOk = false; }
  const hdup = hk.filter((k, idx) => hk.indexOf(k) !== idx);
  if (hdup.length) { console.log("  HTML-dup:", hdup); allOk = false; }
  allMaps[i] = gMap;
}
console.log("\nKey parity: " + (allOk ? "OK" : "MISMATCH"));

const WEEK_OFFSETS_ELEC = {
  sub1: [0, 103, 203, 303, 403],
  sub2: [0, 101, 201, 301, 401],
  sub3: [0, 100, 200, 300, 400],
};

console.log("\n=== Sample row computations (sub, key, W1..W5 absolute rows) ===");
const samples = [
  [1, "sub1", "引込・受電__電圧__KV"],
  [1, "sub1", "低圧分岐__No.3コンデンサー__切/入"],
  [2, "sub2", "Nｏ．１オキシデーションディッチＤＯ計__ＤＯ値__㎎/ℓ"],
  [2, "sub2", "余剰汚泥流量計__No.１余剰汚泥ポンプ流量計読み__㎥/ｈ"],
  [2, "sub2", "余剰汚泥流量計__No.２余剰汚泥ポンプ流量計読み__㎥/ｈ"],
  [2, "sub2", "（放流水ＵＶ）__ＣＯＤ値__㎎/ℓ"],
  [3, "sub3", "全窒素・全燐計__外観の確認"],
  [3, "sub3", "燃料タンク__残量の値__ℓ"],
  [3, "sub3", "通報装置__通報状態の確認"],
];
for (const [pageIdx, sub, key] of samples) {
  const baseRow = allMaps[pageIdx][key];
  const rows = WEEK_OFFSETS_ELEC[sub].map((off) => baseRow + off);
  console.log(`  ${sub} [${key}] base=${baseRow} -> rows: ${rows.join(", ")}`);
}

console.log("\n=== Per-subtab pitch sanity ===");
for (const sub of ["sub1", "sub2", "sub3"]) {
  const o = WEEK_OFFSETS_ELEC[sub];
  const pitches = [o[1] - o[0], o[2] - o[1], o[3] - o[2], o[4] - o[3]];
  console.log(`  ${sub}: W1->W2=${pitches[0]}, W2->W3=${pitches[1]}, W3->W4=${pitches[2]}, W4->W5=${pitches[3]}`);
}

const counts = [29, 26, 12];
const tot = counts.reduce((a,b)=>a+b,0);
console.log(`\nTotal mapping items: ${tot} (expected 67)`);
