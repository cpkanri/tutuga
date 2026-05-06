// scripts/regenerate_template_base64.js
//
// _handover/筒賀_テンプレート.xlsx を読んで base64 にし、
// template.js 1 行目の `var EXCEL_TEMPLATE_B64 = "..."` を上書きする。
//
// 使い方: node scripts/regenerate_template_base64.js
//
// Phase 8 修正08 で初期投入。テンプレ xlsx は .gitignore で除外されており、
// アプリは template.js に埋め込まれた base64 のみを参照するため、
// テンプレ更新時は本スクリプトで再生成して template.js を差し替える運用。

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const XLSX_PATH = path.join(ROOT, '_handover', '筒賀_テンプレート.xlsx');
const TEMPLATE_JS = path.join(ROOT, 'template.js');

if (!fs.existsSync(XLSX_PATH)) {
  console.error('ERROR: テンプレ xlsx が見つかりません: ' + XLSX_PATH);
  process.exit(1);
}
if (!fs.existsSync(TEMPLATE_JS)) {
  console.error('ERROR: template.js が見つかりません: ' + TEMPLATE_JS);
  process.exit(1);
}

const xlsxBytes = fs.readFileSync(XLSX_PATH);
const newB64 = xlsxBytes.toString('base64');

const src = fs.readFileSync(TEMPLATE_JS, 'utf8');

// 先頭の `var EXCEL_TEMPLATE_B64 = "..."` を新 base64 で置換。
// 1 行目に集約されている前提（734KB 級の長文字列）。
const re = /^var\s+EXCEL_TEMPLATE_B64\s*=\s*"[^"]*";/;
if (!re.test(src)) {
  console.error('ERROR: template.js 先頭に EXCEL_TEMPLATE_B64 宣言が見つかりません');
  process.exit(1);
}
const newSrc = src.replace(re, 'var EXCEL_TEMPLATE_B64 = "' + newB64 + '";');

fs.writeFileSync(TEMPLATE_JS, newSrc, 'utf8');

console.log('Updated template.js');
console.log('  xlsx bytes:    ' + xlsxBytes.length);
console.log('  base64 chars:  ' + newB64.length);
