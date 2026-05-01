# 筒賀水質管理センター 週次点検アプリ

筒賀水質管理センター（広島県）の日常水質・主要機器運転時間・日常電気設備・日常機械設備を週単位で入力し、月次の Excel 記録表として出力するための PWA。

公開予定 URL: <https://cpkanri.github.io/tutuga/>

## 実装ステータス

- ✅ Phase 1（基盤、ブランディング、22機器運転時間タブ）
- ✅ Phase 2（日常水質2サブタブ、機械設備3サブタブ、電気設備3サブタブ、§4 複合転記）
- ✅ Phase 2 後追補
  - 機械設備の Excel 行マッピング（155項目、W4→W5 が 176行ピッチ）
  - 電気設備の Excel 行マッピング（67項目、サブタブごとに別の週オフセット sub1=+103/sub2=+101/sub3=+100）
  - 日常水質 + 水質管理報告 行マッピング検証（37項目、+52行ピッチ、平日動的計算）
- 🔜 動作テスト・GitHub Pages デプロイ

詳細は [`実装ステータス.md`](./実装ステータス.md) を参照。

## 概要

- スマートフォン（iOS / Android）のブラウザから利用、ホーム画面追加で PWA としてインストール可能
- 入力データは端末の localStorage（`tutuga_*` キー）と Google スプレッドシート（GAS バックエンド）の両方に保存
- 月末に「Excel 出力」ボタンで `筒賀水質管理センター記録表_YYYY-MM.xlsx` を生成
- 雛形: 加計浄化センター 週報アプリ（`/c/dev/kake-dev`）。タブ構成は筒賀仕様に合わせて再構成。

## タブ構成（筒賀仕様）

| メインタブ | サブタブ | 入力規模 |
|---|---|---|
| 日常水質 | 上段使用量 / 下段水質測定 | 11項目/週 + 34項目/日×5日 |
| 主要機器運転時間 | なし（1タブ） | 22機器×5日 |
| 日常機械設備 | 点検表1/3, 2/3, 3/3 | 約130項目/日×5日 |
| 日常電気設備 | 点検表1/3, 2/3, 3/3 | 約58項目/日×5日 |

## ファイル構成

```
.
├── index.html                              # 本体（HTML + CSS + JS）
├── manifest.json                           # PWA マニフェスト（scope: /tutuga/）
├── sw.js                                   # Service Worker (CACHE_NAME: tutuga-v1)
├── template.js                             # Excel テンプレート（base64）
├── icon-192.png                            # PWA アイコン
├── icon-512.png                            # PWA アイコン
├── tutuga_icon.svg                         # 元アイコン（SVG）
├── tutuga_gscode.txt                       # GAS バックエンド ソース（Apps Script へペースト）
├── 筒賀水質管理センター_template.xlsx        # Excel テンプレート原本（template.js 再生成用）
├── 筒賀_テンプレート_v5.xlsx                 # オリジナル（参考保存）
├── 筒賀アプリ_ClaudeCode指示書.md            # v1 指示書（参考）
├── 筒賀アプリ_ClaudeCode指示書_v2.md         # v2 指示書（最新仕様の典拠）
└── 実装ステータス.md                          # 実装進捗・残タスク一覧
```

## 設定値

| 項目 | 値 |
|---|---|
| GAS URL | `https://script.google.com/macros/s/AKfycbyOvwXmt-GL1cwo5qRuUnl2AnxI4aG0tJd5exg9Z96kqaG9RQCETsJLmMEBtlqlSVAb/exec` |
| Spreadsheet ID（テンプレ） | `1XFPH90_XXvyGQfewQXZORcdvc-jH2jAzP_yxV2pcrpQ` |
| localStorage プレフィックス | `tutuga_` |
| Service Worker キャッシュ | `tutuga-v1` |
| GitHub Pages 公開予定 | `cpkanri.github.io/tutuga/` |

## 開発

### Excel テンプレート更新時

```bash
python -c "
import base64
with open('筒賀水質管理センター_template.xlsx','rb') as f:
    b64 = base64.b64encode(f.read()).decode('ascii')
with open('template.js','w', encoding='utf-8') as o:
    o.write('window.EXCEL_TEMPLATE_B64 = \"' + b64 + '\";\n')
"
```

### Service Worker のキャッシュ更新

`sw.js` の `CACHE_NAME`（例: `tutuga-v1` → `tutuga-v2`）をインクリメントすると、PWA を開き直したときに新しい資産がフェッチされる。

### GAS バックエンド更新

`tutuga_gscode.txt` を編集 → Apps Script エディタにペースト → 新しいバージョンとして再デプロイ → `index.html` の `GAS_URL` を新 URL に置換（デプロイ URL を変えない設定であれば不要）。

## H yamane 開発ルール

- VBA/GAS でシート参照する際は **シート名ではなく index 番号** を使う（エンコーディング問題回避）
  - tutuga_gscode.txt の `SHEET_IDX` 定数を参照
- VBA ファイルは Shift-JIS で保存
- ファイル削除・システム変更以外は確認不要で最後まで一気に完了させる
- 既存の加計・上殿アプリの完成コードを参照可能なら最大限流用する

## 関連リポジトリ

- 雛形: 加計浄化センター 週報アプリ（`/c/dev/kake-dev`）
- 上位ベース: 上殿浄化センター 週報アプリ（`/c/dev/kamitono-dev`）
