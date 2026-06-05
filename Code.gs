// =====================================================================
// 筒賀水質管理センター 週報アプリ GAS バックエンド (Phase 3)
// 作成日: 2026-05-04
//
// 旧 tutuga_gscode.gs (_handover/) のロジックを移植し、ExcelJS 方式に
// 対応するため Excel 出力部 (exportExcel) を削除。クライアント (index.html)
// は GAS から JSON を取得し、ExcelJS で xlsx を生成する方針。
//
// 8 シート構成:
//   日常水質_上段       per-week JSON      [月, 週, データJSON, 更新日時]
//   日常水質_下段       per-day  JSON      [月, 週, 曜日, データJSON, 更新日時]
//   日常電気設備        per-day  JSON      [月, 週, 曜日, データJSON, 更新日時]
//   日常機械設備        per-day  JSON      [月, 週, 曜日, データJSON, 更新日時]
//   機器運転時間        per-week JSON      [月, 週, データJSON, 更新日時]
//   祝日リスト          [日付, 名称]
//   設定                [key, value]
//   ユーザー管理        [ID, パスワードハッシュ, 氏名, 有効, 最終ログイン, 備考]
//
// データ列スキーマは Sheets セル容量上限 (50,000 文字) に収まるよう
// month / week / day を独立したキー列に分け、JSON は data 列のみに格納する。
// =====================================================================

// ===== 定数: Spreadsheet ID / シート名 =====
var SS_ID = '1XFPH90_XXvyGQfewQXZORcdvc-jH2jAzP_yxV2pcrpQ';

var SHEET_WATER_USAGE   = '日常水質_上段';
var SHEET_WATER_MEASURE = '日常水質_下段';
var SHEET_ELECTRICAL    = '日常電気設備';
var SHEET_MECHANICAL    = '日常機械設備';
var SHEET_EQUIPMENT     = '機器運転時間';
var SHEET_HOLIDAYS      = '祝日リスト';
var SHEET_SETTINGS      = '設定';
var SHEET_USERS         = 'ユーザー管理';

// ===== 列ヘッダ定義 =====
var HEADERS_WATER_USAGE   = ['月', '週', 'データJSON', '更新日時'];
var HEADERS_WATER_MEASURE = ['月', '週', '曜日', 'データJSON', '更新日時'];
var HEADERS_ELECTRICAL    = ['月', '週', '曜日', 'データJSON', '更新日時'];
var HEADERS_MECHANICAL    = ['月', '週', '曜日', 'データJSON', '更新日時'];
var HEADERS_EQUIPMENT     = ['月', '週', 'データJSON', '更新日時'];

// ===== 認証: 単一ユーザ (cpkanri / cptutuga) =====
// Phase 3 時点ではハードコード。Phase 5+ で ユーザー管理シートに移行する場合あり。
var TUTUGA_USERS = [
  { id: 'cpkanri', password: 'cptutuga', name: '筒賀' }
];
var PUBLIC_ACTIONS = { ping: 1, login: 1, verifyToken: 1 };

// ===== 曜日マッピング =====
var DAYS         = ['mon', 'tue', 'wed', 'thu', 'fri'];
var DAY_NAMES    = { mon: '月', tue: '火', wed: '水', thu: '木', fri: '金' };
var DAY_REVERSE  = { '月': 'mon', '火': 'tue', '水': 'wed', '木': 'thu', '金': 'fri' };
var DAY_OFFSETS  = { mon: 0, tue: 1, wed: 2, thu: 3, fri: 4 };

// ===== excessSludge1/2・autoFrom 転記キー (電気 sub2/sub3) =====
// 上段使用量フォームの該当キーを、保存時に日常電気設備の per-day JSON に転記する。
// (excessSludge1/2 は日常水質_上段 JSON には含めない)
var T_ELEC_EXCESS_KEY_NO1     = '余剰汚泥流量計__No.１余剰汚泥ポンプ流量計読み__㎥/ｈ';   // sub2 r56
var T_ELEC_EXCESS_KEY_NO2     = '余剰汚泥流量計__No.２余剰汚泥ポンプ流量計読み__㎥/ｈ';   // sub2 r57
var T_ELEC_RETURN_SLUDGE_NO1  = 'Nｏ．１返送汚泥流量計__流量計積算計読み__㎥';            // sub2 r49
var T_ELEC_RETURN_SLUDGE_NO2  = 'Nｏ．２返送汚泥流量計__流量計積算計読み__㎥';            // sub2 r53
var T_ELEC_DISCHARGE          = '放流流量計__流量計積算の読み__㎥';                       // sub2 r61
var T_ELEC_FUEL_TANK          = '燃料タンク__残量の値__ℓ';                                // sub3 r82


// =====================================================================
// ユーティリティ
// =====================================================================

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 入力値を 'YYYY-MM' 形式に正規化する。
// CSV インポート時の落とし穴 ('YYYY-MM-DD' / 先頭アポストロフィ / Date オブジェクト)
// を全てカバーする。判定は instanceof ではなく duck typing (getFullYear) で行う。
function normalizeMonth(v) {
  if (v === null || v === undefined || v === '') return '';
  // Date-like (instanceof Date は GAS で誤判定するので getFullYear 関数の有無で判定)
  if (typeof v === 'object' && typeof v.getFullYear === 'function') {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2);
  }
  var s = String(v).trim();
  if (s.length === 0) return '';
  // 先頭アポストロフィ (Sheets で文字列化されたセル) を剥がす
  if (s.charAt(0) === "'") s = s.substring(1).trim();
  // 'YYYY-MM-DD' / 'YYYY/MM/DD' → 先頭 7 文字 (CSV 落とし穴対策の本体)
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  // 'YYYY-MM' / 'YYYY/MM'
  m = s.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  // 'YYYY年M月'
  m = s.match(/^(\d{4})年(\d{1,2})月$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  // 'M月' (年は当年)
  m = s.match(/^(\d{1,2})月$/);
  if (m) {
    var y = new Date().getFullYear();
    return y + '-' + ('0' + m[1]).slice(-2);
  }
  return '';
}

function toDateStr(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'object' && typeof v.getFullYear === 'function') {
    return v.getFullYear() + '-' +
           ('0' + (v.getMonth() + 1)).slice(-2) + '-' +
           ('0' + v.getDate()).slice(-2);
  }
  return String(v);
}

function matchMonth(cellValue, month) {
  if (cellValue === month) return true;
  return normalizeMonth(cellValue) === month;
}

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8eef6');
    }
  }
  return sh;
}

// 連続する行を 1 回の setValues にまとめて batched 書込する (N+1 根絶用)。
// shiwagi Phase 25-c-γ / kamitono 25-c-β の flushRowUpdates と同手法。
// updates = [{ r: 絶対行番号(1始まり), v: [列値...] }]。startCol は書込開始列(1始まり)。
// 非連続行・幅違いは別レンジに分割し、gap/未対象行には一切触れない。
function flushRowUpdates(sheet, updates, startCol) {
  if (!updates || updates.length === 0) return;
  updates.sort(function(a, b) { return a.r - b.r; });
  var i = 0;
  while (i < updates.length) {
    var runStart = i;
    var width = updates[i].v.length;
    while (i + 1 < updates.length &&
           updates[i + 1].r === updates[i].r + 1 &&
           updates[i + 1].v.length === width) {
      i++;
    }
    var block = [];
    for (var k = runStart; k <= i; k++) block.push(updates[k].v);
    sheet.getRange(updates[runStart].r, startCol, block.length, width).setValues(block);
    i++;
  }
}

// [Phase A/B] 旧 findRowByMonthWeek / findRowByMonthWeekDay は N+1 の元凶 (毎行
// フルシート読み) のため全廃。行特定は各保存ハンドラ内で getValues 1 回 +
// 索引マップ ('week' / 'week|曜日' → 絶対行) に置換済み。


// =====================================================================
// 認証 (TUTUGA_USERS をハードコード参照、トークンは ScriptProperties に保管)
// =====================================================================

function login(data) {
  var userId   = data && data.userId;
  var password = data && data.password;
  if (!userId || !password) {
    return jsonResponse({ error: 'IDとパスワードを入力してください' });
  }
  for (var i = 0; i < TUTUGA_USERS.length; i++) {
    var u = TUTUGA_USERS[i];
    if (u.id !== userId) continue;
    if (u.password !== password) {
      return jsonResponse({ error: 'IDまたはパスワードが正しくありません' });
    }
    var t = issueToken(u.id);
    return jsonResponse({
      status: 'ok', token: t.token, expires: t.expires,
      userId: u.id, name: u.name
    });
  }
  return jsonResponse({ error: 'IDまたはパスワードが正しくありません' });
}

function issueToken(userId) {
  var props = PropertiesService.getScriptProperties();
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var now = new Date().getTime();
  var expires = now + 30 * 24 * 60 * 60 * 1000; // 30 日
  props.setProperty('TK_' + token, JSON.stringify({ userId: userId, expires: expires }));
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
    } catch (e) {
      props.deleteProperty(key);
    }
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
    // TUTUGA_USERS にユーザが存在することを再確認
    for (var i = 0; i < TUTUGA_USERS.length; i++) {
      if (TUTUGA_USERS[i].id === obj.userId) return obj.userId;
    }
    props.deleteProperty('TK_' + token);
    return null;
  } catch (e) {
    return null;
  }
}

function logout(data) {
  var token = data && data.token;
  if (token) PropertiesService.getScriptProperties().deleteProperty('TK_' + token);
  return jsonResponse({ status: 'ok' });
}

function verifyTokenApi(token) {
  var userId = verifyToken(token);
  if (userId) return jsonResponse({ valid: true, userId: userId });
  return jsonResponse({ valid: false });
}

function requireAuth(action, token) {
  if (PUBLIC_ACTIONS[action]) return null;
  var userId = verifyToken(token);
  if (!userId) return jsonResponse({ error: '認証が必要です', authRequired: true });
  return null;
}


// =====================================================================
// 祝日リスト
// =====================================================================

function getHolidays() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_HOLIDAYS);
    if (!sheet) return jsonResponse({ holidays: [] });
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return jsonResponse({ holidays: [] });
    var data = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var d = data[i][0];
      if (!d) continue;
      if (typeof d === 'object' && typeof d.getFullYear === 'function') {
        var y = d.getFullYear();
        var m = ('0' + (d.getMonth() + 1)).slice(-2);
        var day = ('0' + d.getDate()).slice(-2);
        result.push({ date: y + '-' + m + '-' + day, name: data[i][1] || '' });
      } else {
        var ds = toDateStr(d);
        if (ds) result.push({ date: ds, name: data[i][1] || '' });
      }
    }
    return jsonResponse({ holidays: result });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}


// =====================================================================
// HTTP エンドポイント
// =====================================================================

function doGet(e) {
  try {
    var a = e.parameter.action;
    var token = e.parameter.token || '';
    var authFail = requireAuth(a, token);
    if (authFail) return authFail;

    if (a === 'ping')         return jsonResponse({ status: 'ok', message: '筒賀週報API稼働中 (Phase3)' });
    if (a === 'verifyToken')  return verifyTokenApi(token);
    if (a === 'getAllData')   return getAllData(e.parameter.month);
    if (a === 'getHolidays')  return getHolidays();
    return jsonResponse({ error: '不明なアクション: ' + a });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var a = data.action;
    var token = data.token || '';
    var authFail = requireAuth(a, token);
    if (authFail) return authFail;

    if (a === 'login')           return login(data);
    if (a === 'logout')          return logout(data);
    if (a === 'saveWaterUsage')  return saveWaterUsage(data);
    if (a === 'saveWaterMeasure') return saveWaterMeasure(data);
    if (a === 'saveEquipment')   return saveEquipment(data);
    if (a === 'saveElectrical')  return saveElectrical(data);
    if (a === 'saveMechanical')  return saveMechanical(data);
    if (a === 'saveEverything')  return saveEverything(data);
    return jsonResponse({ error: '不明なアクション: ' + a });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}


// =====================================================================
// 保存関数 (5 種) + saveEverything オーケストレータ
// =====================================================================

// 上段使用量 (per-week, JSON)
//
// 仕様:
//   - 11 キー (powerAllDay … excessSludge2) のうち、excessSludge1/2 は
//     日常水質_上段 JSON には書き込まない (物理的に excessSludge は除外)
//   - excessSludge1/2 を含む 6 個の day-level 値を 日常電気設備 JSON に転記:
//       excessSludge1[day] → ELEC[day][T_ELEC_EXCESS_KEY_NO1]    (sub2 r56)
//       excessSludge2[day] → ELEC[day][T_ELEC_EXCESS_KEY_NO2]    (sub2 r57)
//       returnSludge1[day] → ELEC[day][T_ELEC_RETURN_SLUDGE_NO1] (sub2 r49)
//       returnSludge2[day] → ELEC[day][T_ELEC_RETURN_SLUDGE_NO2] (sub2 r53)
//       discharge[day]     → ELEC[day][T_ELEC_DISCHARGE]         (sub2 r61)
//       diesel[day]        → ELEC[day][T_ELEC_FUEL_TANK]         (sub3 r82)
//
//   - 入力フィールドが未入力 (null/undefined/'') の場合は転記しない
//
function saveWaterUsage(data) {
  var month = normalizeMonth(data.month);
  var weeks = data.weeks;
  if (!month || !weeks) return jsonResponse({ error: 'パラメータ不足' });

  var sh = getOrCreateSheet(SHEET_WATER_USAGE, HEADERS_WATER_USAGE);
  var elecSh = getOrCreateSheet(SHEET_ELECTRICAL, HEADERS_ELECTRICAL);
  var now = new Date();
  var savedWeeks = 0;
  var transcribedDays = 0;

  // --- Phase B: 索引マップを各シート 1 回ずつ構築 (当月のみ・先勝ち) ---
  // 水質上段: week → 絶対行
  var wuRowByWeek = {};
  var wuLr = sh.getLastRow();
  if (wuLr > 1) {
    var wuGrid = sh.getRange(2, 1, wuLr - 1, 2).getValues();
    for (var i = 0; i < wuGrid.length; i++) {
      if (!matchMonth(wuGrid[i][0], month)) continue;
      var gw = parseInt(wuGrid[i][1], 10); if (!gw) continue;
      if (wuRowByWeek[gw] === undefined) wuRowByWeek[gw] = i + 2;
    }
  }
  // 電気: 'week|曜日' → { r:絶対行, json:既存data(col4), dirty } (col4 を同時取得し getValue 全廃)
  var elecIdx = {};
  var elLr = elecSh.getLastRow();
  if (elLr > 1) {
    var elGrid = elecSh.getRange(2, 1, elLr - 1, 4).getValues();
    for (var j = 0; j < elGrid.length; j++) {
      if (!matchMonth(elGrid[j][0], month)) continue;
      var ew = parseInt(elGrid[j][1], 10); if (!ew) continue;
      var ekey0 = ew + '|' + String(elGrid[j][2]);
      if (elecIdx[ekey0] === undefined) {
        var ej = {};
        try { ej = JSON.parse(elGrid[j][3] || '{}'); } catch (e) { ej = {}; }
        elecIdx[ekey0] = { r: j + 2, w: ew, dayName: String(elGrid[j][2]), json: ej, dirty: false };
      }
    }
  }

  var wuUpdates = [], wuAppends = [];
  var elecNew = {}, elecNewOrder = [];   // 新規電気行 (key → {w, dayName, json}), 出現順

  for (var wn in weeks) {
    var weekData = weeks[wn]; if (!weekData) continue;
    var w = parseInt(wn, 10);
    if (!w) continue;

    // 1) excessSludge1/2 を分離して 日常水質_上段 から除外
    var sheetCopy = {};
    for (var k in weekData) {
      if (k === 'excessSludge1' || k === 'excessSludge2') continue;
      sheetCopy[k] = weekData[k];
    }
    var rv = [month, w, JSON.stringify(sheetCopy), now];
    if (wuRowByWeek[w] !== undefined) wuUpdates.push({ r: wuRowByWeek[w], v: rv });
    else wuAppends.push(rv);
    savedWeeks++;

    // 2) 日常電気設備 への自動転記 (per-day マージ, in-memory 累積)
    for (var di = 0; di < DAYS.length; di++) {
      var dayKey = DAYS[di];
      var transcribe = {};
      var picked = false;

      var es1 = pickDayValue(weekData.excessSludge1, dayKey);
      var es2 = pickDayValue(weekData.excessSludge2, dayKey);
      var rs1 = pickDayValue(weekData.returnSludge1, dayKey);
      var rs2 = pickDayValue(weekData.returnSludge2, dayKey);
      var dis = pickDayValue(weekData.discharge,     dayKey);
      var dsl = pickDayValue(weekData.diesel,        dayKey);

      if (es1 != null) { transcribe[T_ELEC_EXCESS_KEY_NO1]    = es1; picked = true; }
      if (es2 != null) { transcribe[T_ELEC_EXCESS_KEY_NO2]    = es2; picked = true; }
      if (rs1 != null) { transcribe[T_ELEC_RETURN_SLUDGE_NO1] = rs1; picked = true; }
      if (rs2 != null) { transcribe[T_ELEC_RETURN_SLUDGE_NO2] = rs2; picked = true; }
      if (dis != null) { transcribe[T_ELEC_DISCHARGE]         = dis; picked = true; }
      if (dsl != null) { transcribe[T_ELEC_FUEL_TANK]         = dsl; picked = true; }

      if (!picked) continue;

      // 旧 mergeIntoElectricalDay 相当: dayName 無しは書込スキップ (カウンタは加算)。
      // 自動転記フィールドのみ既存 JSON にマージし、他フィールドは保全。
      var dayName = DAY_NAMES[dayKey];
      if (dayName) {
        var ekey = w + '|' + dayName;
        if (elecIdx[ekey] !== undefined) {
          for (var tk in transcribe) elecIdx[ekey].json[tk] = transcribe[tk];
          elecIdx[ekey].dirty = true;
        } else if (elecNew[ekey] !== undefined) {
          for (var tk2 in transcribe) elecNew[ekey].json[tk2] = transcribe[tk2];
        } else {
          var nj = {};
          for (var tk3 in transcribe) nj[tk3] = transcribe[tk3];
          elecNew[ekey] = { w: w, dayName: dayName, json: nj };
          elecNewOrder.push(ekey);
        }
      }
      transcribedDays++;
    }
  }

  // --- batched 書込 ---
  // 水質上段
  flushRowUpdates(sh, wuUpdates, 1);
  if (wuAppends.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, wuAppends.length, 4).setValues(wuAppends);
  }
  // 電気: 既存行更新 (dirty のみ) + 新規行一括追加 (出現順)
  var elecUpdates = [];
  for (var ukey in elecIdx) {
    if (elecIdx[ukey].dirty) {
      var e = elecIdx[ukey];
      elecUpdates.push({ r: e.r, v: [month, e.w, e.dayName, JSON.stringify(e.json), now] });
    }
  }
  flushRowUpdates(elecSh, elecUpdates, 1);
  if (elecNewOrder.length > 0) {
    var elecAppends = elecNewOrder.map(function (key) {
      var n = elecNew[key];
      return [month, n.w, n.dayName, JSON.stringify(n.json), now];
    });
    elecSh.getRange(elecSh.getLastRow() + 1, 1, elecAppends.length, 5).setValues(elecAppends);
  }

  return jsonResponse({
    status: 'ok',
    message: '水質上段 ' + savedWeeks + ' 週保存 / 電気自動転記 ' + transcribedDays + ' 日',
    savedWeeks: savedWeeks,
    transcribedDays: transcribedDays
  });
}

// per-day オブジェクト ({mon, tue, …, nextMon}) から指定曜日の値を取り出す。
// 未入力 ('' / null / undefined) は null を返す。
function pickDayValue(obj, dayKey) {
  if (!obj || typeof obj !== 'object') return null;
  var v = obj[dayKey];
  if (v === null || v === undefined || v === '') return null;
  return v;
}

// [Phase B] 旧 mergeIntoElectricalDay は saveWaterUsage 内にインライン化。
// 電気シートの索引を 1 回構築し (col4 JSON 同時取得)、自動転記フィールドのみ
// in-memory でマージ → batched 書込。per-row getValue / appendRow を全廃。

// 下段水質測定 (per-day, JSON)
// Phase A: N+1 根絶。全週 payload (data.weeks = { [week]: { [dayCode]: {...} } }) を
// 受領し、索引マップ 1 回読み + flushRowUpdates で batched 書込。
// 旧 single-week 形式 (data.week + data.days) も後方互換で受付 (saveEverything 経由)。
function saveWaterMeasure(data) {
  var month = normalizeMonth(data.month);
  var weeks = data.weeks;
  if (!weeks && data.week != null && data.days) {
    weeks = {}; weeks[data.week] = data.days;   // 後方互換: 単一週 → 全週形へラップ
  }
  if (!month || !weeks) return jsonResponse({ error: 'パラメータ不足' });

  var sh = getOrCreateSheet(SHEET_WATER_MEASURE, HEADERS_WATER_MEASURE);
  var now = new Date();

  // 索引マップ: 'week|曜日' → 絶対行 (当月のみ・先勝ち)
  var lr = sh.getLastRow();
  var rowByKey = {};
  if (lr > 1) {
    var grid = sh.getRange(2, 1, lr - 1, 3).getValues();
    for (var i = 0; i < grid.length; i++) {
      if (!matchMonth(grid[i][0], month)) continue;
      var gw = parseInt(grid[i][1], 10); if (!gw) continue;
      var gkey = gw + '|' + String(grid[i][2]);
      if (rowByKey[gkey] === undefined) rowByKey[gkey] = i + 2;
    }
  }

  var updates = [], appends = [], cnt = 0;
  for (var wn in weeks) {
    var days = weeks[wn]; if (!days) continue;
    var w = parseInt(wn, 10); if (!w) continue;
    for (var dc in days) {
      var dd = days[dc]; if (!dd) continue;
      var dn = DAY_NAMES[dc]; if (!dn) continue;
      var rv = [month, w, dn, JSON.stringify(dd), now];
      var key = w + '|' + dn;
      if (rowByKey[key] !== undefined) updates.push({ r: rowByKey[key], v: rv });
      else appends.push(rv);
      cnt++;
    }
  }
  flushRowUpdates(sh, updates, 1);
  if (appends.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, 5).setValues(appends);
  }
  return jsonResponse({ status: 'ok', message: '水質下段 ' + cnt + ' 日保存', savedCount: cnt });
}

// 機器運転時間 (per-week, JSON)
//   data.equipment = { [week]: { [machineKey]: { mon, tue, wed, thu, fri } } }
function saveEquipment(data) {
  var month = normalizeMonth(data.month);
  var eq = data.equipment;
  if (!month || !eq) return jsonResponse({ error: 'パラメータ不足' });

  var sh = getOrCreateSheet(SHEET_EQUIPMENT, HEADERS_EQUIPMENT);
  var now = new Date();

  // 索引マップ: week → 絶対行 (当月のみ・先勝ち)
  var lr = sh.getLastRow();
  var rowByWeek = {};
  if (lr > 1) {
    var grid = sh.getRange(2, 1, lr - 1, 2).getValues();
    for (var i = 0; i < grid.length; i++) {
      if (!matchMonth(grid[i][0], month)) continue;
      var gw = parseInt(grid[i][1], 10); if (!gw) continue;
      if (rowByWeek[gw] === undefined) rowByWeek[gw] = i + 2;
    }
  }

  var updates = [], appends = [], cnt = 0;
  for (var wn in eq) {
    var wd = eq[wn]; if (!wd) continue;
    var w = parseInt(wn, 10); if (!w) continue;
    var rv = [month, w, JSON.stringify(wd), now];
    if (rowByWeek[w] !== undefined) updates.push({ r: rowByWeek[w], v: rv });
    else appends.push(rv);
    cnt++;
  }
  flushRowUpdates(sh, updates, 1);
  if (appends.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, 4).setValues(appends);
  }
  return jsonResponse({ status: 'ok', message: '機器運転時間 ' + cnt + ' 週保存', savedCount: cnt });
}

// 日常電気設備 (per-day, JSON)
//   data.electrical = { [week]: { [day]: { fieldKey: value, … } } }
function saveElectrical(data) {
  var month = normalizeMonth(data.month);
  var el = data.electrical;
  if (!month || !el) return jsonResponse({ error: 'パラメータ不足' });

  var sh = getOrCreateSheet(SHEET_ELECTRICAL, HEADERS_ELECTRICAL);
  var now = new Date();

  // 索引マップ: 'week|曜日' → { r: 絶対行, json: 既存data(col4) } (当月のみ・先勝ち)
  // 既存の自動転記値 (saveWaterUsage 由来) を破壊しないよう既存 JSON とマージするため、
  // col4(data) も索引構築時に同時取得し per-row getValue を全廃。
  var lr = sh.getLastRow();
  var idx = {};
  if (lr > 1) {
    var grid = sh.getRange(2, 1, lr - 1, 4).getValues();
    for (var i = 0; i < grid.length; i++) {
      if (!matchMonth(grid[i][0], month)) continue;
      var gw = parseInt(grid[i][1], 10); if (!gw) continue;
      var gkey = gw + '|' + String(grid[i][2]);
      if (idx[gkey] === undefined) {
        var gex = {};
        try { gex = JSON.parse(grid[i][3] || '{}'); } catch (e) { gex = {}; }
        idx[gkey] = { r: i + 2, json: gex };
      }
    }
  }

  var updates = [], appends = [], cnt = 0;
  for (var wn in el) {
    var wd = el[wn]; if (!wd) continue;
    var w = parseInt(wn, 10); if (!w) continue;
    for (var dc in wd) {
      var dd = wd[dc]; if (!dd) continue;
      var dn = DAY_NAMES[dc]; if (!dn) continue;
      var key = w + '|' + dn;
      var merged;
      if (idx[key] !== undefined) {
        merged = idx[key].json;
        for (var k in dd) merged[k] = dd[k];
        updates.push({ r: idx[key].r, v: [month, w, dn, JSON.stringify(merged), now] });
      } else {
        merged = {};
        for (var k2 in dd) merged[k2] = dd[k2];
        appends.push([month, w, dn, JSON.stringify(merged), now]);
      }
      cnt++;
    }
  }
  flushRowUpdates(sh, updates, 1);
  if (appends.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, 5).setValues(appends);
  }
  return jsonResponse({ status: 'ok', message: '電気 ' + cnt + ' 日保存', savedCount: cnt });
}

// 日常機械設備 (per-day, JSON)
//   data.mechanical = { [week]: { [day]: { fieldKey: value, … } } }
function saveMechanical(data) {
  var month = normalizeMonth(data.month);
  var mech = data.mechanical;
  if (!month || !mech) return jsonResponse({ error: 'パラメータ不足' });

  var sh = getOrCreateSheet(SHEET_MECHANICAL, HEADERS_MECHANICAL);
  var now = new Date();

  // 索引マップ: 'week|曜日' → 絶対行 (当月のみ・先勝ち)
  var lr = sh.getLastRow();
  var rowByKey = {};
  if (lr > 1) {
    var grid = sh.getRange(2, 1, lr - 1, 3).getValues();
    for (var i = 0; i < grid.length; i++) {
      if (!matchMonth(grid[i][0], month)) continue;
      var gw = parseInt(grid[i][1], 10); if (!gw) continue;
      var gkey = gw + '|' + String(grid[i][2]);
      if (rowByKey[gkey] === undefined) rowByKey[gkey] = i + 2;
    }
  }

  var updates = [], appends = [], cnt = 0;
  for (var wn in mech) {
    var wd = mech[wn]; if (!wd) continue;
    var w = parseInt(wn, 10); if (!w) continue;
    for (var dc in wd) {
      var dd = wd[dc]; if (!dd) continue;
      var dn = DAY_NAMES[dc]; if (!dn) continue;
      var rv = [month, w, dn, JSON.stringify(dd), now];
      var key = w + '|' + dn;
      if (rowByKey[key] !== undefined) updates.push({ r: rowByKey[key], v: rv });
      else appends.push(rv);
      cnt++;
    }
  }
  flushRowUpdates(sh, updates, 1);
  if (appends.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, 5).setValues(appends);
  }
  return jsonResponse({ status: 'ok', message: '機械 ' + cnt + ' 日保存', savedCount: cnt });
}

// オーケストレータ: 全データを 1 リクエストで保存
function saveEverything(data) {
  var month = normalizeMonth(data.month);
  if (!month) return jsonResponse({ error: '月が指定されていません' });
  var msgs = [];
  if (data.waterUsage) {
    var r = saveWaterUsage({ month: month, weeks: data.waterUsage });
    msgs.push(JSON.parse(r.getContent()).message || '');
  }
  if (data.waterMeasure) {
    var totalDays = 0;
    for (var wn in data.waterMeasure) {
      var r2 = saveWaterMeasure({ month: month, week: parseInt(wn, 10), days: data.waterMeasure[wn] });
      try { totalDays += JSON.parse(r2.getContent()).savedCount || 0; } catch (e) {}
    }
    msgs.push('水質下段 ' + totalDays + ' 日保存');
  }
  if (data.equipment) {
    var r3 = saveEquipment({ month: month, equipment: data.equipment });
    msgs.push(JSON.parse(r3.getContent()).message || '');
  }
  if (data.electrical) {
    var r4 = saveElectrical({ month: month, electrical: data.electrical });
    msgs.push(JSON.parse(r4.getContent()).message || '');
  }
  if (data.mechanical) {
    var r5 = saveMechanical({ month: month, mechanical: data.mechanical });
    msgs.push(JSON.parse(r5.getContent()).message || '');
  }
  return jsonResponse({ status: 'ok', message: msgs.filter(function(m){return m;}).join(' / ') });
}


// =====================================================================
// 全データ取得 (旧 tutuga_gscode.gs の戻り値構造を踏襲)
// =====================================================================

function getAllData(rawMonth) {
  var month = normalizeMonth(rawMonth);
  if (!month) return jsonResponse({ error: '月が指定されていません' });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {
    month: month,
    waterUsage:   {},
    waterMeasure: {},
    equipment:    {},
    electrical:   {},
    mechanical:   {},
    holidays:     []
  };

  // 日常水質_上段 (per-week)
  var ws = ss.getSheetByName(SHEET_WATER_USAGE);
  if (ws && ws.getLastRow() > 1) {
    var v = ws.getRange(2, 1, ws.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < v.length; i++) {
      if (!matchMonth(v[i][0], month)) continue;
      var w = parseInt(v[i][1], 10); if (!w) continue;
      try { result.waterUsage[w] = JSON.parse(v[i][2] || '{}'); } catch (e) {}
    }
  }

  // 日常水質_下段 (per-day)
  var wsd = ss.getSheetByName(SHEET_WATER_MEASURE);
  if (wsd && wsd.getLastRow() > 1) {
    var v2 = wsd.getRange(2, 1, wsd.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < v2.length; i++) {
      if (!matchMonth(v2[i][0], month)) continue;
      var w = parseInt(v2[i][1], 10); if (!w) continue;
      var dc = DAY_REVERSE[v2[i][2]]; if (!dc) continue;
      result.waterMeasure[w] = result.waterMeasure[w] || {};
      try { result.waterMeasure[w][dc] = JSON.parse(v2[i][3] || '{}'); } catch (e) {}
    }
  }

  // 機器運転時間 (per-week)
  var es = ss.getSheetByName(SHEET_EQUIPMENT);
  if (es && es.getLastRow() > 1) {
    var v3 = es.getRange(2, 1, es.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < v3.length; i++) {
      if (!matchMonth(v3[i][0], month)) continue;
      var w = parseInt(v3[i][1], 10); if (!w) continue;
      try { result.equipment[w] = JSON.parse(v3[i][2] || '{}'); } catch (e) {}
    }
  }

  // 日常電気設備 (per-day)
  var els = ss.getSheetByName(SHEET_ELECTRICAL);
  if (els && els.getLastRow() > 1) {
    var v4 = els.getRange(2, 1, els.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < v4.length; i++) {
      if (!matchMonth(v4[i][0], month)) continue;
      var w = parseInt(v4[i][1], 10); if (!w) continue;
      var dc = DAY_REVERSE[v4[i][2]]; if (!dc) continue;
      result.electrical[w] = result.electrical[w] || {};
      try { result.electrical[w][dc] = JSON.parse(v4[i][3] || '{}'); } catch (e) {}
    }
  }

  // 日常機械設備 (per-day)
  var ms = ss.getSheetByName(SHEET_MECHANICAL);
  if (ms && ms.getLastRow() > 1) {
    var v5 = ms.getRange(2, 1, ms.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < v5.length; i++) {
      if (!matchMonth(v5[i][0], month)) continue;
      var w = parseInt(v5[i][1], 10); if (!w) continue;
      var dc = DAY_REVERSE[v5[i][2]]; if (!dc) continue;
      result.mechanical[w] = result.mechanical[w] || {};
      try { result.mechanical[w][dc] = JSON.parse(v5[i][3] || '{}'); } catch (e) {}
    }
  }

  // 祝日リスト (全件、クライアントで月でフィルタする)
  var hs = ss.getSheetByName(SHEET_HOLIDAYS);
  if (hs && hs.getLastRow() >= 3) {
    var v6 = hs.getRange(3, 1, hs.getLastRow() - 2, 2).getValues();
    for (var i = 0; i < v6.length; i++) {
      var d = v6[i][0];
      if (!d) continue;
      if (typeof d === 'object' && typeof d.getFullYear === 'function') {
        var y = d.getFullYear();
        var mm = ('0' + (d.getMonth() + 1)).slice(-2);
        var dd = ('0' + d.getDate()).slice(-2);
        result.holidays.push({ date: y + '-' + mm + '-' + dd, name: v6[i][1] || '' });
      } else {
        var ds = toDateStr(d);
        if (ds) result.holidays.push({ date: ds, name: v6[i][1] || '' });
      }
    }
  }

  return jsonResponse(result);
}


// =====================================================================
// 初期セットアップ用ヘルパー (GAS エディタから手動実行)
// =====================================================================

function setupHolidays() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_HOLIDAYS);
  if (sheet && sheet.getLastRow() >= 3) return '既に「祝日リスト」シートが存在します。';
  if (!sheet) sheet = ss.insertSheet(SHEET_HOLIDAYS);
  sheet.getRange(1, 1, 2, 2).setValues([['祝日一覧', ''], ['日付', '名称']]);
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.getRange(2, 1, 1, 2).setFontWeight('bold').setBackground('#e0f2e9');
  var holidays = [
    ['2026-01-01', '元日'], ['2026-01-02', '振替休日'], ['2026-01-12', '成人の日'],
    ['2026-02-11', '建国記念の日'], ['2026-02-23', '天皇誕生日'], ['2026-03-20', '春分の日'],
    ['2026-04-29', '昭和の日'], ['2026-05-03', '憲法記念日'], ['2026-05-04', 'みどりの日'],
    ['2026-05-05', 'こどもの日'], ['2026-05-06', '振替休日'],
    ['2026-07-20', '海の日'], ['2026-08-11', '山の日'], ['2026-09-21', '敬老の日'],
    ['2026-09-23', '秋分の日'], ['2026-10-12', 'スポーツの日'],
    ['2026-11-03', '文化の日'], ['2026-11-23', '勤労感謝の日']
  ];
  var rows = holidays.map(function (h) {
    var p = h[0].split('-').map(Number);
    return [new Date(p[0], p[1] - 1, p[2]), h[1]];
  });
  sheet.getRange(3, 1, rows.length, 2).setValues(rows);
  sheet.getRange(3, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  return '「祝日リスト」シートを作成しました（' + rows.length + '件）。';
}

function setupSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (sheet) return '既に「設定」シートが存在します。';
  sheet = ss.insertSheet(SHEET_SETTINGS);
  sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e0f2e9');
  sheet.getRange(2, 1, 1, 2).setValues([['テンプレートID', '1lcdNltnqxjIYAK1dxZKLZ3D8ubNx6fjGYtsz-5ClqZg']]);
  return '「設定」シートを作成しました。';
}

function setupUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (sheet) return '既に「ユーザー管理」シートが存在します。';
  sheet = ss.insertSheet(SHEET_USERS);
  sheet.getRange(1, 1, 1, 6).setValues([['ID', 'パスワードハッシュ', '氏名', '有効', '最終ログイン', '備考']]);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e0f2e9');
  sheet.getRange(2, 1, 1, 6).setValues([['cpkanri', '(未使用)', '筒賀', true, '', 'Phase 3 では TUTUGA_USERS 配列を使用、本シートは将来の拡張用']]);
  return '「ユーザー管理」シートを作成しました。';
}

function setupAll() {
  var msgs = [setupHolidays(), setupSettings(), setupUsers()];
  return msgs.join('\n');
}
