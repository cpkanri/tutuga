// tutuga Excel 書込マップ (Phase 38-b5)
// 目的: template.js の write 関数群 (L416 以降) の write 構造を 1 か所に集約した reference document。
//       既存の TUTUGA_MEASURE_ROW_OFFSET / TUTUGA_USAGE_ROW_OFFSET 等の first-class 定数の
//       薄い view として動作。実 write 経路の動的化は Phase 38-b5+ Step 1 (null skip 動的化) のみ。
//
// 設計: tutuga 固有 (5 アプリ間で唯一の 2-line plant)
// base commit: 055354b (Phase 23-B-3、6 key null skip 適用済)
// 抽出: Phase 38-b5 Stage 1 報告 (2026-05-25)
//
// 重要: tutuga は他 4 アプリと異なり、書込関数が template.js に分離されている。
// excel-write-map.js は template.js より **前** に読込まれる必要がある。

(function () {
  'use strict';

  // ===== シート定義 =====
  // tutuga 固有: wb.worksheets[idx] で 0-based index ベース取得
  // (shiwagi: 1-based getWorksheet、yoshiwa/kake/kamitono: name-based の第 3 方式)
  const SHEETS = {
    DAILY_WATER:       { index: 0 },  // 日常水質 (上段 USAGE + 下段 MEASURE) ✅ 実装済
    OPERATION_MONTHLY: { index: 1 },  // 運転管理月報 (Phase 5+ 未実装、A6 日付のみ書込)
    WATER_REPORT:      { index: 2 },  // 水質管理報告 (Phase 5+ 未実装、固定値のみ書込)
    DAILY_MECH:        { index: 3 },  // 日常機械設備 (155 項目、別経路で MAP 対象外)
    DAILY_ELEC:        { index: 4 },  // 日常電気設備 (67 項目、別経路で MAP 対象外)
    EQUIPMENT_HOURS:   { index: 5 }   // 主要機器運転時間 (22 機器、別経路で MAP 対象外)
  };

  // ===== helper 関数 =====
  // _daily: daily 書込位置を返す
  //   row_offset: 週内行オフセット (WATER_BLOCK_ROW[w] に加算する想定)
  //   mon_col:    mon 曜日の列番号 (tue/wed/thu/fri は +1/+2/+3/+4)
  //               tutuga 固有 2 列マップ: USAGE=9 (USAGE_DAY_COL), MEASURE=10 (MEASURE_DAY_COL)
  function _daily(row_offset, mon_col) {
    return { sheet: SHEETS.DAILY_WATER, row_offset: row_offset, mon_col: mon_col };
  }

  // _entry: 統一エントリー
  // 注: monthly は Phase 5+ (water_report 実装時) まで全件 null
  function _entry(daily, monthly, null_skip) {
    return {
      storage_unit: 'key',
      section: null,
      daily: daily || null,
      monthly: monthly || null,
      null_skip: !!null_skip
    };
  }

  // ===== 書込マップ本体 (37 entries: MEASURE 28 + USAGE 9) =====
  // 2026-06-22: 旧 Phase 2 追加分 solidChlorine(USAGE) / do1・do2(MEASURE) を廃止。
  //   solidChlorine→運転管理月報 G 列は chlorineDose へ、do1/do2→水質管理報告 I/J 列は
  //   電気設備の DO 計値へそれぞれ統合・再ソース。実 write 経路 (template.js の reportDayMap
  //   駆動) が source of truth で、この map の static スロットは reference document。
  const EXCEL_WRITE_MAP = {
    // ====================================================
    // USAGE_KEYS (上段使用量、ws=0 daily_water, mon_col=9)
    // 9 entries (excessSludge1/2 は ws=4 daily_elec sub2 転記のため除外)
    // ====================================================
    powerAllDay:    _entry(_daily(3,  9), null),
    powerMeasure:   _entry(_daily(4,  9), null),
    powerReactive:  _entry(_daily(5,  9), null),
    demand:         _entry(_daily(6,  9), null),
    water:          _entry(_daily(7,  9), null),
    diesel:         _entry(_daily(8,  9), null),
    returnSludge1:  _entry(_daily(9,  9), null),
    returnSludge2:  _entry(_daily(10, 9), null),
    discharge:      _entry(_daily(11, 9), null),
    // 旧 solidChlorine(固形塩素使用量 → 運転管理月報 G 列) は廃止し chlorineDose へ統合。
    // 実 G 列書込は template.js _tutugaWriteOperationDaily が chlorineDose から供給する。

    // ====================================================
    // MEASURE_FIELDS (下段水質測定、ws=0 daily_water, mon_col=10)
    // 28 entries (daily 20 + null_skip 6 + metadata 2)
    // ====================================================

    // --- 水温 (流入 + ディッチ 1/2 + 放流) ---
    tempIn:         _entry(_daily(15, 10), null),
    tempD1:         _entry(_daily(16, 10), null),
    tempD2:         _entry(_daily(17, 10), null),
    tempOut:        _entry(_daily(18, 10), null),
    // --- 水温 (終沈 1/2 → 水質管理報告専用、Phase 5+ で Excel 書込実装) ---
    tempFinal1:     _entry(null, null, true),  // null_skip 対象 (Phase 23-B-3、1 系終沈)
    tempFinal2:     _entry(null, null, true),  // null_skip 対象 (Phase 23-B-3、2 系終沈)

    // --- 透視度 (流入 + 放流) ---
    transIn:        _entry(_daily(23, 10), null),
    transOut:       _entry(_daily(24, 10), null),
    // --- 透視度 (終沈 1/2) ---
    transFinal1:    _entry(null, null, true),  // null_skip 対象
    transFinal2:    _entry(null, null, true),  // null_skip 対象

    // --- PH (流入 + ディッチ 1/2 + 放流) ---
    phIn:           _entry(_daily(29, 10), null),
    phD1:           _entry(_daily(30, 10), null),
    phD2:           _entry(_daily(31, 10), null),
    phOut:          _entry(_daily(32, 10), null),
    // --- PH (終沈 1/2) ---
    phFinal1:       _entry(null, null, true),  // null_skip 対象
    phFinal2:       _entry(null, null, true),  // null_skip 対象

    // 旧 do1/do2(水質測定タブの測定DO → 水質管理報告 2/2 I/J 列) は廃止。
    // 実 I/J 列書込は template.js _tutugaWriteSuikanDaily が電気設備の DO 計値から再ソースする。

    // --- SV (1 系 / 2 系 並列処理場の特徴) ---
    sv1_30:         _entry(_daily(33, 10), null),
    sv1_24h:        _entry(_daily(34, 10), null),
    sv2_30:         _entry(_daily(35, 10), null),
    sv2_24h:        _entry(_daily(36, 10), null),

    // --- MLSS (1 系 / 2 系) ---
    mlss1:          _entry(_daily(37, 10), null),
    mlss2:          _entry(_daily(38, 10), null),

    // --- 汚泥界面 (1 系終沈 / 2 系終沈) ---
    sludge1:        _entry(_daily(39, 10), null),
    sludge2:        _entry(_daily(40, 10), null),

    // --- 塩素 (残留 / 投入) ---
    chlorine:       _entry(_daily(41, 10), null),
    // 塩素投入量: daily_water(42,10) に加え、運転管理月報(ws=1) G 列(col:7) にも出力
    //   （旧 solidChlorine=固形塩素使用量 と同一値のため統合）。実書込は template.js が担う。
    chlorineDose:   _entry(_daily(42, 10), { sheet: SHEETS.OPERATION_MONTHLY, col: 7 }),

    // --- メタデータ (Excel 書込先なし、client 専用) ---
    inspector:      _entry(null, null),  // 測定者
    bikou:          _entry(null, null)   // 備考
  };

  // ===== NULL_SKIP_KEYS_DIRECT (動的算出、Phase 23-B-3 適用 6 キー) =====
  const NULL_SKIP_KEYS_DIRECT = Object.keys(EXCEL_WRITE_MAP)
    .filter(function (k) { return EXCEL_WRITE_MAP[k].null_skip; });

  // ===== 期待値 (index.html L1805-1834 と整合) =====
  // MEASURE_FIELDS (28 件、metadata 含む。do1/do2 廃止)
  const MEASURE_FIELDS_EXPECTED = [
    'tempIn', 'tempD1', 'tempD2', 'tempOut', 'tempFinal1', 'tempFinal2',
    'transIn', 'transOut', 'transFinal1', 'transFinal2',
    'phIn', 'phD1', 'phD2', 'phOut', 'phFinal1', 'phFinal2',
    'sv1_30', 'sv1_24h', 'sv2_30', 'sv2_24h',
    'mlss1', 'mlss2',
    'sludge1', 'sludge2',
    'chlorine', 'chlorineDose',
    'inspector', 'bikou'
  ];
  // USAGE_KEYS (11 件、ただし excessSludge1/2 は MAP 対象外。solidChlorine 廃止)
  const USAGE_KEYS_EXPECTED = [
    'powerAllDay', 'powerMeasure', 'powerReactive', 'demand',
    'water', 'diesel',
    'returnSludge1', 'returnSludge2', 'discharge',
    'excessSludge1', 'excessSludge2'  // 別シート転記、MAP には含めない
  ];
  // MAP に含めるべき keys = MEASURE 全件 + USAGE 9 件 (excessSludge1/2 除外)
  const EXPECTED_MAP_KEYS = MEASURE_FIELDS_EXPECTED.concat(
    USAGE_KEYS_EXPECTED.filter(function (k) {
      return k !== 'excessSludge1' && k !== 'excessSludge2';
    })
  );

  // ===== 整合性検証 (6 種) =====
  function verifyAgainstTutugaFields() {
    const issues = [];
    const mapKeys = Object.keys(EXCEL_WRITE_MAP);

    // Verify 1: MAP entries 数 = 37 (do1/do2/solidChlorine 廃止: 40 − 3)
    if (mapKeys.length !== 37) {
      issues.push('MAP entries count mismatch: expected 37, got ' + mapKeys.length);
    }

    // Verify 2: MAP keys ↔ EXPECTED_MAP_KEYS 双方向 (37 件、MEASURE 28 + USAGE 9)
    const missingInMap = EXPECTED_MAP_KEYS.filter(function (k) { return mapKeys.indexOf(k) < 0; });
    const extraInMap   = mapKeys.filter(function (k) { return EXPECTED_MAP_KEYS.indexOf(k) < 0; });
    if (missingInMap.length) issues.push('MAP missing keys: ' + missingInMap.join(','));
    if (extraInMap.length)   issues.push('MAP extra keys: '   + extraInMap.join(','));

    // Verify 3: NULL_SKIP_KEYS_DIRECT 期待値 (6 件、Phase 23-B-3 整合)
    const expectedNullSkip = ['tempFinal1', 'tempFinal2', 'transFinal1', 'transFinal2', 'phFinal1', 'phFinal2'];
    const actualSorted   = NULL_SKIP_KEYS_DIRECT.slice().sort();
    const expectedSorted = expectedNullSkip.slice().sort();
    if (JSON.stringify(actualSorted) !== JSON.stringify(expectedSorted)) {
      issues.push('null_skip mismatch: expected=[' + expectedSorted.join(',') + '] actual=[' + actualSorted.join(',') + ']');
    }

    // Verify 4: 6 null_skip keys は daily=null, monthly=null
    for (let i = 0; i < expectedNullSkip.length; i++) {
      const k = expectedNullSkip[i];
      const entry = EXCEL_WRITE_MAP[k];
      if (!entry) { issues.push(k + ': entry missing'); continue; }
      if (entry.daily !== null) issues.push(k + ': null_skip key daily should be null');
      if (entry.monthly !== null) issues.push(k + ': null_skip key monthly should be null (Phase 5+ 未実装)');
    }

    // Verify 5: SHEETS index 0-5 整合 (tutuga 固有 0-based)
    const expectedIndices = [0, 1, 2, 3, 4, 5];
    const actualIndices = Object.keys(SHEETS).map(function (k) { return SHEETS[k].index; }).sort(function (a, b) { return a - b; });
    if (JSON.stringify(actualIndices) !== JSON.stringify(expectedIndices)) {
      issues.push('SHEETS indices mismatch: expected=' + JSON.stringify(expectedIndices) + ' actual=' + JSON.stringify(actualIndices));
    }

    // Verify 6: daily entries の mon_col ∈ {9, 10} (tutuga 2 列マップ整合)
    for (let i = 0; i < mapKeys.length; i++) {
      const k = mapKeys[i];
      const entry = EXCEL_WRITE_MAP[k];
      if (entry.daily && [9, 10].indexOf(entry.daily.mon_col) < 0) {
        issues.push(k + ': daily.mon_col not in {9,10} (' + entry.daily.mon_col + ')');
      }
    }

    if (issues.length) {
      console.error('[excel-write-map.js] verify FAILED (' + issues.length + ' issues):');
      issues.forEach(function (msg) { console.error('  - ' + msg); });
      return false;
    }
    console.log('[excel-write-map.js] verify OK (37 entries, 6 null_skip, 6 sheets, mon_col groups {9,10})');
    return true;
  }

  // ===== グローバル公開 =====
  window.EXCEL_WRITE_MAP = EXCEL_WRITE_MAP;
  window.EXCEL_WRITE_MAP_SHEETS = SHEETS;
  window.NULL_SKIP_KEYS_DIRECT = NULL_SKIP_KEYS_DIRECT;
  window.EXCEL_WRITE_MAP_NULL_SKIP_KEYS_DIRECT = NULL_SKIP_KEYS_DIRECT;
  window.verifyAgainstTutugaFields = verifyAgainstTutugaFields;

  // ===== 自動 verify (読込時に 1 回) =====
  try {
    verifyAgainstTutugaFields();
  } catch (e) {
    console.error('[excel-write-map.js] verify threw:', e);
  }
})();
