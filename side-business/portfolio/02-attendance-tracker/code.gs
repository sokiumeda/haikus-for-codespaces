/**
 * ============================================================================
 *  勤怠管理・自動集計ツール（ポートフォリオ用デモ）
 * ============================================================================
 * 想定利用者：中小企業の総務・経理担当者（非エンジニア）
 * 目的      ：タイムカード（打刻記録）の手集計をなくし、月次の実働時間・
 *             残業時間・深夜時間の自動計算と、残業アラートの表示を行う。
 *
 * 使い方はすべて Google スプレッドシートのカスタムメニュー
 * 「勤怠管理ツール」から実行できます（詳細は README.md 参照）。
 *
 * 外部サービスとの通信は一切行いません（GAS標準機能のみで完結）。
 * ============================================================================
 */

// ============================================================================
// 定数定義
// ============================================================================

// シート名
const SHEET_EMP = '従業員マスタ';
const SHEET_PUNCH = '打刻記録';
const SHEET_SUMMARY = '月次集計';
const SHEET_SETTINGS = '設定';
const SHEET_ERROR = 'エラー一覧';

// 各シートのヘッダー行
const EMP_HEADERS = ['社員ID', '氏名', '雇用形態', '時給（パートのみ）', '所定労働時間(h/日)'];
const PUNCH_HEADERS = ['日付', '社員ID', '出勤時刻', '退勤時刻', '休憩時間(分)'];
const SUMMARY_HEADERS = ['社員ID', '氏名', '出勤日数', '実働時間(h)', '残業時間(h)', '深夜時間(h)', '概算支給額(円)', 'アラート'];
const SETTINGS_HEADERS = ['項目', '値', '説明'];
const ERROR_HEADERS = ['発生日時', 'シート', '行番号', 'エラー内容', '詳細'];

// 月次集計シートの列番号（1始まり）。ヘッダーを変更する場合は合わせて変更すること。
const SUM_COL = { ID: 1, NAME: 2, DAYS: 3, WORKED: 4, OVERTIME: 5, NIGHT: 6, PAY: 7, ALERT: 8 };

// 設定シートの「項目」ラベル
const SETTING_KEYS = {
  ALERT_THRESHOLD: '残業アラート閾値(h/月)',
  NIGHT_START: '深夜開始時刻',
  NIGHT_END: '深夜終了時刻',
  TARGET_MONTH: '集計対象年月'
};

// 1回の勤務としてあり得る上限時間。これを超える場合は入力ミスとみなしエラーにする
// （例：退勤時刻が出勤時刻より前で、かつ日またぎ勤務としても長すぎる＝不正な時刻）
const MAX_SHIFT_HOURS = 20;

// デフォルトの深夜帯（22:00〜翌5:00）・残業アラート閾値
const DEFAULT_NIGHT_START = '22:00';
const DEFAULT_NIGHT_END = '5:00';
const DEFAULT_ALERT_THRESHOLD = 45;

// アラート時の行の背景色
const ALERT_BG_COLOR = '#f4cccc';


// ============================================================================
// メニュー（非エンジニアはここからしか操作しない）
// ============================================================================

/**
 * スプレッドシートを開いたときに自動実行され、カスタムメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('勤怠管理ツール')
    .addItem('① 初期セットアップ', 'initSheets')
    .addItem('② 月次集計の実行', 'runMonthlyAggregation')
    .addItem('③ 残業アラートチェック', 'checkOvertimeAlerts')
    .addToUi();
}


// ============================================================================
// ① 初期セットアップ
// ============================================================================

/**
 * 必要なシート（従業員マスタ／打刻記録／月次集計／設定／エラー一覧）を
 * すべて自動作成する。既に存在するシートのデータは壊さない（冪等）。
 * 新規作成したシートにのみ、サンプルデータ・初期設定値を入れる。
 */
function initSheets() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const empIsNew = !ss.getSheetByName(SHEET_EMP);
    const punchIsNew = !ss.getSheetByName(SHEET_PUNCH);
    const settingsIsNew = !ss.getSheetByName(SHEET_SETTINGS);

    const empSheet = setupSheet_(ss, SHEET_EMP, EMP_HEADERS);
    const punchSheet = setupSheet_(ss, SHEET_PUNCH, PUNCH_HEADERS);
    setupSheet_(ss, SHEET_SUMMARY, SUMMARY_HEADERS);
    const settingsSheet = setupSheet_(ss, SHEET_SETTINGS, SETTINGS_HEADERS);
    setupSheet_(ss, SHEET_ERROR, ERROR_HEADERS);

    // 従業員マスタ：雇用形態はプルダウン選択にしておく（入力ミス防止）
    const empTypeRange = empSheet.getRange(2, 3, 199, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['正社員', 'パート'], true)
      .setAllowInvalid(false)
      .build();
    empTypeRange.setDataValidation(rule);

    // 設定シート：新規作成時のみデフォルト値を書き込む
    if (settingsIsNew) {
      const today = new Date();
      const defaultTargetMonth = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM');
      settingsSheet.getRange(2, 1, 4, 3).setValues([
        [SETTING_KEYS.ALERT_THRESHOLD, DEFAULT_ALERT_THRESHOLD, 'この時間（月間残業時間）を超えたら月次集計シートで警告表示します'],
        [SETTING_KEYS.NIGHT_START, DEFAULT_NIGHT_START, '深夜時間帯の開始時刻（この時刻以降を深夜としてカウント）'],
        [SETTING_KEYS.NIGHT_END, DEFAULT_NIGHT_END, '深夜時間帯の終了時刻（翌日のこの時刻までを深夜としてカウント）'],
        [SETTING_KEYS.TARGET_MONTH, defaultTargetMonth, '月次集計を行う対象の年月（例：2026-07）を入力してください']
      ]);
      settingsSheet.setColumnWidths(1, 3, 220);
    }

    // 従業員マスタ・打刻記録：新規作成時のみデモ用サンプルデータを入れる
    // （動作確認用です。実データを入力する際は削除・上書きしてください）
    if (empIsNew) {
      empSheet.getRange(2, 1, 2, EMP_HEADERS.length).setValues([
        ['E001', '山田 太郎', '正社員', '', 8],
        ['E002', '鈴木 花子', 'パート', 1200, 6]
      ]);
    }
    if (punchIsNew) {
      const today = new Date();
      const y = today.getFullYear();
      const m = today.getMonth(); // 0始まり。今月の日付として使う
      const d = (day) => new Date(y, m, day);
      punchSheet.getRange(2, 1, 6, PUNCH_HEADERS.length).setValues([
        // 通常勤務（正社員・所定8h・残業なし）
        [d(1), 'E001', '9:00', '18:00', 60],
        // 通常勤務（パート）
        [d(2), 'E002', '13:00', '22:00', 45],
        // 日またぎ勤務の例：22:00出勤→翌7:00退勤（このツールの品質確認ポイント）
        [d(3), 'E001', '22:00', '7:00', 60],
        // わざと不正データを入れています（エラー一覧の動作確認用）
        [d(4), 'E999', '9:00', '18:00', 60],   // 社員IDがマスタに存在しない
        [d(5), 'E001', '', '18:00', 60],       // 出勤時刻が空
        [d(6), 'E002', '9:00', '8:30', 30]     // 退勤<出勤かつ日またぎとしても長すぎる不正時刻
      ]);
    }

    ss.setActiveSheet(empSheet);
    ui.alert(
      '初期セットアップ完了',
      '必要なシートを作成しました。\n\n' +
      '・「従業員マスタ」「打刻記録」にはサンプルデータが入っています。実際のデータに置き換えてください。\n' +
      '・「設定」シートで集計対象年月やアラート閾値を確認・変更できます。\n' +
      '準備ができたら、メニューの「② 月次集計の実行」を行ってください。',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('初期セットアップでエラーが発生しました: ' + e.message);
  }
}

/**
 * シートが無ければ作成し、ヘッダー行を整える共通処理。
 * @return {Sheet} 対象シート
 */
function setupSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  // ヘッダー行が未設定（1行目が空）の場合のみヘッダーを書き込む
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeader = firstRow.some((v) => v !== '' && v !== null);
  if (!hasHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8eaed');
  sheet.setFrozenRows(1);
  return sheet;
}


// ============================================================================
// ② 月次集計の実行
// ============================================================================

/**
 * 打刻記録から月次集計を作成するメイン処理。
 * データに問題がある行は処理を止めず「エラー一覧」に記録して読み飛ばす。
 */
function runMonthlyAggregation() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(SHEET_EMP);
    const punchSheet = ss.getSheetByName(SHEET_PUNCH);
    const summarySheet = ss.getSheetByName(SHEET_SUMMARY);
    const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);

    if (!empSheet || !punchSheet || !summarySheet || !settingsSheet) {
      ui.alert('必要なシートが見つかりません。先にメニューの「① 初期セットアップ」を実行してください。');
      return;
    }

    clearErrorSheet_(ss);

    const settings = readSettings_(ss);
    const targetYM = parseTargetYearMonth_(settings.targetMonth);
    if (!targetYM) {
      ui.alert(
        '「設定」シートの「' + SETTING_KEYS.TARGET_MONTH + '」の形式が正しくありません。\n' +
        '「2026-07」のような「年-月」の形式で入力してください。'
      );
      return;
    }

    // --- 従業員マスタの読み込み ---------------------------------------
    const empData = getDataRows_(empSheet, EMP_HEADERS.length);
    const empOrder = [];
    const empMap = {}; // key: 社員ID(文字列) -> {name, type, wage, scheduledHours}
    empData.forEach((row, idx) => {
      const rowNumber = idx + 2;
      const [idRaw, name, type, wage, scheduledHours] = row;
      if (isRowBlank_(row)) return; // 完全な空行は無視

      const id = String(idRaw).trim();
      if (!id) {
        logError_(ss, SHEET_EMP, rowNumber, '社員IDが空です', '氏名:' + name);
        return;
      }

      let sh = extractNumber_(scheduledHours);
      if (sh === null || sh <= 0) {
        logError_(ss, SHEET_EMP, rowNumber, '所定労働時間が未入力または不正なため、8時間として計算します', '社員ID:' + id);
        sh = 8;
      }

      const empType = String(type || '').trim();
      let hourlyWage = extractNumber_(wage);
      if (empType === 'パート' && (hourlyWage === null || hourlyWage <= 0)) {
        logError_(ss, SHEET_EMP, rowNumber, '時給が未入力または不正なため、支給額は0円として計算します', '社員ID:' + id);
        hourlyWage = 0;
      }

      empMap[id] = {
        name: String(name || '').trim() || id,
        type: empType,
        wage: hourlyWage,
        scheduledHours: sh,
        days: 0,
        workedMin: 0,
        overtimeMin: 0,
        nightMin: 0
      };
      empOrder.push(id);
    });

    // --- 打刻記録の読み込み・検証・集計 --------------------------------
    const punchData = getDataRows_(punchSheet, PUNCH_HEADERS.length);
    let errorCount = 0;

    punchData.forEach((row, idx) => {
      const rowNumber = idx + 2;
      if (isRowBlank_(row)) return; // 完全な空行はスキップ（エラーにしない）

      const result = validatePunchRow_(row, empMap, settings);
      if (result.error) {
        logError_(ss, SHEET_PUNCH, rowNumber, result.error, result.detail || '');
        errorCount++;
        return;
      }
      if (result.skip) {
        // 集計対象月と異なる月のデータなので何もしない（エラーではない）
        return;
      }

      const emp = empMap[result.empId];
      emp.days += 1;
      emp.workedMin += result.workedMin;
      emp.overtimeMin += result.overtimeMin;
      emp.nightMin += result.nightMin;
    });

    // --- 月次集計シートへ出力 ------------------------------------------
    const outRows = empOrder.map((id) => {
      const e = empMap[id];
      const payAmount = e.type === 'パート' ? Math.round(e.wage * (e.workedMin / 60)) : '';
      return [
        id,
        e.name,
        e.days,
        round1_(e.workedMin),
        round1_(e.overtimeMin),
        round1_(e.nightMin),
        payAmount,
        '' // アラート列：「③ 残業アラートチェック」で更新
      ];
    });

    writeSummary_(summarySheet, outRows);

    const monthLabel = targetYM.year + '年' + targetYM.month + '月';
    if (errorCount > 0) {
      ui.alert(
        '月次集計 完了',
        '対象月：' + monthLabel + '\n従業員数：' + empOrder.length + '名\n' +
        'データの問題：' + errorCount + '件見つかりました。\n\n' +
        '「エラー一覧」シートを確認し、該当する打刻記録を修正のうえ、再度集計を実行してください。',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '月次集計 完了',
        '対象月：' + monthLabel + '\n従業員数：' + empOrder.length + '名\nエラーはありませんでした。',
        ui.ButtonSet.OK
      );
    }
  } catch (e) {
    ui.alert('月次集計の実行中にエラーが発生しました: ' + e.message);
  }
}

/**
 * 打刻記録1行分を検証し、問題なければ勤務時間の計算結果を返す。
 * 戻り値のパターン：
 *  - { error: string, detail: string }  … エラー（エラー一覧に記録して読み飛ばす）
 *  - { skip: true }                      … 集計対象月と異なるので無視
 *  - { empId, workedMin, overtimeMin, nightMin } … 正常な計算結果
 */
function validatePunchRow_(row, empMap, settings) {
  const [dateRaw, empIdRaw, startRaw, endRaw, breakRaw] = row;
  const detailParts = [
    '日付:' + formatForDetail_(dateRaw),
    '社員ID:' + formatForDetail_(empIdRaw),
    '出勤:' + formatForDetail_(startRaw),
    '退勤:' + formatForDetail_(endRaw)
  ];
  const detail = detailParts.join(' / ');

  // 日付チェック
  if (dateRaw === '' || dateRaw === null || dateRaw === undefined) {
    return { error: '日付が空です', detail: detail };
  }
  const baseDate = (dateRaw instanceof Date) ? dateRaw : new Date(dateRaw);
  if (isNaN(baseDate.getTime())) {
    return { error: '日付の形式が不正です', detail: detail };
  }

  // 社員IDチェック
  const empIdStr = String(empIdRaw || '').trim();
  if (!empIdStr) {
    return { error: '社員IDが空です', detail: detail };
  }
  if (!empMap[empIdStr]) {
    return { error: '社員IDが従業員マスタに存在しません（社員ID: ' + empIdStr + '）', detail: detail };
  }
  const emp = empMap[empIdStr];

  // 対象年月チェック（対象外の月は静かにスキップ＝エラーではない）
  const targetYM = parseTargetYearMonth_(settings.targetMonth);
  if (targetYM && (baseDate.getFullYear() !== targetYM.year || baseDate.getMonth() + 1 !== targetYM.month)) {
    return { skip: true };
  }

  // 出勤・退勤時刻チェック
  const startMin = parseTimeToMinutes_(startRaw);
  if (startMin === null) {
    return {
      error: (startRaw === '' || startRaw === null || startRaw === undefined) ? '出勤時刻が空です' : '出勤時刻の形式が不正です',
      detail: detail
    };
  }
  const endMin = parseTimeToMinutes_(endRaw);
  if (endMin === null) {
    return {
      error: (endRaw === '' || endRaw === null || endRaw === undefined) ? '退勤時刻が空です' : '退勤時刻の形式が不正です',
      detail: detail
    };
  }

  // 休憩時間チェック（「60分」のように単位付きで入力されていても数値部分を読み取る）
  let breakMin = 0;
  if (breakRaw !== '' && breakRaw !== null && breakRaw !== undefined) {
    const parsedBreak = extractNumber_(breakRaw);
    if (parsedBreak === null || parsedBreak < 0) {
      return { error: '休憩時間(分)の形式が不正です', detail: detail };
    }
    breakMin = parsedBreak;
  }

  // 出勤・退勤の日時を組み立てる（日またぎ対応）
  const startDT = buildDateTime_(baseDate, startMin);
  let endDT = buildDateTime_(baseDate, endMin);
  if (endDT.getTime() <= startDT.getTime()) {
    // 退勤が出勤以前の時刻 → 日をまたいだ勤務とみなし、退勤日を+1日する
    // （例：22:00出勤→翌7:00退勤の「7:00」はこの処理で「翌日7:00」になる）
    endDT = new Date(endDT.getTime());
    endDT.setDate(endDT.getDate() + 1);
  }

  const shiftHours = (endDT.getTime() - startDT.getTime()) / (60 * 60 * 1000);
  if (shiftHours >= MAX_SHIFT_HOURS) {
    // 日またぎとして解釈してもなお長すぎる＝入力ミス（例：退勤時刻の入力間違い）
    return {
      error: '退勤時刻が出勤時刻より前で、日またぎ勤務と解釈しても勤務時間が' + MAX_SHIFT_HOURS + '時間以上となるため、不正な時刻データと判断しました',
      detail: detail
    };
  }

  const workedMin = (endDT.getTime() - startDT.getTime()) / 60000 - breakMin;
  if (workedMin <= 0) {
    return { error: '休憩時間が勤務時間以上のため、実働時間が0以下になります', detail: detail + ' / 休憩:' + breakMin + '分' };
  }

  const scheduledMin = emp.scheduledHours * 60;
  const overtimeMin = Math.max(0, workedMin - scheduledMin);

  let nightMin = computeNightMinutes_(startDT, endDT, settings.nightStartMin, settings.nightEndMin);
  if (nightMin > workedMin) nightMin = workedMin; // 休憩控除後の実働時間を超えないよう調整（簡易計算）

  return {
    empId: empIdStr,
    workedMin: workedMin,
    overtimeMin: overtimeMin,
    nightMin: nightMin
  };
}


// ============================================================================
// ③ 残業アラートチェック
// ============================================================================

/**
 * 月次集計シートを見て、残業時間が閾値を超える従業員の行を色付けし、
 * アラート列に超過時間を表示する。閾値以下の行は色・表示をクリアする。
 */
function checkOvertimeAlerts() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName(SHEET_SUMMARY);
    const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);

    if (!summarySheet || !settingsSheet) {
      ui.alert('必要なシートが見つかりません。先にメニューの「① 初期セットアップ」を実行してください。');
      return;
    }

    const lastRow = summarySheet.getLastRow();
    if (lastRow < 2) {
      ui.alert('月次集計のデータがありません。先にメニューの「② 月次集計の実行」を行ってください。');
      return;
    }

    const settings = readSettings_(ss);
    const threshold = settings.alertThreshold;

    const numRows = lastRow - 1;
    const numCols = SUMMARY_HEADERS.length;
    const range = summarySheet.getRange(2, 1, numRows, numCols);
    const values = range.getValues();

    let alertCount = 0;
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      // 完全な空行はスキップ
      if (isRowBlank_(row)) continue;

      const overtimeH = Number(row[SUM_COL.OVERTIME - 1]) || 0;
      const rowRange = summarySheet.getRange(2 + i, 1, 1, numCols);
      const alertCell = summarySheet.getRange(2 + i, SUM_COL.ALERT);

      if (overtimeH > threshold) {
        rowRange.setBackground(ALERT_BG_COLOR);
        const over = Math.round((overtimeH - threshold) * 10) / 10;
        alertCell.setValue('⚠ 閾値超過（+' + over + 'h）');
        alertCount++;
      } else {
        rowRange.setBackground(null);
        alertCell.setValue('');
      }
    }

    ui.alert(
      '残業アラートチェック 完了',
      '残業アラート閾値：' + threshold + 'h/月\n閾値を超えた従業員：' + alertCount + '名',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('残業アラートチェック中にエラーが発生しました: ' + e.message);
  }
}


// ============================================================================
// 時刻・日付の計算ユーティリティ（このツールの心臓部：日またぎ・深夜計算）
// ============================================================================

/**
 * セルの値（Date型 または "HH:mm" 文字列）を「0:00からの経過分」に変換する。
 * 変換できない場合は null を返す。
 */
function parseTimeToMinutes_(value) {
  if (value === '' || value === null || value === undefined) return null;

  if (value instanceof Date) {
    return value.getHours() * 60 + value.getMinutes();
  }

  if (typeof value === 'string') {
    const m = value.trim().match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return null;
    const h = Number(m[1]);
    const mi = Number(m[2]);
    if (h > 29 || mi > 59) return null; // 常識的な範囲を超える値は不正とみなす
    return h * 60 + mi;
  }

  if (typeof value === 'number') {
    // シリアル値などが渡ってきた場合は不正データとして扱う
    return null;
  }

  return null;
}

/**
 * 基準日(baseDate)の 0:00 に、指定した「経過分」を加えた Date を作る。
 * minutesOfDay が 60 や 1440 を超える値でも Date が自動的に繰り上げる。
 */
function buildDateTime_(baseDate, minutesOfDay) {
  const dt = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate(), 0, 0, 0, 0);
  dt.setMinutes(minutesOfDay);
  return dt;
}

/**
 * 2つの時間区間 [aStart, aEnd) と [bStart, bEnd) が重なっている分数を返す。
 */
function overlapMinutes_(aStart, aEnd, bStart, bEnd) {
  const s = Math.max(aStart.getTime(), bStart.getTime());
  const e = Math.min(aEnd.getTime(), bEnd.getTime());
  return e > s ? (e - s) / 60000 : 0;
}

/**
 * 勤務区間 [startDT, endDT) のうち、深夜時間帯（設定シートで指定、既定22:00〜翌5:00）
 * に含まれる分数を計算する。日またぎ勤務にも対応するため、勤務区間が含む日を
 * 1日ずつ走査し、その日の深夜帯との重なりを合計する。
 */
function computeNightMinutes_(startDT, endDT, nightStartMin, nightEndMin) {
  let total = 0;
  // 前日22:00起点の深夜帯（〜当日5:00）と重なるケース（深夜0時台の出勤など）を
  // 取りこぼさないよう、走査は勤務開始日の前日から始める
  const cursor = new Date(startDT.getFullYear(), startDT.getMonth(), startDT.getDate() - 1);
  const endDay = new Date(endDT.getFullYear(), endDT.getMonth(), endDT.getDate());

  // 無限ループ防止のための安全装置（通常はあり得ないが念のため）
  let guard = 0;
  while (cursor.getTime() <= endDay.getTime() && guard < 40) {
    const winStart = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate());
    winStart.setMinutes(nightStartMin);

    const winEnd = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate());
    if (nightEndMin <= nightStartMin) {
      // 深夜帯が日をまたぐ設定（例：22:00〜翌5:00）の場合
      winEnd.setDate(winEnd.getDate() + 1);
    }
    winEnd.setMinutes(nightEndMin);

    total += overlapMinutes_(startDT, endDT, winStart, winEnd);
    cursor.setDate(cursor.getDate() + 1);
    guard++;
  }
  return total;
}

/** 分を時間(h)に変換し、小数第1位で丸める。 */
function round1_(minutes) {
  return Math.round((minutes / 60) * 10) / 10;
}

/**
 * セルの値から数値部分だけを読み取る（「8時間」「60分」のように単位付きで
 * 入力された場合でも、非エンジニアの入力ミスとして処理を止めないための緩和策）。
 * 数値が見つからない場合は null を返す。
 */
function extractNumber_(value) {
  if (value === '' || value === null || value === undefined) return null;
  if (typeof value === 'number') return value;
  const m = String(value).trim().match(/-?\d+(\.\d+)?/);
  if (!m) return null;
  const n = Number(m[0]);
  return isNaN(n) ? null : n;
}


// ============================================================================
// 設定シートの読み込み
// ============================================================================

/**
 * 「設定」シートを項目名で検索し、必要な設定値をまとめて返す。
 * 値が入っていない・不正な場合はデフォルト値を使う（設定不備で処理停止させない）。
 */
function readSettings_(ss) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  const map = {};
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      values.forEach((row) => {
        const key = String(row[0] || '').trim();
        if (key) map[key] = row[1];
      });
    }
  }

  let alertThreshold = Number(map[SETTING_KEYS.ALERT_THRESHOLD]);
  if (isNaN(alertThreshold) || alertThreshold <= 0) alertThreshold = DEFAULT_ALERT_THRESHOLD;

  let nightStartMin = parseTimeToMinutes_(map[SETTING_KEYS.NIGHT_START]);
  if (nightStartMin === null) nightStartMin = parseTimeToMinutes_(DEFAULT_NIGHT_START);

  let nightEndMin = parseTimeToMinutes_(map[SETTING_KEYS.NIGHT_END]);
  if (nightEndMin === null) nightEndMin = parseTimeToMinutes_(DEFAULT_NIGHT_END);

  return {
    alertThreshold: alertThreshold,
    nightStartMin: nightStartMin,
    nightEndMin: nightEndMin,
    targetMonth: map[SETTING_KEYS.TARGET_MONTH]
  };
}

/**
 * 「集計対象年月」の値（Date型 または "2026-07" 等の文字列）を
 * { year, month(1-12) } の形に変換する。解釈できない場合は null。
 */
function parseTargetYearMonth_(value) {
  if (value === '' || value === null || value === undefined) return null;

  if (value instanceof Date) {
    return { year: value.getFullYear(), month: value.getMonth() + 1 };
  }

  if (typeof value === 'string' || typeof value === 'number') {
    const s = String(value).trim();
    const m = s.match(/^(\d{4})[\/\-年](\d{1,2})/);
    if (m) {
      const month = Number(m[2]);
      if (month >= 1 && month <= 12) {
        return { year: Number(m[1]), month: month };
      }
    }
  }
  return null;
}


// ============================================================================
// シート読み書き・エラー記録の共通処理
// ============================================================================

/**
 * ヘッダー行を除いたデータ行を2次元配列で返す。データが無い場合は空配列。
 */
function getDataRows_(sheet, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
}

/** 行のすべてのセルが空かどうかを判定する。 */
function isRowBlank_(row) {
  return row.every((v) => v === '' || v === null || v === undefined);
}

/** エラー一覧シートへの表示用に、値を短い文字列に整形する。 */
function formatForDetail_(value) {
  if (value === '' || value === null || value === undefined) return '(空欄)';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }
  return String(value);
}

/** 月次集計シートのデータ部分を書き換える（見出しは残す）。 */
function writeSummary_(sheet, rows) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, SUMMARY_HEADERS.length).clearContent();
    sheet.getRange(2, 1, lastRow - 1, SUMMARY_HEADERS.length).setBackground(null);
  }
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, SUMMARY_HEADERS.length).setValues(rows);
  }
}

/** エラー一覧シートの中身を消去する（見出しは残す）。集計を実行するたびに最新の状態にする。 */
function clearErrorSheet_(ss) {
  const sheet = ss.getSheetByName(SHEET_ERROR);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, ERROR_HEADERS.length).clearContent();
  }
}

/** エラー一覧シートに1行追加する。処理は止めずに続行するためのログ。 */
function logError_(ss, sheetName, rowNumber, message, detail) {
  const sheet = ss.getSheetByName(SHEET_ERROR);
  if (!sheet) return;
  sheet.appendRow([new Date(), sheetName, rowNumber, message, detail || '']);
}
