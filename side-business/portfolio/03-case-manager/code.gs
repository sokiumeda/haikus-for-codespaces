/**
 * 顧客・案件進捗管理ツール
 * ------------------------------------------------------------
 * 対象：税理士・社労士等の士業事務所／中小企業の営業担当者
 * 目的：案件の期限・進捗管理の「抜け漏れ」を防ぐ
 *
 * 使い方はスプレッドシート上部メニュー「案件管理ツール」から実行してください。
 * コードを直接編集する必要はありません。
 * ------------------------------------------------------------
 */

// ============================================================
// 定数定義（シート名・列番号・設定値などはすべてここに集約）
// ============================================================

// シート名
const SHEET_CUSTOMERS = '顧客マスタ';
const SHEET_CASES = '案件一覧';
const SHEET_DASHBOARD = 'ダッシュボード';
const SHEET_SETTINGS = '設定';
const SHEET_ERRORS = 'エラー一覧';

// 案件一覧の列番号（1始まり）
const CASE_COL = {
  ID: 1,           // 案件ID
  CUSTOMER_ID: 2,  // 顧客ID
  NAME: 3,         // 案件名
  TYPE: 4,         // 案件種別
  OWNER: 5,        // 担当者
  RECEIVED_DATE: 6,// 受任日
  DUE_DATE: 7,     // 期限日
  STATUS: 8,       // ステータス
  MEMO: 9          // 進捗メモ
};
const CASE_HEADER = ['案件ID', '顧客ID', '案件名', '案件種別', '担当者', '受任日', '期限日', 'ステータス', '進捗メモ'];
const CASE_LAST_COL = CASE_HEADER.length;

// 顧客マスタの列番号
const CUSTOMER_COL = {
  ID: 1,       // 顧客ID
  NAME: 2,     // 会社名
  STAFF: 3,    // 担当者名
  CONTACT: 4   // 連絡先
};
const CUSTOMER_HEADER = ['顧客ID', '会社名', '担当者名', '連絡先'];

// エラー一覧の列
const ERROR_HEADER = ['発生日時', '実行機能', '対象シート', '行番号', '項目', '入力されていた値', 'エラー内容'];

// 設定シートのセル位置
const SETTINGS_HEADER = ['設定項目', '値', '説明'];
const SETTINGS_ROW_THRESHOLD = 2; // 「期限間近日数」を書き込む行
const DEFAULT_THRESHOLD_DAYS = 7;

// 有効なステータス一覧
const STATUS_LIST = ['着手前', '対応中', '確認待ち', '完了'];
const STATUS_DONE = '完了';

// 色分け用の背景色
const COLOR_OVERDUE = '#f4c7c3';   // 期限超過：赤系
const COLOR_UPCOMING = '#fff2cc';  // 期限間近：黄系
const COLOR_NONE = null;           // 色クリア（既定色に戻す）


// ============================================================
// メニュー登録
// ============================================================

/**
 * スプレッドシートを開いたときに自動実行され、カスタムメニューを追加する。
 * 非エンジニアの利用者はこのメニューからすべての操作を行う。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('案件管理ツール')
    .addItem('① 初期セットアップ', 'initSheets')
    .addItem('② 期限チェック実行', 'checkDeadlines')
    .addItem('③ ダッシュボード更新', 'updateDashboard')
    .addToUi();
}


// ============================================================
// ① 初期セットアップ
// ============================================================

/**
 * 必要なシートが存在しない場合に自動作成する。
 * 既に存在するシートのデータは上書きしない（複数回実行しても安全）。
 */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  setupCustomerSheet_(ss);
  setupCaseSheet_(ss);
  setupSettingsSheet_(ss);
  setupErrorSheet_(ss);
  setupDashboardSheet_(ss);

  SpreadsheetApp.getUi().alert(
    '初期セットアップが完了しました。\n' +
    '「顧客マスタ」「案件一覧」に情報を入力してから、\n' +
    '「② 期限チェック実行」「③ ダッシュボード更新」をお使いください。'
  );
}

/** 顧客マスタシートを作成する（存在しない場合のみ） */
function setupCustomerSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_CUSTOMERS);
  if (sheet) return; // 既にある場合は何もしない（データ保護のため）

  sheet = ss.insertSheet(SHEET_CUSTOMERS);
  sheet.getRange(1, 1, 1, CUSTOMER_HEADER.length).setValues([CUSTOMER_HEADER]);
  formatHeaderRow_(sheet, CUSTOMER_HEADER.length);

  // 記入例（サンプルデータ）。実データを入力する際は削除・上書きしてください。
  sheet.getRange(2, 1, 1, CUSTOMER_HEADER.length).setValues([
    ['C001', '（記入例）サンプル商事株式会社', '山田 太郎', 'yamada@example.com']
  ]);
  sheet.setColumnWidths(1, CUSTOMER_HEADER.length, 160);
}

/** 案件一覧シートを作成する（存在しない場合のみ） */
function setupCaseSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_CASES);
  if (sheet) return;

  sheet = ss.insertSheet(SHEET_CASES);
  sheet.getRange(1, 1, 1, CASE_HEADER.length).setValues([CASE_HEADER]);
  formatHeaderRow_(sheet, CASE_HEADER.length);

  // 記入例（サンプルデータ）
  const today = new Date();
  const received = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
  const due = new Date(today.getTime() + 3 * 24 * 60 * 60 * 1000);
  sheet.getRange(2, 1, 1, CASE_HEADER.length).setValues([
    ['J001', 'C001', '（記入例）決算申告書類作成', '税務顧問', '佐藤', received, due, '対応中', '資料回収済み、申告書ドラフト作成中']
  ]);
  sheet.getRange(2, CASE_COL.RECEIVED_DATE).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(2, CASE_COL.DUE_DATE).setNumberFormat('yyyy/MM/dd');
  sheet.setColumnWidths(1, CASE_HEADER.length, 130);
  sheet.setColumnWidth(CASE_COL.NAME, 220);
  sheet.setColumnWidth(CASE_COL.MEMO, 260);

  // ステータス列にプルダウン（入力規則）を設定し、入力ミスを防ぐ
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_LIST, true)
    .setAllowInvalid(true) // 手入力での自由記述もエラー一覧で検知できるよう許容する
    .build();
  sheet.getRange(2, CASE_COL.STATUS, 998, 1).setDataValidation(statusRule);
}

/** 設定シートを作成する（存在しない場合のみ） */
function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (sheet) return;

  sheet = ss.insertSheet(SHEET_SETTINGS);
  sheet.getRange(1, 1, 1, SETTINGS_HEADER.length).setValues([SETTINGS_HEADER]);
  formatHeaderRow_(sheet, SETTINGS_HEADER.length);

  sheet.getRange(SETTINGS_ROW_THRESHOLD, 1, 1, 3).setValues([
    ['期限間近日数', DEFAULT_THRESHOLD_DAYS, '期限日までの残り日数がこの値以下になった案件を「期限間近」として黄色で表示します']
  ]);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 420);
}

/** エラー一覧シートを作成する（存在しない場合のみ） */
function setupErrorSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_ERRORS);
  if (sheet) return;

  sheet = ss.insertSheet(SHEET_ERRORS);
  sheet.getRange(1, 1, 1, ERROR_HEADER.length).setValues([ERROR_HEADER]);
  formatHeaderRow_(sheet, ERROR_HEADER.length);
  sheet.setColumnWidths(1, ERROR_HEADER.length, 160);
  sheet.setColumnWidth(7, 320);
}

/** ダッシュボードシートを作成する（存在しない場合のみ） */
function setupDashboardSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_DASHBOARD);
  if (sheet) return;

  sheet = ss.insertSheet(SHEET_DASHBOARD);
  sheet.getRange(1, 1).setValue('「③ ダッシュボード更新」を実行すると、ここに集計結果が表示されます。');
  // ダッシュボードはシートの並び順で一番左に置くと視認性が良いため先頭へ移動
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
}

/** ヘッダー行の共通書式（太字・背景色・固定表示） */
function formatHeaderRow_(sheet, numCols) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setFontWeight('bold');
  range.setBackground('#4a86e8');
  range.setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}


// ============================================================
// ② 期限チェック実行
// ============================================================

/**
 * 案件一覧を1行ずつ検証し、
 * ・期限超過 → 赤
 * ・期限間近 → 黄
 * ・完了     → 色クリア
 * に自動で色分けする。
 * 不正なデータ（未登録の顧客ID、日付でない期限日、不正なステータス等）は
 * 処理を止めずに「エラー一覧」シートに記録して次の行へ進む。
 */
function checkDeadlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const caseSheet = ss.getSheetByName(SHEET_CASES);
  const errorSheet = ss.getSheetByName(SHEET_ERRORS);

  if (!caseSheet || !errorSheet) {
    SpreadsheetApp.getUi().alert('「案件一覧」または「エラー一覧」シートが見つかりません。先に「① 初期セットアップ」を実行してください。');
    return;
  }

  const thresholdResult = getThresholdDays_(ss, '期限チェック実行');
  const { validRows, errors } = loadAndValidateCases_(ss, '期限チェック実行');
  if (thresholdResult.error) errors.push(thresholdResult.error);
  const thresholdDays = thresholdResult.value;

  writeErrorLog_(errorSheet, errors);

  const lastRow = caseSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('「案件一覧」にデータがありません。案件を入力してから再実行してください。');
    return;
  }

  const numDataRows = lastRow - 1;
  const colorMatrix = [];
  const today = truncateToDate_(new Date());

  // 案件一覧の全データ行分の色配列を用意（検証で除外された行・空行は色クリア）
  for (let i = 0; i < numDataRows; i++) {
    colorMatrix.push(new Array(CASE_LAST_COL).fill(COLOR_NONE));
  }

  validRows.forEach(function (row) {
    const idx = row.rowNumber - 2; // 配列上のインデックス（ヘッダー分を引く）
    let color = COLOR_NONE;

    if (row.status === STATUS_DONE) {
      color = COLOR_NONE; // 完了案件は色をクリア
    } else {
      const diffDays = daysBetween_(today, truncateToDate_(row.dueDate));
      if (diffDays < 0) {
        color = COLOR_OVERDUE;
      } else if (diffDays <= thresholdDays) {
        color = COLOR_UPCOMING;
      } else {
        color = COLOR_NONE;
      }
    }

    colorMatrix[idx] = new Array(CASE_LAST_COL).fill(color);
  });

  // まとめて1回で書き込むことで処理を高速化
  caseSheet.getRange(2, 1, numDataRows, CASE_LAST_COL).setBackgrounds(colorMatrix);

  const message = errors.length > 0
    ? '期限チェックが完了しました。\n色分けできなかった行が ' + errors.length + ' 件あります。「エラー一覧」シートを確認してください。'
    : '期限チェックが完了しました。エラーはありませんでした。';
  SpreadsheetApp.getUi().alert(message);
}


// ============================================================
// ③ ダッシュボード更新
// ============================================================

/**
 * 案件一覧を集計し、ダッシュボードシートに
 * ・ステータス別件数
 * ・担当者別の対応中件数／進捗率
 * ・期限超過案件リスト
 * ・期限間近案件リスト
 * をまとめて出力する。
 */
function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(SHEET_DASHBOARD);
  const errorSheet = ss.getSheetByName(SHEET_ERRORS);

  if (!dashboardSheet || !errorSheet) {
    SpreadsheetApp.getUi().alert('「ダッシュボード」または「エラー一覧」シートが見つかりません。先に「① 初期セットアップ」を実行してください。');
    return;
  }

  const thresholdResult = getThresholdDays_(ss, 'ダッシュボード更新');
  const { validRows, errors } = loadAndValidateCases_(ss, 'ダッシュボード更新');
  if (thresholdResult.error) errors.push(thresholdResult.error);
  const thresholdDays = thresholdResult.value;
  writeErrorLog_(errorSheet, errors);

  dashboardSheet.clear();
  dashboardSheet.setColumnWidths(1, 8, 130);
  dashboardSheet.setColumnWidth(2, 220);

  let cursorRow = 1;

  // --- タイトル・更新日時 ---
  dashboardSheet.getRange(cursorRow, 1).setValue('案件進捗ダッシュボード')
    .setFontWeight('bold').setFontSize(14);
  cursorRow++;
  dashboardSheet.getRange(cursorRow, 1).setValue(
    '更新日時：' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm')
  );
  cursorRow += 2;

  if (validRows.length === 0) {
    dashboardSheet.getRange(cursorRow, 1).setValue('集計対象の案件データがありません（「案件一覧」を入力してから再実行してください）。');
    SpreadsheetApp.getUi().alert('ダッシュボードを更新しました（対象データなし）。');
    return;
  }

  // --- ステータス別件数 ---
  cursorRow = writeSectionTitle_(dashboardSheet, cursorRow, 'ステータス別件数');
  const statusCounts = {};
  STATUS_LIST.forEach(function (s) { statusCounts[s] = 0; });
  let unknownStatusCount = 0;
  validRows.forEach(function (row) {
    if (Object.prototype.hasOwnProperty.call(statusCounts, row.status)) {
      statusCounts[row.status]++;
    } else {
      unknownStatusCount++;
    }
  });
  const statusTable = [['ステータス', '件数']];
  STATUS_LIST.forEach(function (s) { statusTable.push([s, statusCounts[s]]); });
  if (unknownStatusCount > 0) statusTable.push(['不明・その他', unknownStatusCount]);
  cursorRow = writeTable_(dashboardSheet, cursorRow, statusTable);
  cursorRow++;

  // --- 担当者別の対応中件数・進捗率 ---
  cursorRow = writeSectionTitle_(dashboardSheet, cursorRow, '担当者別の状況');
  const byOwner = {};
  validRows.forEach(function (row) {
    const owner = row.owner || '（担当者未設定）';
    if (!byOwner[owner]) {
      byOwner[owner] = { total: 0, inProgress: 0, done: 0 };
    }
    byOwner[owner].total++;
    if (row.status === '対応中') byOwner[owner].inProgress++;
    if (row.status === STATUS_DONE) byOwner[owner].done++;
  });
  const ownerTable = [['担当者', '対応中件数', '完了件数', '総件数', '進捗率（完了÷総件数）']];
  Object.keys(byOwner).sort().forEach(function (owner) {
    const o = byOwner[owner];
    const rate = o.total > 0 ? Math.round((o.done / o.total) * 100) + '%' : '-';
    ownerTable.push([owner, o.inProgress, o.done, o.total, rate]);
  });
  cursorRow = writeTable_(dashboardSheet, cursorRow, ownerTable);
  cursorRow++;

  // --- 期限超過／期限間近リスト ---
  const today = truncateToDate_(new Date());
  const overdueList = [];
  const upcomingList = [];
  validRows.forEach(function (row) {
    if (row.status === STATUS_DONE) return;
    const diffDays = daysBetween_(today, truncateToDate_(row.dueDate));
    if (diffDays < 0) {
      overdueList.push([row.id, row.name, row.customerId, row.owner, formatDate_(row.dueDate), (-diffDays) + '日超過']);
    } else if (diffDays <= thresholdDays) {
      upcomingList.push([row.id, row.name, row.customerId, row.owner, formatDate_(row.dueDate), '残り' + diffDays + '日']);
    }
  });

  cursorRow = writeSectionTitle_(dashboardSheet, cursorRow, '期限超過案件（要対応）');
  if (overdueList.length === 0) {
    dashboardSheet.getRange(cursorRow, 1).setValue('該当なし');
    cursorRow += 2;
  } else {
    const header = [['案件ID', '案件名', '顧客ID', '担当者', '期限日', '超過状況']];
    cursorRow = writeTable_(dashboardSheet, cursorRow, header.concat(overdueList), COLOR_OVERDUE);
    cursorRow++;
  }

  cursorRow = writeSectionTitle_(dashboardSheet, cursorRow, '期限間近案件（' + thresholdDays + '日以内）');
  if (upcomingList.length === 0) {
    dashboardSheet.getRange(cursorRow, 1).setValue('該当なし');
    cursorRow += 2;
  } else {
    const header = [['案件ID', '案件名', '顧客ID', '担当者', '期限日', '残り日数']];
    cursorRow = writeTable_(dashboardSheet, cursorRow, header.concat(upcomingList), COLOR_UPCOMING);
    cursorRow++;
  }

  const message = errors.length > 0
    ? 'ダッシュボードを更新しました。\n集計から除外したエラー行が ' + errors.length + ' 件あります。「エラー一覧」シートを確認してください。'
    : 'ダッシュボードを更新しました。';
  SpreadsheetApp.getUi().alert(message);
}

/** ダッシュボード用の見出しを1行書き込み、次の行番号を返す */
function writeSectionTitle_(sheet, row, title) {
  sheet.getRange(row, 1).setValue(title).setFontWeight('bold').setBackground('#d9e2f3');
  return row + 1;
}

/**
 * ダッシュボード用の表を書き込む（1行目は見出し行として太字・グレー背景）。
 * optionalRowColor を指定すると、見出し行以外の全行にその背景色を付ける
 * （期限超過／期限間近リストを一目で分かるようにするため）。
 * 書き込み後、次に使える行番号を返す。
 */
function writeTable_(sheet, startRow, table, optionalRowColor) {
  if (table.length === 0) return startRow;
  const numCols = table[0].length;
  const range = sheet.getRange(startRow, 1, table.length, numCols);
  range.setValues(table);
  sheet.getRange(startRow, 1, 1, numCols).setFontWeight('bold').setBackground('#f3f3f3');
  if (optionalRowColor && table.length > 1) {
    sheet.getRange(startRow + 1, 1, table.length - 1, numCols).setBackground(optionalRowColor);
  }
  return startRow + table.length;
}


// ============================================================
// 共通ロジック：データ読み込み・検証
// ============================================================

/**
 * 「案件一覧」を読み込み、行ごとに検証する。
 * ・完全な空行はスキップ（エラーにしない）
 * ・顧客IDが未入力／顧客マスタに存在しない → エラー
 * ・期限日が日付として解釈できない → エラー
 * ・ステータスが定義済みの4種類以外 → エラー
 * 検証をパスした行だけを validRows として返し、
 * 不備のある行は errors に理由付きで積んで返す（処理は止めない）。
 */
function loadAndValidateCases_(ss, executedFrom) {
  const caseSheet = ss.getSheetByName(SHEET_CASES);
  const customerSheet = ss.getSheetByName(SHEET_CUSTOMERS);
  const validRows = [];
  const errors = [];

  if (!caseSheet) {
    return { validRows: validRows, errors: errors };
  }

  // 顧客マスタのIDを集合として取得（存在チェック用）
  const customerIdSet = {};
  if (customerSheet && customerSheet.getLastRow() >= 2) {
    const custIds = customerSheet.getRange(2, CUSTOMER_COL.ID, customerSheet.getLastRow() - 1, 1).getValues();
    custIds.forEach(function (r) {
      const id = String(r[0]).trim();
      if (id !== '') customerIdSet[id] = true;
    });
  }

  const lastRow = caseSheet.getLastRow();
  if (lastRow < 2) {
    return { validRows: validRows, errors: errors }; // データ行なし（0件）でもエラーにしない
  }

  const numRows = lastRow - 1;
  const values = caseSheet.getRange(2, 1, numRows, CASE_LAST_COL).getValues();

  values.forEach(function (r, i) {
    const rowNumber = i + 2; // 実際のシート上の行番号
    const id = String(r[CASE_COL.ID - 1]).trim();
    const customerId = String(r[CASE_COL.CUSTOMER_ID - 1]).trim();
    const name = String(r[CASE_COL.NAME - 1]).trim();
    const caseType = String(r[CASE_COL.TYPE - 1]).trim();
    const owner = String(r[CASE_COL.OWNER - 1]).trim();
    const receivedDateRaw = r[CASE_COL.RECEIVED_DATE - 1];
    const dueDateRaw = r[CASE_COL.DUE_DATE - 1];
    const status = String(r[CASE_COL.STATUS - 1]).trim();
    const memo = String(r[CASE_COL.MEMO - 1]).trim();

    // 完全な空行はスキップ（エラーにしない）
    const isEmptyRow = [id, customerId, name, caseType, owner, status, memo].every(function (v) { return v === ''; })
      && !isValidDate_(receivedDateRaw) && !isValidDate_(dueDateRaw);
    if (isEmptyRow) return;

    let rowHasError = false;

    // 顧客IDの検証
    if (customerId === '') {
      errors.push([executedFrom, SHEET_CASES, rowNumber, '顧客ID', customerId, '顧客IDが未入力です']);
      rowHasError = true;
    } else if (!customerIdSet[customerId]) {
      errors.push([executedFrom, SHEET_CASES, rowNumber, '顧客ID', customerId, '顧客マスタに存在しない顧客IDです']);
      rowHasError = true;
    }

    // ステータスの検証
    if (STATUS_LIST.indexOf(status) === -1) {
      errors.push([executedFrom, SHEET_CASES, rowNumber, 'ステータス', status, '不正なステータス値です（着手前/対応中/確認待ち/完了のいずれかを入力してください）']);
      rowHasError = true;
    }

    // 期限日の検証
    const dueDate = isValidDate_(dueDateRaw) ? new Date(dueDateRaw) : null;
    if (!dueDate) {
      errors.push([executedFrom, SHEET_CASES, rowNumber, '期限日', dueDateRaw, '期限日が日付形式ではないか、空欄です']);
      rowHasError = true;
    }

    if (rowHasError) return; // この行は集計・色分けの対象から除外して次の行へ

    validRows.push({
      rowNumber: rowNumber,
      id: id,
      customerId: customerId,
      name: name,
      caseType: caseType,
      owner: owner,
      receivedDate: isValidDate_(receivedDateRaw) ? new Date(receivedDateRaw) : null,
      dueDate: dueDate,
      status: status,
      memo: memo
    });
  });

  return { validRows: validRows, errors: errors };
}

/** エラー一覧シートを最新の実行結果で上書きする（毎回リセットして書き直す） */
function writeErrorLog_(errorSheet, errors) {
  const lastRow = errorSheet.getLastRow();
  if (lastRow > 1) {
    errorSheet.getRange(2, 1, lastRow - 1, ERROR_HEADER.length).clearContent();
  }
  if (errors.length === 0) return;

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  const rows = errors.map(function (e) {
    // e = [実行機能, シート名, 行番号, 項目, 入力値, エラー内容]
    return [now].concat(e);
  });
  errorSheet.getRange(2, 1, rows.length, ERROR_HEADER.length).setValues(rows);
}

/**
 * 設定シートから「期限間近日数」を取得する。
 * 値が不正な場合は既定値にフォールバックし、その旨をエラー情報として一緒に返す
 * （呼び出し側で他のエラーとまとめて1回でエラー一覧に書き込むため、ここではシートに書き込まない）。
 */
function getThresholdDays_(ss, executedFrom) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return { value: DEFAULT_THRESHOLD_DAYS, error: null };

  const value = sheet.getRange(SETTINGS_ROW_THRESHOLD, 2).getValue();
  const num = Number(value);
  if (isFinite(num) && num >= 0) {
    return { value: num, error: null };
  }

  // 不正な設定値の場合は既定値にフォールバックし、エラー情報を返す
  const error = [executedFrom, SHEET_SETTINGS, SETTINGS_ROW_THRESHOLD, '期限間近日数', value, '数値として認識できないため既定値(' + DEFAULT_THRESHOLD_DAYS + '日)を使用しました'];
  return { value: DEFAULT_THRESHOLD_DAYS, error: error };
}


// ============================================================
// 共通ロジック：日付ユーティリティ
// ============================================================

/**
 * 値が「日付セル」として入力された正しい日付かどうかを判定する。
 * セルが日付形式で入力されていれば getValues() は Date型を返すため、
 * ここではDate型のみを有効とする（数値や文字列は誤入力とみなしエラーにする）。
 * これにより「123」のような数値や、日付として認識されない文字列の
 * 誤入力を「期限日が日付でない」として確実に検出できる。
 */
function isValidDate_(value) {
  if (Object.prototype.toString.call(value) !== '[object Date]') return false;
  return !isNaN(value.getTime());
}

/** 時刻情報を切り捨てて日付だけの比較ができるようにする */
function truncateToDate_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/** 2つの日付（時刻切り捨て済み）の差を「日数」で返す（to - from） */
function daysBetween_(from, to) {
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.round((to.getTime() - from.getTime()) / msPerDay);
}

/** 日付を yyyy/MM/dd 形式の文字列にする */
function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}
