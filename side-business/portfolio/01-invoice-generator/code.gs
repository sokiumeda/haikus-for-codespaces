/**
 * 見積書・請求書・納品書 自動発行ツール
 * ------------------------------------------------------------
 * 【このファイルについて】
 * ・非エンジニアの事務担当者が「カスタムメニュー」からボタン操作で使うことを前提に作っています。
 * ・処理を1か所止めても全体が止まらないよう、行ごとにエラーを捕まえて「エラー一覧」シートに記録します。
 * ・外部サービス通信は一切行いません（スプレッドシート・Drive内で完結します）。
 * ・改修する場合は、まずシート名や列名の定数（このすぐ下の設定エリア）を変更するだけで
 *   多くの調整ができるようにしてあります。
 * ------------------------------------------------------------
 */

/* ============================================================
 * 設定エリア（シート名・列構成・税率などをまとめて管理）
 * ============================================================ */

// シート名
const SHEET_CUSTOMER = '顧客マスタ';
const SHEET_ITEM = '品目マスタ';
const SHEET_INPUT = '発行入力';
const SHEET_HISTORY = '発行履歴';
const SHEET_ERROR = 'エラー一覧';
const SHEET_TEMPLATE = '帳票テンプレート';
const SHEET_SUMMARY = '集計';

// Drive上の保存先フォルダ名（スプレッドシートと同じ場所の直下に作成される）
const DRIVE_FOLDER_NAME = '帳票';

// 税率（品目マスタの「税区分」列で選択する値と対応）
const TAX_RATE_MAP = {
  '10%': 0.10,
  '軽減8%': 0.08,
};

// 発行入力シートの列（1始まり）
const INPUT_COL = {
  TYPE: 1,       // 発行区分（見積書/請求書/納品書）
  DATE: 2,       // 発行日
  CUSTOMER_ID: 3,// 顧客ID
  ITEM_ID: 4,    // 品目ID
  QTY: 5,        // 数量
  NOTE: 6,       // 備考
  STATUS: 7,     // ステータス（未発行/発行済）
  DOC_NO: 8,     // 帳票番号（発行時に自動記入。処理の重複判定にも使う）
};

const STATUS_UNISSUED = '未発行';
const STATUS_ISSUED = '発行済';

/* ============================================================
 * メニュー
 * ============================================================ */

// スプレッドシートを開いたときに自動実行され、カスタムメニューを追加する
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('帳票発行ツール')
    .addItem('① 初期セットアップ（シート自動作成）', 'initSheets')
    .addSeparator()
    .addItem('② 選択行の帳票PDF作成', 'issueSelectedRows')
    .addItem('③ 月次集計更新', 'updateMonthlySummary')
    .addToUi();
}

/* ============================================================
 * ① 初期セットアップ
 * ============================================================ */

/**
 * 必要なシートを一式自動作成する。
 * すでに同名シートがある場合は作り直さず、見出し行だけ確認・補完する（データを消さないため）。
 */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeaders(ss, SHEET_CUSTOMER,
    ['顧客ID', '会社名', '担当者名', '住所', 'メールアドレス', '振込先情報']);

  ensureSheetWithHeaders(ss, SHEET_ITEM,
    ['品目ID', '品目名', '単価', '税区分（10% or 軽減8%）']);

  ensureSheetWithHeaders(ss, SHEET_INPUT,
    ['発行区分（見積書/請求書/納品書）', '発行日', '顧客ID', '品目ID', '数量', '備考', 'ステータス', '帳票番号']);

  ensureSheetWithHeaders(ss, SHEET_HISTORY,
    ['発行日時', '帳票番号', '区分', '顧客名', '金額（税込）', 'PDFのDriveリンク']);

  ensureSheetWithHeaders(ss, SHEET_ERROR,
    ['検出日時', '対象シート', '行番号', 'エラー内容', '対処のヒント']);

  ensureSheetWithHeaders(ss, SHEET_SUMMARY,
    ['年月', '区分', '合計金額（税込）']);

  ensureTemplateSheet(ss);

  // Drive保存先フォルダも作っておく
  getOrCreateOutputFolder();

  SpreadsheetApp.getUi().alert(
    '初期セットアップが完了しました。\n' +
    '「顧客マスタ」「品目マスタ」にデータを入力し、「発行入力」に発行したい行を追加してください。'
  );
}

// シートが無ければ作成し、見出し行（1行目）が無ければ設定する共通処理
function ensureSheetWithHeaders(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isEmpty = firstRow.every((v) => v === '' || v === null);
  if (isEmpty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 帳票の下書きとして使う「帳票テンプレート」シートを作る（値を差し込んでからPDF化する台紙）
function ensureTemplateSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_TEMPLATE);
  if (sheet) return sheet; // すでにあれば手を加えない（レイアウトを崩さないため）

  sheet = ss.insertSheet(SHEET_TEMPLATE);

  // シンプルな帳票レイアウトを組む。セル位置はfillTemplateAndConvertToPdf()内の定数と対応させている。
  sheet.getRange('A1').setValue('【帳票タイトル】').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A3').setValue('帳票番号：');
  sheet.getRange('B3').setValue('');
  sheet.getRange('A4').setValue('発行日：');
  sheet.getRange('B4').setValue('');

  sheet.getRange('A6').setValue('宛先：');
  sheet.getRange('B6').setValue('').setFontSize(14);
  sheet.getRange('A7').setValue('ご担当者：');
  sheet.getRange('B7').setValue('');
  sheet.getRange('A8').setValue('住所：');
  sheet.getRange('B8').setValue('');

  sheet.getRange('A10').setValue('発行元：本ツール利用事業者');

  // 明細ヘッダー
  sheet.getRange('A12:E12').setValues([['品目', '単価', '数量', '税区分', '金額']]);
  sheet.getRange('A12:E12').setFontWeight('bold');

  // 明細行はA13〜A22の10行分を確保（fillTemplateAndConvertToPdf側でクリア・再記入する）
  sheet.getRange('A23').setValue('小計：');
  sheet.getRange('A24').setValue('消費税：');
  sheet.getRange('A25').setValue('合計：');

  sheet.getRange('A27').setValue('振込先：');
  sheet.getRange('B27').setValue('');
  sheet.getRange('A29').setValue('備考：');
  sheet.getRange('B29').setValue('');

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidths(2, 4, 110);

  return sheet;
}

// 保存先Driveフォルダを取得（無ければ作成）。スプレッドシートと同じ階層に作る。
function getOrCreateOutputFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const existing = parentFolder.getFoldersByName(DRIVE_FOLDER_NAME);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parentFolder.createFolder(DRIVE_FOLDER_NAME);
}

/* ============================================================
 * ② 選択行の帳票PDF作成
 * ============================================================ */

/**
 * 「発行入力」シートで選択中の行（複数選択可）を対象に、
 * 未発行かつ入力内容に問題のない行だけPDFを発行する。
 * 問題のある行はスキップし、「エラー一覧」に理由を記録する。
 */
function issueSelectedRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SHEET_INPUT);
  const ui = SpreadsheetApp.getUi();

  if (!inputSheet) {
    ui.alert('「発行入力」シートが見つかりません。先に「① 初期セットアップ」を実行してください。');
    return;
  }

  const lastRow = inputSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('「発行入力」シートにデータがありません。行を追加してから実行してください。');
    return;
  }

  // 選択範囲から対象行番号を取得する。選択が無い/1行のみの場合も考慮する。
  const activeRange = inputSheet.getActiveRange();
  let targetRows = [];
  if (activeRange && activeRange.getSheet().getName() === SHEET_INPUT) {
    const startRow = Math.max(activeRange.getRow(), 2); // 見出し行は除外
    const endRow = activeRange.getLastRow();
    for (let r = startRow; r <= endRow; r++) {
      targetRows.push(r);
    }
  }
  if (targetRows.length === 0) {
    ui.alert('「発行入力」シート上で、発行したい行を選択してから実行してください。');
    return;
  }

  // マスタをまとめて読み込んでおく（毎行アクセスすると遅いため）
  const customerMap = loadCustomerMap(ss);
  const itemMap = loadItemMap(ss);

  const inputValues = inputSheet.getRange(1, 1, lastRow, 8).getValues();
  const folder = getOrCreateOutputFolder();

  let issuedCount = 0;
  let errorCount = 0;
  let skippedAlreadyIssued = 0;

  targetRows.forEach((rowNum) => {
    const rowIndex = rowNum - 1; // inputValues配列上のインデックス（0始まり、1行目が見出し）
    const row = inputValues[rowIndex];
    if (!row) return;

    // 完全な空行はスキップ（エラーとして記録しない）
    const isBlank = row.every((v) => v === '' || v === null);
    if (isBlank) return;

    const status = row[INPUT_COL.STATUS - 1];
    if (status === STATUS_ISSUED) {
      skippedAlreadyIssued++;
      return; // すでに発行済みの行は二重発行しない
    }

    try {
      const result = issueOneRow(ss, inputSheet, rowNum, row, customerMap, itemMap, folder);
      if (result.ok) {
        issuedCount++;
      } else {
        recordError(ss, SHEET_INPUT, rowNum, result.message, result.hint);
        errorCount++;
      }
    } catch (e) {
      // 想定外のエラーも必ずここで受け止め、処理全体を止めない
      recordError(ss, SHEET_INPUT, rowNum, '想定外のエラー: ' + e.message, '入力内容を確認し、再実行してください。改善しない場合は開発担当に連絡してください。');
      errorCount++;
    }
  });

  let msg = `発行完了：${issuedCount}件\nエラー：${errorCount}件（詳細は「エラー一覧」シートを確認してください）`;
  if (skippedAlreadyIssued > 0) {
    msg += `\n発行済みのためスキップ：${skippedAlreadyIssued}件`;
  }
  ui.alert(msg);
}

// 顧客マスタをMapに読み込む（キー：顧客ID）
function loadCustomerMap(ss) {
  const sheet = ss.getSheetByName(SHEET_CUSTOMER);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  values.forEach((row) => {
    const id = String(row[0]).trim();
    if (id === '') return;
    map[id] = {
      id: id,
      companyName: row[1],
      contactName: row[2],
      address: row[3],
      email: row[4],
      bankInfo: row[5],
    };
  });
  return map;
}

// 品目マスタをMapに読み込む（キー：品目ID）
function loadItemMap(ss) {
  const sheet = ss.getSheetByName(SHEET_ITEM);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  values.forEach((row) => {
    const id = String(row[0]).trim();
    if (id === '') return;
    map[id] = {
      id: id,
      name: row[1],
      unitPrice: row[2],
      taxCategory: String(row[3]).trim(),
    };
  });
  return map;
}

/**
 * 発行入力の1行を検証し、問題なければPDFを発行して「発行入力」「発行履歴」に反映する。
 * 戻り値：{ ok: true } または { ok: false, message, hint }
 */
function issueOneRow(ss, inputSheet, rowNum, row, customerMap, itemMap, folder) {
  const type = String(row[INPUT_COL.TYPE - 1]).trim();
  const dateValue = row[INPUT_COL.DATE - 1];
  const customerId = String(row[INPUT_COL.CUSTOMER_ID - 1]).trim();
  const itemId = String(row[INPUT_COL.ITEM_ID - 1]).trim();
  const qty = row[INPUT_COL.QTY - 1];
  const note = row[INPUT_COL.NOTE - 1] || '';

  // --- 入力チェック（1つずつ理由を分かりやすく返す） ---
  if (['見積書', '請求書', '納品書'].indexOf(type) === -1) {
    return { ok: false, message: `発行区分「${type}」が不正です。`, hint: '「見積書」「請求書」「納品書」のいずれかを入力してください。' };
  }
  if (!dateValue) {
    return { ok: false, message: '発行日が空です。', hint: '発行日を入力してください。' };
  }
  if (customerId === '') {
    return { ok: false, message: '顧客IDが空です。', hint: '顧客マスタに登録済みの顧客IDを入力してください。' };
  }
  const customer = customerMap[customerId];
  if (!customer) {
    return { ok: false, message: `顧客ID「${customerId}」が顧客マスタに見つかりません。`, hint: '顧客マスタの顧客ID列と一致しているか確認してください（全角/半角の違いにも注意）。' };
  }
  if (itemId === '') {
    return { ok: false, message: '品目IDが空です。', hint: '品目マスタに登録済みの品目IDを入力してください。' };
  }
  const item = itemMap[itemId];
  if (!item) {
    return { ok: false, message: `品目ID「${itemId}」が品目マスタに見つかりません。`, hint: '品目マスタの品目ID列と一致しているか確認してください。' };
  }
  if (qty === '' || qty === null || typeof qty !== 'number' || isNaN(qty)) {
    return { ok: false, message: '数量が空、または数値ではありません。', hint: '数量に半角数字を入力してください。' };
  }
  if (qty <= 0) {
    return { ok: false, message: `数量が不正です（入力値：${qty}）。`, hint: '数量には1以上の数値を入力してください。' };
  }
  const taxRate = TAX_RATE_MAP[item.taxCategory];
  if (taxRate === undefined) {
    return { ok: false, message: `品目「${item.name}」の税区分「${item.taxCategory}」が不正です。`, hint: '品目マスタの税区分を「10%」または「軽減8%」に修正してください。' };
  }
  if (typeof item.unitPrice !== 'number' || isNaN(item.unitPrice) || item.unitPrice < 0) {
    return { ok: false, message: `品目「${item.name}」の単価が不正です。`, hint: '品目マスタの単価に0以上の数値を入力してください。' };
  }

  // --- 金額計算 ---
  const subtotal = item.unitPrice * qty;
  const tax = Math.floor(subtotal * taxRate); // 端数は切り捨て（貼付先の運用に合わせて変更可）
  const total = subtotal + tax;

  // --- 帳票番号の採番（INV-YYYYMM-連番） ---
  const issueDate = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
  if (isNaN(issueDate.getTime())) {
    return { ok: false, message: '発行日の形式が不正です。', hint: '発行日は日付として認識できる形式（例：2026/07/08）で入力してください。' };
  }
  const docNo = generateDocNumber(ss, issueDate);

  // --- テンプレートに差し込んでPDF化 ---
  const pdfUrl = fillTemplateAndConvertToPdf(ss, folder, {
    type: type,
    docNo: docNo,
    issueDate: issueDate,
    customer: customer,
    item: item,
    qty: qty,
    subtotal: subtotal,
    tax: tax,
    total: total,
    note: note,
  });

  // --- 「発行入力」を更新 ---
  inputSheet.getRange(rowNum, INPUT_COL.STATUS).setValue(STATUS_ISSUED);
  inputSheet.getRange(rowNum, INPUT_COL.DOC_NO).setValue(docNo);

  // --- 「発行履歴」に記録 ---
  appendHistory(ss, {
    issuedAt: new Date(),
    docNo: docNo,
    type: type,
    customerName: customer.companyName,
    total: total,
    pdfUrl: pdfUrl,
  });

  return { ok: true };
}

// 帳票番号を「INV-YYYYMM-連番」形式で発行する。連番は発行履歴シートを見て、その月の最大値+1にする。
function generateDocNumber(ss, issueDate) {
  const yyyymm = Utilities.formatDate(issueDate, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMM');
  const prefix = `INV-${yyyymm}-`;

  const historySheet = ss.getSheetByName(SHEET_HISTORY);
  let maxSeq = 0;
  if (historySheet && historySheet.getLastRow() >= 2) {
    const docNumbers = historySheet.getRange(2, 2, historySheet.getLastRow() - 1, 1).getValues();
    docNumbers.forEach((r) => {
      const v = String(r[0]);
      if (v.indexOf(prefix) === 0) {
        const seq = parseInt(v.substring(prefix.length), 10);
        if (!isNaN(seq) && seq > maxSeq) {
          maxSeq = seq;
        }
      }
    });
  }
  const nextSeq = maxSeq + 1;
  const seqStr = ('000' + nextSeq).slice(-3); // 3桁ゼロ埋め（4桁目以降はそのまま伸びる）
  return prefix + seqStr;
}

/**
 * 「帳票テンプレート」シートに値を差し込み、PDFに変換してDriveに保存する。
 * 戻り値：保存したPDFのDrive URL
 */
function fillTemplateAndConvertToPdf(ss, folder, data) {
  const sheet = ss.getSheetByName(SHEET_TEMPLATE);
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';

  const titleMap = { '見積書': '御見積書', '請求書': 'ご請求書', '納品書': '納品書' };
  sheet.getRange('A1').setValue('【' + (titleMap[data.type] || data.type) + '】');

  sheet.getRange('B3').setValue(data.docNo);
  sheet.getRange('B4').setValue(Utilities.formatDate(data.issueDate, tz, 'yyyy/MM/dd'));

  sheet.getRange('B6').setValue(data.customer.companyName + ' 御中');
  sheet.getRange('B7').setValue(data.customer.contactName);
  sheet.getRange('B8').setValue(data.customer.address);

  // 明細行（1品目のみだが、将来複数品目対応する場合はここをループに変更する）
  sheet.getRange('A13:E22').clearContent();
  sheet.getRange('A13:E13').setValues([[
    data.item.name,
    data.item.unitPrice,
    data.qty,
    data.item.taxCategory,
    data.subtotal,
  ]]);

  sheet.getRange('B23').setValue(data.subtotal);
  sheet.getRange('B24').setValue(data.tax);
  sheet.getRange('B25').setValue(data.total);

  sheet.getRange('B27').setValue(data.customer.bankInfo);
  sheet.getRange('B29').setValue(data.note);

  SpreadsheetApp.flush(); // 差し込んだ値を確実にシートへ反映させてからPDF化する

  const pdfBlob = exportSheetAsPdf(sheet);
  const fileName = `${data.docNo}_${data.type}_${data.customer.companyName}.pdf`;
  const pdfFile = folder.createFile(pdfBlob).setName(fileName);

  return pdfFile.getUrl();
}

/**
 * 指定シートだけをPDFに変換する。
 * ※外部サービス通信（UrlFetchApp等）は使わず、SpreadsheetApp / DriveApp の標準機能だけで実現している。
 *   手順：①対象シートだけをコピーした一時スプレッドシートを作る → ②それをPDF化 → ③一時ファイルを削除する
 */
function exportSheetAsPdf(sheet) {
  // ①対象シートのコピーを持つ、一時的なスプレッドシートを新規作成する
  const tempSs = SpreadsheetApp.create('_tmp_帳票出力_' + new Date().getTime());
  try {
    const copiedSheet = sheet.copyTo(tempSs);

    // 新規作成時に自動でできる「シート1」など、コピーしたシート以外は削除する（PDFに余計なページが出ないようにするため）
    tempSs.getSheets().forEach((s) => {
      if (s.getSheetId() !== copiedSheet.getSheetId()) {
        tempSs.deleteSheet(s);
      }
    });
    SpreadsheetApp.flush();

    // ②DriveApp標準の変換機能でPDF化する
    const pdfBlob = DriveApp.getFileById(tempSs.getId()).getAs(MimeType.PDF);
    return pdfBlob;
  } finally {
    // ③一時スプレッドシートはゴミ箱に移動して片付ける（Drive内に残り続けないようにする）
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
  }
}

// 「発行履歴」に1行追記する
function appendHistory(ss, entry) {
  const sheet = ss.getSheetByName(SHEET_HISTORY);
  sheet.appendRow([
    entry.issuedAt,
    entry.docNo,
    entry.type,
    entry.customerName,
    entry.total,
    entry.pdfUrl,
  ]);
}

// 「エラー一覧」に1行記録する
function recordError(ss, sheetName, rowNum, message, hint) {
  const sheet = ss.getSheetByName(SHEET_ERROR);
  if (!sheet) return; // 万一シートが無い場合も処理は止めない
  sheet.appendRow([new Date(), sheetName, rowNum, message, hint || '']);
}

/* ============================================================
 * ③ 月次集計更新
 * ============================================================ */

/**
 * 「発行履歴」を集計し、「年月×区分」の合計金額を「集計」シートに出力する。
 * 発行履歴が0件でもエラーにせず、見出しだけの状態で終える。
 */
function updateMonthlySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(SHEET_HISTORY);
  const summarySheet = ensureSheetWithHeaders(ss, SHEET_SUMMARY, ['年月', '区分', '合計金額（税込）']);

  // 既存の集計データ行をクリア（見出しは残す）
  if (summarySheet.getLastRow() > 1) {
    summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 3).clearContent();
  }

  if (!historySheet || historySheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('「発行履歴」にデータがまだありません。帳票を発行すると集計されます。');
    return;
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const values = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 5).getValues();

  // キー「年月_区分」で合計する
  const totals = {};
  values.forEach((row) => {
    const issuedAt = row[0];
    const type = row[2];
    const amount = row[4];
    if (!(issuedAt instanceof Date) || isNaN(issuedAt.getTime())) return; // 不正な行はスキップ
    if (typeof amount !== 'number' || isNaN(amount)) return;

    const yyyymm = Utilities.formatDate(issuedAt, tz, 'yyyy-MM');
    const key = yyyymm + '_' + type;
    if (!totals[key]) {
      totals[key] = { yyyymm: yyyymm, type: type, total: 0 };
    }
    totals[key].total += amount;
  });

  const rows = Object.keys(totals)
    .map((k) => totals[k])
    .sort((a, b) => (a.yyyymm + a.type).localeCompare(b.yyyymm + b.type))
    .map((t) => [t.yyyymm, t.type, t.total]);

  if (rows.length > 0) {
    summarySheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  SpreadsheetApp.getUi().alert(`月次集計を更新しました（${rows.length}件の年月×区分の組み合わせ）。`);
}
