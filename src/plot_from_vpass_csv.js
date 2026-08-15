/**
 * Vpass利用明細CSVを自動仕訳して円グラフ化するスクリプト(プロトタイプ)
 *
 * 【事前準備】
 * 1. このコードをGoogleスプレッドシートの「拡張機能 > Apps Script」のエディタに貼り付ける
 * 2. FOLDER_ID に、月次CSVが入っているGoogle DriveフォルダのIDを設定する
 *    (フォルダを開いたときのURL末尾: https://drive.google.com/drive/folders/【ここ】)
 * 3. CATEGORY_DICT に、店舗名(部分一致)と分類の対応をどんどん追加していく
 * 4. classifyExpenses を実行(初回はGoogleアカウントの権限承認が必要)
 * 5. 実行後、スプレッドシートに「明細」「集計」「推移」シートが作成される
 *    - 「集計」シート: CSV(ファイル)ごと・全ファイル合計の分類別円グラフ
 *    - 「推移」シート: ファイルごとの分類別支出推移の積み上げ棒グラフ
 */

// ===== 設定 =====

// CSVが保存されているDriveフォルダのID
const FOLDER_ID = 'フォルダidをここに入力';

// CSVの文字コード(Vpassのダウンロードcsvは Shift_JIS のことが多い。文字化けする場合はUTF-8に変更)
const CSV_ENCODING = 'Shift_JIS';

const DETAIL_SHEET_NAME = '明細';
const SUMMARY_SHEET_NAME = '集計';
const TRANSITION_SHEET_NAME = '推移';

// 「集計」シートで、ファイル1つ分の表+円グラフに割り当てる行数(ブロック間の間隔)
const SUMMARY_ROWS_PER_BLOCK = 25;

// 店舗名(部分一致)→ 分類 の対応表。上から順にチェックし、最初にマッチしたものを採用
// 分類は以下の6種類固定: スーパー・コンビニ・自販機 / 外食 / 交通費 / 光熱費 / 娯楽・嗜好品 / 未分類
// ※クレジットカードCSVは英数字が全角(ＡＢＣ等)で入っていることが多いため、
//   半角・全角・カタカナ表記ゆれをそれぞれキーとして用意している
const CATEGORY_DICT = {
  // ---- スーパー・コンビニ・自販機 ----
  'イオン': 'スーパー・コンビニ・自販機',
  'セブン': 'スーパー・コンビニ・自販機',
  'ファミリーマート': 'スーパー・コンビニ・自販機',
  'ローソン': 'スーパー・コンビニ・自販機',
  'ミニストップ': 'スーパー・コンビニ・自販機',
  '西友': 'スーパー・コンビニ・自販機',
  'ライフ': 'スーパー・コンビニ・自販機',
  'いなげや': 'スーパー・コンビニ・自販機',
  '東急ストア': 'スーパー・コンビニ・自販機',
  'ハックドラッグ': 'スーパー・コンビニ・自販機',
  '飲料自販機': 'スーパー・コンビニ・自販機',
  'コカコ': 'スーパー・コンビニ・自販機', // コカコーラボトラーズ(自販機)
  'Ｆｉｔ　Ｃａｒｅ　ＤＥＰＯＴ': 'スーパー・コンビニ・自販機',
  'アピタ': 'スーパー・コンビニ・自販機',
  '大野屋': 'スーパー・コンビニ・自販機',

  // ---- 外食 ----
  'マクドナルド': '外食',
  'スターバックス': '外食',
  'ドトール': '外食',
  '吉野家': '外食',
  'すき家': '外食',
  'サイゼリヤ': '外食',
  'ガスト': '外食',
  '丸亀製麺': '外食',
  '味奈登庵': '外食',
  'ブドウヤ': '外食',
  'ウーバー': '外食', // ウーバーイーツ
  'UBER': '外食',    // UBER ONE MEMBERSHIP など半角表記
  'ＣＲＩＳＰ': '外食', // CRISP SALAD WORKS(全角表記)
  'クリスプ': '外食',   // クリスプサラダワークス(カタカナ表記)
  '上星商店': '外食',
  'いちごの樹': '外食',
  'キッチンオリジン': '外食',

  // ---- 交通費 ----
  'JR': '交通費',
  'ENEOS': '交通費',
  'ＥＮＥＯＳ': '交通費',
  '出光': '交通費',
  'Suica': '交通費',
  'PASMO': '交通費',
  'タイムズ': '交通費',
  'ETC': '交通費',
  'ＥＴＣ': '交通費', // 全角表記
  'キグナス': '交通費', // ガソリンスタンド
  '中日本高速道路': '交通費',

  // ---- 光熱費 ----
  '東京電力': '光熱費',
  '東京ガス': '光熱費',
  '東京都水道局': '光熱費',
  '水道': '光熱費', // 「〇〇市水道料金」等
  'ガスリョウキン': '光熱費', // 「〇〇　ガスリョウキン」表記
  'Ｕ−ＰＯＷＥＲ': '光熱費', // 電力会社

  // ---- 娯楽・嗜好品 ----
  'Amazon': '娯楽・嗜好品',
  'ＡＭＡＺＯＮ': '娯楽・嗜好品', // 全角表記
  'ＡＰＰＬＥ': '娯楽・嗜好品',   // App Store/iCloud等のサブスク
  'TSUTAYA': '娯楽・嗜好品',
  'ゲオ': '娯楽・嗜好品',
  'ヨドバシカメラ': '娯楽・嗜好品',
  'ルミネ': '娯楽・嗜好品',
  '伊勢丹': '娯楽・嗜好品',
  'キュービックプラザ': '娯楽・嗜好品',
  'キュ−ビックプラザ': '娯楽・嗜好品', // ダッシュ記号違いの表記ゆれ
  'ノースポート': '娯楽・嗜好品',
  'トレッサ': '娯楽・嗜好品',
  'ランドマーク': '娯楽・嗜好品',
  'サクラステージ': '娯楽・嗜好品',
  'ブルク': '娯楽・嗜好品', // 映画館
  'セリア': '娯楽・嗜好品',
  'かたぎり塾': '娯楽・嗜好品', // パーソナルトレーニングジム
  'ホテル': '娯楽・嗜好品',

  // 必要に応じて追加していく
};

const UNCLASSIFIED_LABEL = '未分類';

// 積み上げ棒グラフ・推移表で使う分類の並び順(固定)
const CATEGORY_ORDER = ['スーパー・コンビニ・自販機', '外食', '交通費', '光熱費', '娯楽・嗜好品', UNCLASSIFIED_LABEL];

// ===== メイン処理 =====

function classifyExpenses() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = getCsvFilesInFolder(folder);

  const transactions = []; // {fileName, date, store, amount, category}

  files.forEach(file => {
    const rows = readCsvRows(file);

    rows.forEach(row => {
      const rawDate = (row[0] || '').toString().trim();
      if (!isDateString(rawDate)) return; // 見出し行・合計行などはスキップ
      const dateStr = toHalfWidth(rawDate);

      const store = (row[1] || '').toString().trim();
      const amount = parseAmount(row[2]);
      if (!store || isNaN(amount)) return;

      transactions.push({
        fileName: file.getName(),
        date: dateStr,
        store: store,
        amount: amount,
        category: classifyStore(store),
      });
    });
  });

  writeDetailSheet(transactions);
  const summary = writeSummarySheet(transactions, files);
  drawPieCharts(summary);

  const transitionSummary = writeTransitionSheet(transactions, files);
  drawTransitionChart(transitionSummary.sheet, transitionSummary.transition);
}

// ===== CSV読み込み =====

function readCsvRows(file) {
  let csvText;
  try {
    csvText = file.getBlob().getDataAsString(CSV_ENCODING);
  } catch (e) {
    csvText = file.getBlob().getDataAsString('UTF-8');
  }
  return Utilities.parseCsv(csvText);
}

function getCsvFilesInFolder(folder) {
  // MIMEタイプでの判定は環境によって text/csv にならないことがあるため、
  // ファイル名の拡張子(.csv)で判定する
  const allFiles = folder.getFiles();
  const csvFiles = [];
  while (allFiles.hasNext()) {
    const file = allFiles.next();
    if (/\.csv$/i.test(file.getName())) {
      csvFiles.push(file);
    }
  }
  // ファイル名順(例: 202605.csv, 202606.csv)に並べ、集計・グラフの並び順を安定させる
  csvFiles.sort((a, b) => a.getName().localeCompare(b.getName(), 'ja'));
  return csvFiles;
}

function toHalfWidth(str) {
  // 全角の数字・記号を半角に変換する(Vpassのcsvは全角表記が混ざっていることがあるため)
  const zenkakuMap = {
    '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
    '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
    '／': '/', '－': '-', '．': '.', '，': ',', '￥': '¥',
  };
  return str.replace(/[０-９／－．，￥]/g, ch => zenkakuMap[ch] || ch);
}

function isDateString(str) {
  // yyyy/mm/dd, yy/mm/dd, yyyy-mm-dd など(全角数字・2桁年も許容)のみレコードとして扱う
  const normalized = toHalfWidth(str);
  return /^\d{2,4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(normalized);
}

function parseAmount(value) {
  if (value === null || value === undefined) return NaN;
  const normalized = toHalfWidth(value.toString());
  const cleaned = normalized.replace(/[,¥\s]/g, '');
  return Number(cleaned);
}

// ===== 分類ロジック =====

function classifyStore(store) {
  for (const keyword in CATEGORY_DICT) {
    if (store.indexOf(keyword) !== -1) {
      return CATEGORY_DICT[keyword];
    }
  }
  return UNCLASSIFIED_LABEL;
}

// ===== シート書き出し =====

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  sheet.clear();
  return sheet;
}

function writeDetailSheet(transactions) {
  const sheet = getOrCreateSheet(DETAIL_SHEET_NAME);
  sheet.appendRow(['ファイル名', '日付', '店舗', '金額', '分類']);
  const rows = transactions.map(t => [t.fileName, t.date, t.store, t.amount, t.category]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

function writeSummarySheet(transactions, files) {
  const sheet = getOrCreateSheet(SUMMARY_SHEET_NAME);
  const blocks = [];

  // ファイルごとに分類・合計金額のブロックを積み上げて書き出す
  files.forEach((file, i) => {
    const fileName = file.getName();
    const fileTransactions = transactions.filter(t => t.fileName === fileName);
    blocks.push(writeSummaryBlock(sheet, i * SUMMARY_ROWS_PER_BLOCK + 1, fileName, fileTransactions));
  });

  // 最後に全ファイル合計のブロックを追加
  const totalStartRow = files.length * SUMMARY_ROWS_PER_BLOCK + 1;
  blocks.push(writeSummaryBlock(sheet, totalStartRow, '全ファイル合計', transactions));

  return { sheet, blocks };
}

// 「推移」シート: 円グラフの集計(「集計」シート)とは別タブに、
// ファイルごとの分類別推移の表+積み上げ棒グラフを書き出す
function writeTransitionSheet(transactions, files) {
  const sheet = getOrCreateSheet(TRANSITION_SHEET_NAME);
  const transition = writeTransitionSection(sheet, 1, files, transactions);
  return { sheet, transition };
}

function writeSummaryBlock(sheet, startRow, label, transactions) {
  const totals = {};
  transactions.forEach(t => {
    totals[t.category] = (totals[t.category] || 0) + t.amount;
  });

  sheet.getRange(startRow, 1).setValue(label).setFontWeight('bold');
  sheet.getRange(startRow + 1, 1, 1, 2).setValues([['分類', '合計金額']]);

  const rows = Object.keys(totals).map(category => [category, totals[category]]);
  if (rows.length > 0) {
    sheet.getRange(startRow + 2, 1, rows.length, 2).setValues(rows);
  }

  return {
    label: label,
    headerRow: startRow + 1, // ['分類','合計金額']の見出し行
    rowCount: rows.length,
    chartAnchorRow: startRow + 2 + rows.length + 1,
  };
}

function writeTransitionSection(sheet, startRow, files, transactions) {
  sheet.getRange(startRow, 1).setValue('ファイルごとの推移').setFontWeight('bold');

  const headerRow = startRow + 1;
  const columnCount = CATEGORY_ORDER.length + 1; // ファイル名 + 各分類
  sheet.getRange(headerRow, 1, 1, columnCount).setValues([['ファイル名'].concat(CATEGORY_ORDER)]);

  // ファイルごとに、分類の並び順(CATEGORY_ORDER)に沿って合計金額を並べる(0円の分類も含める)
  const dataRows = files.map(file => {
    const fileName = file.getName();
    const totals = {};
    transactions
      .filter(t => t.fileName === fileName)
      .forEach(t => {
        totals[t.category] = (totals[t.category] || 0) + t.amount;
      });
    return [fileName].concat(CATEGORY_ORDER.map(category => totals[category] || 0));
  });

  if (dataRows.length > 0) {
    sheet.getRange(headerRow + 1, 1, dataRows.length, columnCount).setValues(dataRows);
  }

  return {
    headerRow: headerRow,
    columnCount: columnCount,
    rowCount: dataRows.length,
    chartAnchorRow: headerRow + dataRows.length + 2,
  };
}

// ===== 円グラフ =====

function drawPieCharts(summary) {
  const { sheet, blocks } = summary;

  // 既存のグラフを削除してから作り直す(再実行しても増殖しないように)
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));

  blocks.forEach(block => {
    if (block.rowCount === 0) return; // データが無いファイルはグラフを描かない

    const dataRange = sheet.getRange(block.headerRow, 1, block.rowCount + 1, 2);
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dataRange)
      .setPosition(block.chartAnchorRow, 1, 0, 0)
      .setOption('title', block.label + ' の支出分類')
      .setOption('width', 380)
      .setOption('height', 260)
      .build();

    sheet.insertChart(chart);
  });
}

// ===== 積み上げ棒グラフ(推移) =====

function drawTransitionChart(sheet, transition) {
  // 既存のグラフを削除してから作り直す(再実行しても増殖しないように)
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));

  if (transition.rowCount === 0) return; // ファイルが無ければ何もしない

  const dataRange = sheet.getRange(transition.headerRow, 1, transition.rowCount + 1, transition.columnCount);
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    // 1行目(ファイル名, 分類1, 分類2, ...)をヘッダーとして扱い、
    // 各分類名が積み上げ棒グラフのレジェンドに表示されるようにする
    .setNumHeaders(1)
    .setPosition(transition.chartAnchorRow, 1, 0, 0)
    .setOption('title', 'ファイルごとの分類別支出推移')
    .setOption('isStacked', true)
    .setOption('legend', { position: 'right' })
    .setOption('width', 600)
    .setOption('height', 350)
    .build();

  sheet.insertChart(chart);
}

// ===== デバッグ用 =====

/**
 * フォルダ内の最初のCSVファイルの先頭10行(A〜C列の生データ)をログに出力する。
 * classifyExpenses実行後もシートが空の場合、この関数を実行してログを確認すると
 * 「日付列がどんな文字列になっているか」「列がずれていないか」を切り分けられる。
 * ログの確認方法: Apps Scriptエディタで実行後、左側メニューの「実行数」から
 * 該当の実行を開くか、上部メニューの「表示 > ログ」で確認できる。
 */
function debugPrintFirstRows() {
  const folder = DriveApp.getFolderById(FOLDER_ID);

  // まずフォルダ内の全ファイルとMIMEタイプを一覧表示(想定外のファイルが無いか確認用)
  const allFiles = folder.getFiles();
  let fileCount = 0;
  while (allFiles.hasNext()) {
    const file = allFiles.next();
    Logger.log(`ファイル: ${file.getName()} / MimeType: ${file.getMimeType()}`);
    fileCount++;
  }
  if (fileCount === 0) {
    Logger.log('フォルダ内にファイルが1つもありません。FOLDER_IDや対象フォルダを確認してください。');
    return;
  }

  const csvFiles = getCsvFilesInFolder(folder);
  if (csvFiles.length === 0) {
    Logger.log('.csv拡張子のファイルが見つかりませんでした。上記のファイル名一覧を確認してください。');
    return;
  }

  const file = csvFiles[0];
  Logger.log('対象ファイル: ' + file.getName());
  const rows = readCsvRows(file);
  rows.slice(0, 10).forEach((row, i) => {
    Logger.log(`行${i}: A=[${row[0]}] B=[${row[1]}] C=[${row[2]}] isDate=${isDateString((row[0] || '').toString().trim())}`);
  });
}