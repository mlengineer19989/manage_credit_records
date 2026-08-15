/**
 * Vpass(三井住友カード)・楽天カードの両方の利用明細CSVをまとめて仕訳し、
 * 年月ごとに統合して円グラフ・積み上げ棒グラフ化するスクリプト(プロトタイプ)
 *
 * 【plot_from_vpass_csv.js / plot_from_rakten_csv.js との違い】
 * 2つのDriveフォルダ(Vpass用・楽天カード用)を読み込み、ファイル名から年月を取り出して
 * 同じ年月同士を1つの集計ブロックにまとめる。同じスプレッドシートに他の2つのスクリプトを
 * 貼り付けても衝突しないよう、定数名には "ALL_" を、関数名には "AllCards" を付与している。
 *
 * 【対象ファイル名】
 * ・Vpass:     YYYYMM.csv                (例: 202605.csv)
 * ・楽天カード: enaviYYYYMM(カード番号等).csv (例: enavi202605(1234).csv)
 *   ※ enavi の直後の6桁を年月として扱うため、カッコの中身(カード番号)は問わない。
 *     1つの年月に対して楽天側のファイルが複数(カードごと)あっても合算される。
 *
 * 【事前準備】
 * 1. このコードをGoogleスプレッドシートの「拡張機能 > Apps Script」のエディタに貼り付ける
 * 2. ALL_VPASS_FOLDER_ID / ALL_RAKUTEN_FOLDER_ID に、それぞれの月次CSVが入っている
 *    Google DriveフォルダのIDを設定する
 *    (フォルダを開いたときのURL末尾: https://drive.google.com/drive/folders/【ここ】)
 * 3. ALL_CATEGORY_DICT に、店舗名(部分一致)と分類の対応をどんどん追加していく
 *    (plot_from_vpass_csv.js / plot_from_rakten_csv.js と内容を揃えておくこと)
 * 4. classifyExpensesAllCards を実行(初回はGoogleアカウントの権限承認が必要)
 * 5. 実行後、スプレッドシートに「全体明細」「全体集計」「全体推移」シートが作成される
 *    - 「全体集計」シート: 年月ごと・全期間合計の分類別円グラフ
 *    - 「全体推移」シート: 年月ごとの分類別支出推移の積み上げ棒グラフ
 */

// ===== 設定 =====

// Vpass(三井住友カード)のCSVが保存されているDriveフォルダのID
const ALL_VPASS_FOLDER_ID = 'フォルダidをここに入力';

// 楽天カードのCSVが保存されているDriveフォルダのID
const ALL_RAKUTEN_FOLDER_ID = 'フォルダidをここに入力';

// CSVの文字コード(文字化けする場合は入れ替える)
const ALL_VPASS_CSV_ENCODING = 'Shift_JIS';
const ALL_RAKUTEN_CSV_ENCODING = 'UTF-8';

const ALL_DETAIL_SHEET_NAME = '全体明細';
const ALL_SUMMARY_SHEET_NAME = '全体集計';
const ALL_TRANSITION_SHEET_NAME = '全体推移';

// 「全体集計」シートで、年月1つ分の表+円グラフに割り当てる行数(ブロック間の間隔)
const ALL_SUMMARY_ROWS_PER_BLOCK = 25;

// 店舗名(部分一致)→ 分類 の対応表。上から順にチェックし、最初にマッチしたものを採用
// 分類は以下の8種類固定: スーパー / コンビニ / 自販機 / 外食 / 交通費 / 光熱費 / 娯楽・嗜好品 / サブスク / 未分類
// ※plot_from_vpass_csv.js / plot_from_rakten_csv.js と同じ内容。キーワードを追加する際は
//   3ファイルとも揃えておくこと。
const ALL_CATEGORY_DICT = {
  // ---- スーパー ----
  'イオン': 'スーパー',
  '西友': 'スーパー',
  'ライフ': 'スーパー',
  'いなげや': 'スーパー',
  '東急ストア': 'スーパー',
  'ハックドラッグ': 'スーパー', // ドラッグストア
  'おどや': 'スーパー',
  'ハナマサ': 'スーパー', // 肉のハナマサ
  'カルディ': 'スーパー',
  '食べチョク': 'スーパー', // 産直野菜通販
  'ドウェイ': 'スーパー', // フードウェイ(表記ゆれで先頭の長音記号がずれることがあるため末尾側で判定)
  'フレッシュオ': 'スーパー', // フレッシュオーワダ(八百屋)
  'Ｆｉｔ　Ｃａｒｅ　ＤＥＰＯＴ': 'スーパー',
  'アピタ': 'スーパー',
  '大野屋': 'スーパー',
  'ヤッキョク': 'スーパー', // 「ドリーム薬局」など(半角カナ変換後の表記)
  'ヤツキヨク': 'スーパー', // 「杉本薬局」など(小さい"ッ""ョ"が大きい字になる半角カナの表記ゆれ)

  // ---- コンビニ ----
  'セブン': 'コンビニ',
  'ファミリーマート': 'コンビニ',
  'ローソン': 'コンビニ',
  'ミニストップ': 'コンビニ',
  'ＮｅｗＤａｙｓ': 'コンビニ', // JR東日本の駅ナカコンビニ
  'ニューデイズ': 'コンビニ', // NewDaysのカナ表記
  'キヨスク': 'コンビニ', // グランドキヨスク・ベルマートキヨスク・東海キヨスク等
  'キオスク': 'コンビニ', // 表記ゆれ(「キヨスク」ではなく「キオスク」)

  // ---- 自販機 ----
  '飲料自販機': '自販機',
  'コカコ': '自販機', // コカコーラボトラーズ(自販機)
  'コカ・コーラ': '自販機', // 「コカ・コーラ Cmode 自販機」など(・区切り表記)
  'ダイドー': '自販機', // ダイドードリンコの自販機
  'エキナカ自販機': '自販機', // JR東日本 acure など

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
  '市場食堂': '外食',
  'コウジクラ': '外食',
  'ＨＯＮＥＹＢＥＥ': '外食',
  '東京土山人': '外食',
  '鉄板ステーキ＆ハンバーグ': '外食',
  'ＳＵＳＨＩ　ＳＯＵＬ': '外食',
  '新宿さぼてん': '外食',
  '四代目菊川': '外食',
  '金伝丸': '外食', // らーめん金伝丸
  'シェイクシャック': '外食',
  '銚子丸': '外食', // すし銚子丸
  'オムライスのひまわり': '外食',
  'ケンタッキー': '外食',
  'マルデナポリ': '外食',
  'はま寿司': '外食',
  'スシロ': '外食', // スシロー(長音記号が「—」表記になることがあるため末尾を含めない)
  'こがね食堂': '外食',
  'イカリヤシヨクドウ': '外食',
  'ＡＦＵＲＩ': '外食',
  'ＣｏＣｏ壱番屋': '外食',
  'ふふふあん': '外食',
  'アミーゴ': '外食', // イタリアンバー アミーゴ
  'ドンク': '外食', // ベーカリーチェーン
  'Ｒ　Ｂａｋｅｒ': '外食',
  'デリフランス': '外食',
  'ｍｍＴＨＡＩ': '外食',
  'バーガージョーズ': '外食',
  'プレミアムカルビ': '外食',
  '伝ＯＮ　ＴＨＥ　ＴＡＢＬＥ': '外食',
  'ｔｏ　ｙｏｕｒ　Ｆａｒｍ': '外食', // フレッシュオーワダ直営のサラダ専門店
  '三渓園茶寮': '外食',
  'プレジデン': '外食', // ベーカリー「プレジデンテ」
  'シュウマイ': '外食', // 毎週シュウマイ/マイシュウシュウマイ
  'エチゼンカニ': '外食',
  'たか松': '外食',
  '神戸スパイス': '外食',
  '駅弁にぎわい': '外食',
  'モスのネット注文': '外食',
  'ＭＲ．ＬＥＥ': '外食',
  'ポンパドウル': '外食', // ベーカリーチェーン
  'ドミノピザ': '外食',
  '回転寿司みさき': '外食',
  'そば処': '外食',
  'ブトウキヨウ': '外食', // GARB TOKYO(カフェ&レストラン)の長音記号ゆれ表記
  '越前かに': '外食', // 越前かに職人甲羅組
  'うま囲': '外食',
  'ジャノメヤ': '外食', // 麺処 蛇の目屋
  'ケンタツキー': '外食', // ケンタッキーの半角カナ表記ゆれ(小さい"ッ"が大きい字になる)
  'ヤキトリ': '外食',
  'ハブシブヤドウゲンザカ': '外食', // 英国風パブ「HUB」渋谷道玄坂店
  'サヴール': '外食', // サンス・エ・サヴール(ベーカリー)
  'ビストロアグリ': '外食',

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
  'Ｓｕｉｃａ': '交通費', // モバイルSuica(Apple)など全角表記
  'タクシー': '交通費',
  '三和交通': '交通費',
  '交通利用': '交通費', // 「遠鉄バス／交通利用」など
  '彌榮自動車': '交通費', // 京都のタクシー会社(通称やさか交通)
  '６２５タクシ': '交通費', // 「株式会社625タクシー横浜」(長音記号の表記ゆれを避けて手前で判定)
  'スマートイーエツクス': '交通費', // JRのスマートEX(半角カナ表記)
  'モバイルスイカ': '交通費', // モバイルSuicaのカナ表記

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
  'AMAZON': '娯楽・嗜好品', // 半角表記(AMAZON.CO.JPなど)
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
  'ホテル': '娯楽・嗜好品',
  'ＮＥＷｏＭａｎ': '娯楽・嗜好品',
  'ジョイナス': '娯楽・嗜好品',
  'シァル横浜': '娯楽・嗜好品',
  'メイワン': '娯楽・嗜好品', // 浜松駅の商業施設
  'そごう': '娯楽・嗜好品',
  'ソゴウ': '娯楽・嗜好品', // そごうのカナ表記
  '高島屋': '娯楽・嗜好品',
  '三越': '娯楽・嗜好品',
  '横浜モアーズ': '娯楽・嗜好品',
  'ワールドポーターズ': '娯楽・嗜好品',
  '三井アウトレットパーク': '娯楽・嗜好品',
  'ＭＡＲＫ　ＩＳ': '娯楽・嗜好品',
  'クイズミナトミライ': '娯楽・嗜好品', // 「マークイズみなとみらい」の長音記号ゆれ表記
  'ブリックスクエア': '娯楽・嗜好品',
  '日吉東急アベニュー': '娯楽・嗜好品',
  'ヒヨシトウキユウアベニユ': '娯楽・嗜好品', // 上記のカナ表記ゆれ(長音記号を含まない箇所で判定)
  'アトレ': '娯楽・嗜好品',
  '東急百貨店': '娯楽・嗜好品',
  'ユナイテッド・シネマ': '娯楽・嗜好品',
  'Ｔジョイ': '娯楽・嗜好品', // 映画館(T・ジョイ)
  'アソビュー': '娯楽・嗜好品',
  'Ｐｒｏｇａｔｅ': '娯楽・嗜好品',
  'ダイソー': '娯楽・嗜好品',
  'ニトリ': '娯楽・嗜好品',
  'ユニクロ': '娯楽・嗜好品',
  'エディオン': '娯楽・嗜好品',
  '三溪園': '娯楽・嗜好品', // 庭園の入園料等(「三渓園茶寮」は外食に分類)
  'ハンマーヘッド': '娯楽・嗜好品',
  'シヨツプアンドレス': '娯楽・嗜好品', // 「ハンマーヘッド SHOP&REST」のカナ表記ゆれ
  'コーナン': '娯楽・嗜好品',
  '道の駅': '娯楽・嗜好品',
  '海ほたる': '娯楽・嗜好品',
  'アイハーブ': '娯楽・嗜好品',
  'ラクテントラベル': '娯楽・嗜好品',
  '鍵善良房': '娯楽・嗜好品',
  'パティスリ': '娯楽・嗜好品', // 各種パティスリー(洋菓子店)
  'コージーコーナー': '娯楽・嗜好品',
  '亀屋万年堂': '娯楽・嗜好品',
  'エトモ': '娯楽・嗜好品', // etomo綱島(カナ表記)
  'ｅｔｏｍｏ': '娯楽・嗜好品', // etomo綱島(全角アルファベット表記)
  'クプラザ': '娯楽・嗜好品', // 「ランドマークプラザ」の長音記号ゆれ表記
  'デンハウスミナトミライ': '娯楽・嗜好品', // ガーデンハウスみなとみらい
  '東急スクエア': '娯楽・嗜好品',
  '阪神百貨店': '娯楽・嗜好品',
  '横浜ベイクオーター': '娯楽・嗜好品',
  'ＰＩＥＲＲＥ　ＭＡＲＣＯＬＩＮＩ': '娯楽・嗜好品',
  '半兵衛麸': '娯楽・嗜好品', // 京都の老舗麸店
  'エキマルシェ': '娯楽・嗜好品',
  'ポルタ': '娯楽・嗜好品', // 京都ポルタなど
  'ＡＳＴＹ': '娯楽・嗜好品',
  'サイト高崎': '娯楽・嗜好品', // イーサイト高崎
  'ＰＬＵＳＴＡ': '娯楽・嗜好品',
  '遠鉄百貨店': '娯楽・嗜好品',
  '新丸ビル': '娯楽・嗜好品',
  '丸の内ビルディング': '娯楽・嗜好品',
  'グランスタ': '娯楽・嗜好品', // 東京駅グランスタ
  '渋谷ストリーム': '娯楽・嗜好品',
  'ロンロン': '娯楽・嗜好品',
  'グツド　ネイチヤ': '娯楽・嗜好品', // Good Nature Station
  'ジヨイナス': '娯楽・嗜好品', // ジョイナスの半角カナ表記ゆれ(小さい"ョ"が大きい字になる)
  'エキマルシエ': '娯楽・嗜好品', // エキマルシェの半角カナ表記ゆれ
  'アステイ': '娯楽・嗜好品', // ASTY(京都/新大阪の駅ビル)の半角カナ表記
  'コレツトマーレ': '娯楽・嗜好品', // コレットマーレ(横須賀の商業施設)
  'イーサイトタカサキ': '娯楽・嗜好品', // イーサイト高崎の半角カナ表記
  'ヴイアイン': '娯楽・嗜好品', // ホテル「ヴィアイン」
  'トウヨコイン': '娯楽・嗜好品', // ホテル「東横INN」
  'ハンキユウサンバンガイ': '娯楽・嗜好品', // 阪急三番街
  'ニジユウバシスクエ': '娯楽・嗜好品', // 丸の内二重橋スクエア
  'トレフアクスタイル': '娯楽・嗜好品', // リユースショップ
  'ジーユー': '娯楽・嗜好品', // GU
  'ビームス': '娯楽・嗜好品', // BEAMS
  'カルバンクライン': '娯楽・嗜好品',
  'LILYSILK': '娯楽・嗜好品',
  '書泉': '娯楽・嗜好品',
  'ｂｏｏｋｆａｎ': '娯楽・嗜好品',
  '楽天ブックス': '娯楽・嗜好品',
  '楽天トラベル': '娯楽・嗜好品',
  '渋谷スクランブルスクエア': '娯楽・嗜好品',
  'デイーエムエム': '娯楽・嗜好品', // DMM.comの各種課金(半角カナ表記)

  // ---- サブスク ----
  'かたぎり塾': 'サブスク', // パーソナルトレーニングジム(月額会費)
  'パーソナルトレーニングジム': 'サブスク', // かたぎり塾以外のパーソナルジムも含む汎用キーワード
  'アマゾンプライム': 'サブスク', // Amazon Prime会費(カナ表記)
  'CURSOR': 'サブスク', // Cursorの利用料金
  'GITHUB': 'サブスク',
  'ＮＥＴＦＬＩＸ': 'サブスク',
  'ネットフリックス': 'サブスク', // カナ表記
  'ｃｈｏｃｏＺＡＰ': 'サブスク',
  'ソフトバンク': 'サブスク', // 携帯電話料金(月額)
  'ＡＰＰＬＥ': 'サブスク', // App Store/iCloud等のサブスク(全角表記)
  'APPLE COM BILL': 'サブスク', // 半角表記の請求
  'APPLE.COM': 'サブスク', // 半角表記の請求
  'ソニーインタラクティブ': 'サブスク', // PlayStation Networkの課金
  'プレイステーションネットワーク': 'サブスク',
  'ダウンタウンプラス': 'サブスク',

  // 必要に応じて追加していく
};

const ALL_UNCLASSIFIED_LABEL = '未分類';

// 積み上げ棒グラフ・推移表で使う分類の並び順(固定)
const ALL_CATEGORY_ORDER = ['スーパー', 'コンビニ', '自販機', '外食', '交通費', '光熱費', '娯楽・嗜好品', 'サブスク', ALL_UNCLASSIFIED_LABEL];

// ===== メイン処理 =====

function classifyExpensesAllCards() {
  const vpassFolder = DriveApp.getFolderById(ALL_VPASS_FOLDER_ID);
  const rakutenFolder = DriveApp.getFolderById(ALL_RAKUTEN_FOLDER_ID);

  const vpassFiles = getCsvFilesInFolderAllCards(vpassFolder);
  const rakutenFiles = getCsvFilesInFolderAllCards(rakutenFolder);

  const transactions = []; // {yearMonth, source, fileName, date, store, amount, category}

  vpassFiles.forEach(file => {
    const yearMonth = extractYearMonthFromVpassFileNameAllCards(file.getName());
    if (!yearMonth) return; // "YYYYMM.csv" 形式でないファイルは対象外
    const rows = readCsvRowsAllCards(file, ALL_VPASS_CSV_ENCODING);

    rows.forEach(row => {
      const rawDate = (row[0] || '').toString().trim();
      if (!isDateStringAllCards(rawDate)) return; // 見出し行・合計行などはスキップ
      const dateStr = toHalfWidthAllCards(rawDate);

      const store = toFullWidthKatakanaAllCards((row[1] || '').toString().trim());
      const amount = parseAmountAllCards(row[2]); // Vpassは C列(index=2)が金額
      if (!store || isNaN(amount)) return;

      transactions.push({
        yearMonth: yearMonth,
        source: 'Vpass',
        fileName: file.getName(),
        date: dateStr,
        store: store,
        amount: amount,
        category: classifyStoreAllCards(store),
      });
    });
  });

  rakutenFiles.forEach(file => {
    const yearMonth = extractYearMonthFromRakutenFileNameAllCards(file.getName());
    if (!yearMonth) return; // "enaviYYYYMM(...).csv" 形式でないファイルは対象外
    const rows = readCsvRowsAllCards(file, ALL_RAKUTEN_CSV_ENCODING);

    rows.forEach(row => {
      const rawDate = (row[0] || '').toString().trim();
      if (!isDateStringAllCards(rawDate)) return; // 見出し行・合計行などはスキップ
      const dateStr = toHalfWidthAllCards(rawDate);

      const store = toFullWidthKatakanaAllCards((row[1] || '').toString().trim());
      const amount = parseAmountAllCards(row[4]); // 楽天カードCSVはE列(index=4)が金額
      if (!store || isNaN(amount)) return;

      transactions.push({
        yearMonth: yearMonth,
        source: '楽天カード',
        fileName: file.getName(),
        date: dateStr,
        store: store,
        amount: amount,
        category: classifyStoreAllCards(store),
      });
    });
  });

  // 年月(YYYYMM)の文字列としての昇順 = 時系列順。これによりVpass・楽天カードの
  // 同じ年月のファイルが自動的に同じブロックに統合される。
  const yearMonths = Array.from(new Set(transactions.map(t => t.yearMonth))).sort();

  writeDetailSheetAllCards(transactions);
  const summary = writeSummarySheetAllCards(transactions, yearMonths);
  drawPieChartsAllCards(summary);

  const transitionSummary = writeTransitionSheetAllCards(transactions, yearMonths);
  drawTransitionChartAllCards(transitionSummary.sheet, transitionSummary.transition);
}

// ===== CSV読み込み =====

function readCsvRowsAllCards(file, encoding) {
  let csvText;
  try {
    csvText = file.getBlob().getDataAsString(encoding);
  } catch (e) {
    const fallbackEncoding = encoding === 'UTF-8' ? 'Shift_JIS' : 'UTF-8';
    csvText = file.getBlob().getDataAsString(fallbackEncoding);
  }
  // UTF-8のBOM(先頭の目に見えない1文字)が残っていると1行目の1列目だけ余計な文字が付くことがあるため取り除く
  if (csvText.charCodeAt(0) === 0xFEFF) {
    csvText = csvText.slice(1);
  }
  return Utilities.parseCsv(csvText);
}

function getCsvFilesInFolderAllCards(folder) {
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
  // ファイル名順に並べておく(年月抽出後にさらにソートするため、ここでは厳密でなくてよい)
  csvFiles.sort((a, b) => a.getName().localeCompare(b.getName(), 'ja'));
  return csvFiles;
}

// "202605.csv" のようなファイル名から年月"202605"を取り出す。マッチしなければnull。
function extractYearMonthFromVpassFileNameAllCards(fileName) {
  const m = fileName.match(/^(\d{6})\.csv$/i);
  return m ? m[1] : null;
}

// "enavi202605(1234).csv" のようなファイル名から年月"202605"を取り出す。
// カッコの中身(カード番号等)や拡張子より前の余分な文字列は問わない。マッチしなければnull。
function extractYearMonthFromRakutenFileNameAllCards(fileName) {
  const m = fileName.match(/^enavi(\d{6})/i);
  return m ? m[1] : null;
}

function toHalfWidthAllCards(str) {
  // 全角の数字・記号を半角に変換する(CSVによっては全角表記が混ざっていることがあるため)
  const zenkakuMap = {
    '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
    '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
    '／': '/', '－': '-', '．': '.', '，': ',', '￥': '¥',
  };
  return str.replace(/[０-９／－．，￥]/g, ch => zenkakuMap[ch] || ch);
}

// 半角カタカナ(ｱｲｳｴｵ...)→ 全角カタカナ の変換表。クレジットカードCSVの店舗名が
// 半角カナで記録されているケースがあり、そのままだとALL_CATEGORY_DICT(全角カナ)に
// マッチしないため、分類前に必ずこの変換をかける。
const ALL_HALFWIDTH_KATAKANA_MAP = {
  'ｦ': 'ヲ', 'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ', 'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ', 'ｯ': 'ッ',
  'ｰ': 'ー', 'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
  'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
  'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ',
  'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
  'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
  'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
  'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
  'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
  'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
  'ﾜ': 'ワ', 'ﾝ': 'ン',
  '｡': '。', '｢': '「', '｣': '」', '､': '、', '･': '・',
};
// 濁点(ﾞ)・半濁点(ﾟ)付きの半角カナは、単独の全角カナ2文字ではなく1文字の濁音/半濁音になる
const ALL_HALFWIDTH_DAKUTEN_MAP = {
  'ｶ': 'ガ', 'ｷ': 'ギ', 'ｸ': 'グ', 'ｹ': 'ゲ', 'ｺ': 'ゴ',
  'ｻ': 'ザ', 'ｼ': 'ジ', 'ｽ': 'ズ', 'ｾ': 'ゼ', 'ｿ': 'ゾ',
  'ﾀ': 'ダ', 'ﾁ': 'ヂ', 'ﾂ': 'ヅ', 'ﾃ': 'デ', 'ﾄ': 'ド',
  'ﾊ': 'バ', 'ﾋ': 'ビ', 'ﾌ': 'ブ', 'ﾍ': 'ベ', 'ﾎ': 'ボ',
  'ｳ': 'ヴ',
};
const ALL_HALFWIDTH_HANDAKUTEN_MAP = {
  'ﾊ': 'パ', 'ﾋ': 'ピ', 'ﾌ': 'プ', 'ﾍ': 'ペ', 'ﾎ': 'ポ',
};

function toFullWidthKatakanaAllCards(str) {
  let result = str.replace(/[\uFF61-\uFF9F][\uFF9E\uFF9F]?/g, match => {
    const base = match[0];
    const mark = match[1];
    if (mark === 'ﾞ' && ALL_HALFWIDTH_DAKUTEN_MAP[base]) return ALL_HALFWIDTH_DAKUTEN_MAP[base];
    if (mark === 'ﾟ' && ALL_HALFWIDTH_HANDAKUTEN_MAP[base]) return ALL_HALFWIDTH_HANDAKUTEN_MAP[base];
    return (ALL_HALFWIDTH_KATAKANA_MAP[base] || base) + (mark || '');
  });
  // 半角カナでは長音記号(ー)の代わりに半角ハイフン(-)が使われていることがあるため、
  // カタカナの直後のハイフンは長音記号とみなす
  result = result.replace(/([\u30A1-\u30FC])-/g, '$1ー');
  return result;
}

function isDateStringAllCards(str) {
  // yyyy/mm/dd, yy/mm/dd, yyyy-mm-dd など(全角数字・2桁年も許容)のみレコードとして扱う
  const normalized = toHalfWidthAllCards(str);
  return /^\d{2,4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(normalized);
}

function parseAmountAllCards(value) {
  if (value === null || value === undefined) return NaN;
  const normalized = toHalfWidthAllCards(value.toString());
  const cleaned = normalized.replace(/[,¥\s]/g, '');
  return Number(cleaned);
}

// "202605" → "2026/05" のように表示用に整形する
function formatYearMonthLabelAllCards(yearMonth) {
  return yearMonth.slice(0, 4) + '/' + yearMonth.slice(4, 6);
}

// ===== 分類ロジック =====

function classifyStoreAllCards(store) {
  for (const keyword in ALL_CATEGORY_DICT) {
    if (store.indexOf(keyword) !== -1) {
      return ALL_CATEGORY_DICT[keyword];
    }
  }
  return ALL_UNCLASSIFIED_LABEL;
}

// ===== シート書き出し =====

function getOrCreateSheetAllCards(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  sheet.clear();
  return sheet;
}

function writeDetailSheetAllCards(transactions) {
  const sheet = getOrCreateSheetAllCards(ALL_DETAIL_SHEET_NAME);
  sheet.appendRow(['年月', 'カード', 'ファイル名', '日付', '店舗', '金額', '分類']);
  const rows = transactions.map(t => [formatYearMonthLabelAllCards(t.yearMonth), t.source, t.fileName, t.date, t.store, t.amount, t.category]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
}

function writeSummarySheetAllCards(transactions, yearMonths) {
  const sheet = getOrCreateSheetAllCards(ALL_SUMMARY_SHEET_NAME);
  const blocks = [];

  // 年月ごとに(Vpass・楽天カードを合算した)分類・合計金額のブロックを積み上げて書き出す
  yearMonths.forEach((yearMonth, i) => {
    const monthTransactions = transactions.filter(t => t.yearMonth === yearMonth);
    blocks.push(writeSummaryBlockAllCards(sheet, i * ALL_SUMMARY_ROWS_PER_BLOCK + 1, formatYearMonthLabelAllCards(yearMonth), monthTransactions));
  });

  // 最後に全期間合計のブロックを追加
  const totalStartRow = yearMonths.length * ALL_SUMMARY_ROWS_PER_BLOCK + 1;
  blocks.push(writeSummaryBlockAllCards(sheet, totalStartRow, '全期間合計', transactions));

  return { sheet, blocks };
}

// 「全体推移」シート: 円グラフの集計(「全体集計」シート)とは別タブに、
// 年月ごとの分類別推移の表+積み上げ棒グラフを書き出す
function writeTransitionSheetAllCards(transactions, yearMonths) {
  const sheet = getOrCreateSheetAllCards(ALL_TRANSITION_SHEET_NAME);
  const transition = writeTransitionSectionAllCards(sheet, 1, yearMonths, transactions);
  return { sheet, transition };
}

function writeSummaryBlockAllCards(sheet, startRow, label, transactions) {
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

function writeTransitionSectionAllCards(sheet, startRow, yearMonths, transactions) {
  sheet.getRange(startRow, 1).setValue('年月ごとの推移').setFontWeight('bold');

  const headerRow = startRow + 1;
  const columnCount = ALL_CATEGORY_ORDER.length + 1; // 年月 + 各分類
  sheet.getRange(headerRow, 1, 1, columnCount).setValues([['年月'].concat(ALL_CATEGORY_ORDER)]);

  // 年月ごとに、分類の並び順(ALL_CATEGORY_ORDER)に沿って合計金額を並べる(0円の分類も含める)
  const dataRows = yearMonths.map(yearMonth => {
    const totals = {};
    transactions
      .filter(t => t.yearMonth === yearMonth)
      .forEach(t => {
        totals[t.category] = (totals[t.category] || 0) + t.amount;
      });
    return [formatYearMonthLabelAllCards(yearMonth)].concat(ALL_CATEGORY_ORDER.map(category => totals[category] || 0));
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

function drawPieChartsAllCards(summary) {
  const { sheet, blocks } = summary;

  // 既存のグラフを削除してから作り直す(再実行しても増殖しないように)
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));

  blocks.forEach(block => {
    if (block.rowCount === 0) return; // データが無い年月はグラフを描かない

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

function drawTransitionChartAllCards(sheet, transition) {
  // 既存のグラフを削除してから作り直す(再実行しても増殖しないように)
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));

  if (transition.rowCount === 0) return; // 年月データが無ければ何もしない

  const dataRange = sheet.getRange(transition.headerRow, 1, transition.rowCount + 1, transition.columnCount);
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    // 1行目(年月, 分類1, 分類2, ...)をヘッダーとして扱い、
    // 各分類名が積み上げ棒グラフのレジェンドに表示されるようにする
    .setNumHeaders(1)
    .setPosition(transition.chartAnchorRow, 1, 0, 0)
    .setOption('title', '年月ごとの分類別支出推移(Vpass+楽天カード合算)')
    .setOption('isStacked', true)
    .setOption('legend', { position: 'right' })
    .setOption('width', 600)
    .setOption('height', 350)
    .build();

  sheet.insertChart(chart);
}

// ===== デバッグ用 =====

/**
 * Vpass用・楽天カード用の両フォルダについて、ファイル一覧と年月の抽出結果、
 * 先頭ファイルの先頭10行の生データをログに出力する。
 * classifyExpensesAllCards実行後もシートが空/意図しない結果の場合、この関数を実行して
 * ログを確認すると「ファイル名から年月が正しく取れているか」「列がずれていないか」を
 * 切り分けられる。
 * ログの確認方法: Apps Scriptエディタで実行後、左側メニューの「実行数」から
 * 該当の実行を開くか、上部メニューの「表示 > ログ」で確認できる。
 */
function debugPrintFirstRowsAllCards() {
  debugPrintFolderAllCards('Vpass', ALL_VPASS_FOLDER_ID, ALL_VPASS_CSV_ENCODING, extractYearMonthFromVpassFileNameAllCards, 2);
  debugPrintFolderAllCards('楽天カード', ALL_RAKUTEN_FOLDER_ID, ALL_RAKUTEN_CSV_ENCODING, extractYearMonthFromRakutenFileNameAllCards, 4);
}

function debugPrintFolderAllCards(label, folderId, encoding, extractYearMonthFn, amountColumnIndex) {
  Logger.log(`===== ${label} =====`);
  const folder = DriveApp.getFolderById(folderId);

  const allFiles = folder.getFiles();
  let fileCount = 0;
  while (allFiles.hasNext()) {
    const file = allFiles.next();
    const yearMonth = extractYearMonthFn(file.getName());
    Logger.log(`ファイル: ${file.getName()} / MimeType: ${file.getMimeType()} / 抽出した年月: ${yearMonth}`);
    fileCount++;
  }
  if (fileCount === 0) {
    Logger.log(`${label}: フォルダ内にファイルが1つもありません。フォルダIDや対象フォルダを確認してください。`);
    return;
  }

  const csvFiles = getCsvFilesInFolderAllCards(folder);
  if (csvFiles.length === 0) {
    Logger.log(`${label}: .csv拡張子のファイルが見つかりませんでした。上記のファイル名一覧を確認してください。`);
    return;
  }

  const file = csvFiles[0];
  Logger.log(`${label} 対象ファイル: ` + file.getName());
  const rows = readCsvRowsAllCards(file, encoding);
  rows.slice(0, 10).forEach((row, i) => {
    Logger.log(`行${i}: A=[${row[0]}] B=[${row[1]}] 金額列[${amountColumnIndex}]=[${row[amountColumnIndex]}] isDate=${isDateStringAllCards((row[0] || '').toString().trim())}`);
  });
}
