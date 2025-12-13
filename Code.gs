/**
 * CF自動更新システム v5.0 - Cash Flow管理特化
 * Actual（実績）とPlan（予定）の完全分離
 * 日付スパイン + 残高連続表示
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 CF自動更新')
    .addItem('🎛️ コントロールパネルを開く', 'showSidebar')
    .addSeparator()
    .addItem('🚀 システム初期化', 'initializeDatabase')
    .addSeparator()
    .addSubMenu(ui.createMenu('🏦 データ管理')
      .addItem('資金台帳確認', 'refreshTransactions')
      .addItem('振替検出', 'detectTransfers')
      .addItem('DB_Transactions再構築', 'resetTransactionsSheet'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 表示切替')
      .addItem('Month_Viewを開く', 'openMonthView')
      .addItem('DB_Transactionsを開く', 'openTransactions')
      .addItem('Settingsを開く', 'openSettings'))
    .addSeparator()
    .addItem('📋 全シート状態確認', 'checkAllSheets')
    .addToUi();

  showToast('💰 CF自動更新 v5.0', 'Cash Flow管理 稼働中', 5);
}

/**
 * HTMLサイドバーを表示
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('💰 CF自動更新 v5.0')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
  showToast('🎛️ コントロールパネル', 'サイドバーを開きました', 2);
}

/**
 * Toast通知
 */
function showToast(title, message = '', duration = 3) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (message) {
    ss.toast(message, title, duration);
  } else {
    ss.toast(title, '通知', duration);
  }
}

/**
 * データベース全体を初期化
 * v5.0: Cash Flow管理特化アーキテクチャ
 */
function initializeDatabase() {
  showToast('🚀 初期化開始', 'Cash Flow管理システムを構築中...', 3);

  try {
    // 汎用ソースシート（6つ）
    for (let i = 1; i <= 6; i++) {
      setupSourceSheet(i);
    }

    // 新アーキテクチャのシート群
    setupDB_Transactions();  // 資金台帳（旧DB_Integrated）
    setupDB_Master();        // キーワードルール
    setupInput_CashPlan();   // 予定取引（新規）
    setupCalendar();         // 日付スパイン（新規）
    setupSettings();         // 設定（対象月・期首残高）
    setupMonth_View();       // 月次資金予実表（メイン画面）

    showToast('✅ 初期化完了！', 'Cash Flow管理システムが稼働しました', 5);

    return {
      success: true,
      message: '初期化完了',
      sheets: ['Source_1-6', 'DB_Transactions', 'DB_Master', 'Input_CashPlan', 'Calendar', 'Settings', 'Month_View']
    };
  } catch (error) {
    showToast('❌ エラー', error.message, 10);
    Logger.log('初期化エラー: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * 汎用ソースシート作成（Source_1 〜 Source_6）
 * MoneyForwardのデフォルト形式に準拠
 * @param {number} num - シート番号（1-6）
 */
function setupSourceSheet(num) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Source_${num}`;
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // 既に設定済みの場合はスキップ
  if (sheet.getRange('A1').getValue() !== '' && sheet.getRange('A1').getValue() !== '☑') {
    Logger.log(`${sheetName} は既に設定済み`);
    return;
  }

  // ヘッダー行（MoneyForwardデフォルト形式 + チェックボックス）
  const headers = ['☑', '日付', '内容', '金額', '残高', '連携サービス', 'ステータス', '取引No'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーのスタイル（番号ごとに色を変える）
  const colors = ['#1a73e8', '#34a853', '#fbbc04', '#ea4335', '#9c27b0', '#ff6d00'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(colors[num - 1]);
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // 列幅設定
  sheet.setColumnWidth(1, 50);   // ☑
  sheet.setColumnWidth(2, 100);  // 日付
  sheet.setColumnWidth(3, 250);  // 内容
  sheet.setColumnWidth(4, 120);  // 金額
  sheet.setColumnWidth(5, 120);  // 残高
  sheet.setColumnWidth(6, 150);  // 連携サービス
  sheet.setColumnWidth(7, 100);  // ステータス
  sheet.setColumnWidth(8, 100);  // 取引No

  // K1セルに大きくメモ欄を作成
  const memoCell = sheet.getRange('K1');
  memoCell.setValue(`ここは【　　　　　　】`);
  memoCell.setFontSize(14);
  memoCell.setFontWeight('bold');
  memoCell.setFontColor(colors[num - 1]);
  memoCell.setBackground('#fff3e0');
  memoCell.setBorder(true, true, true, true, true, true, '#ff6d00', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 使い方説明（K列以降）
  sheet.getRange('K2').setValue('💡 使い方');
  sheet.getRange('K3').setValue(`1. 上のメモ欄に銀行名を記入`);
  sheet.getRange('K4').setValue('2. MoneyForwardで該当口座を絞り込み');
  sheet.getRange('K5').setValue('3. 全期間を選択してコピー');
  sheet.getRange('K6').setValue('4. A2セル（ヘッダーの下）に貼り付け');
  sheet.getRange('K7').setValue('5. A列は空欄でOK（チェックボックス用）');
  sheet.getRange('K8').setValue('6. 毎回全期間上書きでOK！');
  sheet.getRange('K9').setValue('');
  sheet.getRange('K10').setValue('⚠️ 注意: A1ヘッダーは削除しないこと');

  // 列幅調整
  sheet.setColumnWidth(11, 280); // K列

  Logger.log(`${sheetName} 作成完了（MF形式）`);
}

/**
 * DB_Transactions シート作成（資金台帳）
 * v5.0: Cash Flow管理特化
 * 列: 日付, 口座, 摘要, 金額(+/-), 科目, タグ, UID, 転記元
 */
function setupDB_Transactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DB_Transactions');

  if (!sheet) {
    sheet = ss.insertSheet('DB_Transactions');
  }

  sheet.clear();

  // ═══════════════════════════════════════════════════
  // Step 1: ヘッダー設定（A1:H1）
  // ═══════════════════════════════════════════════════
  const headers = ['日付', '口座', '摘要', '金額', '科目', 'タグ', 'UID', '転記元'];
  sheet.getRange('A1:H1').setValues([headers]);

  // ヘッダースタイル
  const headerRange = sheet.getRange('A1:H1');
  headerRange.setBackground('#0b5394');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setFontSize(11);

  // ═══════════════════════════════════════════════════
  // Step 2: ARRAYFORMULA列の構築（A2-H列）
  // ═══════════════════════════════════════════════════

  // Source_1-6から統合データを取得するための内部シート参照用
  // J2からデータ開始（ヘッダーなし）
  const rawQueryFormula = `=QUERY({Source_1!A2:H; Source_2!A2:H; Source_3!A2:H; Source_4!A2:H; Source_5!A2:H; Source_6!A2:H}, "where Col2 is not null", 0)`;
  sheet.getRange('J2').setFormula(rawQueryFormula);

  // MoneyForwardフォーマット（J列以降）:
  // J列=☑, K列=日付, L列=内容, M列=金額, N列=残高, O列=連携サービス, P列=ステータス, Q列=取引No

  // A2: 日付(整形) - K列（日付）から
  sheet.getRange('A2').setFormula('=ARRAYFORMULA(IF(K2:K="", "", DATEVALUE(LEFT(K2:K, 10))))');

  // B2: 口座 - O列（連携サービス）から
  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(O2:O="", "", O2:O))');

  // C2: 摘要 - L列（内容）から
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(L2:L="", "", L2:L))');

  // D2: 金額(+/-) - M列（金額）を数値化（入金+/出金-）
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(M2:M="", "", VALUE(REGEXREPLACE(TO_TEXT(M2:M), "[^0-9-]", ""))))');

  // E2: 科目 - DB_Masterからキーワードマッチング
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(C2:C="", "", IFERROR(INDEX(DB_Master!B:B, MATCH(TRUE, ISNUMBER(SEARCH(DB_Master!A:A, C2:C)), 0)), "未分類")))');

  // F2: タグ - DB_Masterから詳細タグ
  sheet.getRange('F2').setFormula('=ARRAYFORMULA(IF(C2:C="", "", IFERROR(INDEX(DB_Master!C:C, MATCH(TRUE, ISNUMBER(SEARCH(DB_Master!A:A, C2:C)), 0)), "")))');

  // G2: UID - 口座+取引Noで一意キー生成
  sheet.getRange('G2').setFormula('=ARRAYFORMULA(IF(B2:B="", "", B2:B & "-" & Q2:Q))');

  // H2: 転記元 - 固定値「MF連携」
  sheet.getRange('H2').setFormula('=ARRAYFORMULA(IF(K2:K="", "", "MF連携"))');

  // ═══════════════════════════════════════════════════
  // Step 3: 列幅調整
  // ═══════════════════════════════════════════════════
  sheet.setColumnWidth(1, 100);  // 日付
  sheet.setColumnWidth(2, 150);  // 口座
  sheet.setColumnWidth(3, 250);  // 摘要
  sheet.setColumnWidth(4, 120);  // 金額
  sheet.setColumnWidth(5, 150);  // 科目
  sheet.setColumnWidth(6, 150);  // タグ
  sheet.setColumnWidth(7, 200);  // UID
  sheet.setColumnWidth(8, 100);  // 転記元

  // J列以降は非表示（内部データ）
  sheet.hideColumns(10, 10);

  // ═══════════════════════════════════════════════════
  // Step 4: 説明欄
  // ═══════════════════════════════════════════════════
  sheet.getRange('T1').setValue('💰 資金台帳（DB_Transactions）');
  sheet.getRange('T1').setFontSize(14).setFontWeight('bold').setFontColor('#0b5394');
  sheet.getRange('T2').setValue('');
  sheet.getRange('T3').setValue('【原則】');
  sheet.getRange('T4').setValue('✅ 真実は「実際に口座残高が動いた取引」だけ');
  sheet.getRange('T5').setValue('✅ 入金はプラス、出金はマイナスで統一');
  sheet.getRange('T6').setValue('✅ UPSIDERも銀行口座と同格');
  sheet.getRange('T7').setValue('');
  sheet.getRange('T8').setValue('【列の意味】');
  sheet.getRange('T9').setValue('日付: 取引発生日');
  sheet.getRange('T10').setValue('口座: 資金が動いた口座・サービス名');
  sheet.getRange('T11').setValue('摘要: 取引内容');
  sheet.getRange('T12').setValue('金額: 入金+/出金-');
  sheet.getRange('T13').setValue('科目: 自動仕訳（DB_Master参照）');
  sheet.getRange('T14').setValue('タグ: 詳細分類');
  sheet.getRange('T15').setValue('UID: 一意キー（重複検知用）');
  sheet.getRange('T16').setValue('転記元: データソース');
  sheet.getRange('T17').setValue('');
  sheet.getRange('T18').setValue('【禁止事項】');
  sheet.getRange('T19').setValue('❌ このシートに直接入力しない');
  sheet.getRange('T20').setValue('❌ 数式を変更しない');

  sheet.setColumnWidth(20, 280); // T列

  Logger.log('DB_Transactions 作成完了（資金台帳 v5.0）');
}

/**
 * 資金台帳データを確認
 * ※ ARRAYFORMULA により自動更新されるため、通常は不要
 * ※ 数式が壊れた場合の緊急復旧用
 */
function refreshTransactions() {
  showToast('🔄 確認中...', '資金台帳の状態を確認します', 2);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_Transactions');

  if (!sheet) {
    showToast('❌ エラー', 'DB_Transactionsシートがありません', 5);
    return { success: false, message: 'シートが見つかりません' };
  }

  try {
    // A列のデータ行数を取得
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      showToast('⚠️ データなし', 'Source_1〜6にデータを貼り付けてください', 5);
      return { success: false, message: 'データがありません' };
    }

    // ARRAYFORMULAの存在確認
    const a2Formula = sheet.getRange('A2').getFormula();
    const d2Formula = sheet.getRange('D2').getFormula();

    if (!a2Formula || !d2Formula) {
      showToast('⚠️ 数式エラー', 'DB_Transactionsを再構築してください', 5);
      return { success: false, message: '数式が見つかりません。resetTransactionsSheet()を実行してください。' };
    }

    showToast('✅ 正常稼働中！', `${lastRow - 1}行のデータが自動処理されています`, 5);
    Logger.log(`DB_Transactions確認: ${lastRow - 1}行 (ARRAYFORMULA稼働中)`);

    return {
      success: true,
      message: `${lastRow - 1}行処理完了（自動更新中）`,
      rowCount: lastRow - 1
    };
  } catch (error) {
    showToast('❌ エラー', error.message, 10);
    Logger.log('資金台帳確認エラー: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * DB_Master シート（脳みそ）
 * 自動仕訳のルールを管理
 */
function setupDB_Master() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DB_Master');

  if (!sheet) {
    sheet = ss.insertSheet('DB_Master');
  }

  if (sheet.getRange('A1').getValue() !== '') {
    Logger.log('DB_Master は既に設定済み');
    return;
  }

  // ヘッダー
  const headers = ['検索キーワード', '判定カテゴリ', '詳細タグ', '優先度'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#674ea7');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // サンプルデータ（ユーザーの例に準拠 + 優先度追加）
  const sampleData = [
    ['振込手数料', '支払手数料', '銀行手数料', 1],
    ['UnivaPay', '売上', '決済入金', 1],
    ['UPSIDER', '立替金', 'カード利用', 2],
    ['GOOGLE', '広告宣伝費', 'Google広告', 1],
    ['カ）オールエーアイ', '外注費', 'All AI', 1],
    ['振込＊モカ', '役員報酬', '代表報酬', 1],
    ['PayPay', '売上', 'PayPay決済', 1],
    ['Amazon', '消耗品費', 'Amazon購入', 2],
    ['さくら', '通信費', 'さくらサーバー', 1],
    ['Adobe', '新聞図書費', 'Adobe CC', 1],
    ['みずほ', '手数料', 'みずほ銀行', 3],
    ['SBI', '手数料', 'SBI銀行', 3],
    ['楽天', '手数料', '楽天銀行', 3],
    ['Notion', '通信費', 'Notion利用料', 2],
    ['GitHub', '通信費', 'GitHub利用料', 2],
    ['AWS', '通信費', 'AWS利用料', 2]
  ];

  sheet.getRange(2, 1, sampleData.length, 4).setValues(sampleData);

  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 80);

  // 使い方説明
  sheet.getRange('E1').setValue('🧠 自動仕訳の脳みそ');
  sheet.getRange('E1').setFontSize(14).setFontWeight('bold').setFontColor('#674ea7');
  sheet.getRange('E2').setValue('');
  sheet.getRange('E3').setValue('【仕組み】');
  sheet.getRange('E4').setValue('DB_Integratedの「内容」列に');
  sheet.getRange('E5').setValue('A列のキーワードが含まれるか検索');
  sheet.getRange('E6').setValue('→ 該当したらB列・C列を自動入力');
  sheet.getRange('E7').setValue('');
  sheet.getRange('E8').setValue('【運用ルール】');
  sheet.getRange('E9').setValue('✅ 上の行ほど優先度が高い');
  sheet.getRange('E10').setValue('✅ 部分一致で検索（前方一致不要）');
  sheet.getRange('E11').setValue('✅ 「未分類」が出たらここに追加');
  sheet.getRange('E12').setValue('✅ 追加した瞬間、自動で反映される');
  sheet.getRange('E13').setValue('');
  sheet.getRange('E14').setValue('⚠️ A列は大文字小文字を区別します');

  // 列幅調整
  sheet.setColumnWidth(5, 280); // E列

  Logger.log('DB_Master 作成完了（脳みそ）');
}

/**
 * Input_CashPlan シート（予定取引）
 * v5.0: 未来の予定される資金移動を管理
 */
function setupInput_CashPlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Input_CashPlan');

  if (!sheet) {
    sheet = ss.insertSheet('Input_CashPlan');
  }

  if (sheet.getRange('A1').getValue() !== '') {
    Logger.log('Input_CashPlan は既に設定済み');
    return;
  }

  // ヘッダー
  const headers = ['予定日', '口座', '科目', 'タグ', '予定金額', '繰り返し', 'ステータス', 'メモ'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#e67e22');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // サンプルデータ
  const sampleData = [
    [new Date(2025, 0, 25), 'みずほ銀行', '家賃', '事務所家賃', -200000, '毎月25日', '予定', ''],
    [new Date(2025, 0, 31), 'みずほ銀行', '人件費', '給与', -300000, '毎月末日', '予定', ''],
    [new Date(2025, 1, 10), 'UPSIDER', '広告宣伝費', 'Google広告', -150000, '', '予定', '代表枠'],
    [new Date(2025, 1, 15), 'みずほ銀行', '売上', 'クライアントA', 500000, '', '予定', '']
  ];

  sheet.getRange(2, 1, sampleData.length, 8).setValues(sampleData);

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // 予定日
  sheet.setColumnWidth(2, 150);  // 口座
  sheet.setColumnWidth(3, 150);  // 科目
  sheet.setColumnWidth(4, 150);  // タグ
  sheet.setColumnWidth(5, 120);  // 予定金額
  sheet.setColumnWidth(6, 100);  // 繰り返し
  sheet.setColumnWidth(7, 80);   // ステータス
  sheet.setColumnWidth(8, 200);  // メモ

  // 説明欄
  sheet.getRange('J1').setValue('📅 予定取引（Input_CashPlan）');
  sheet.getRange('J1').setFontSize(14).setFontWeight('bold').setFontColor('#e67e22');
  sheet.getRange('J2').setValue('');
  sheet.getRange('J3').setValue('【原則】');
  sheet.getRange('J4').setValue('✅ 未来の予定される資金移動のみ');
  sheet.getRange('J5').setValue('✅ 家賃/人件費/代表枠/UPSIDER枠など');
  sheet.getRange('J6').setValue('');
  sheet.getRange('J7').setValue('【使い方】');
  sheet.getRange('J8').setValue('サイドバーから「Plan登録」で追加');
  sheet.getRange('J9').setValue('テンプレート登録で繰り返し入力を簡略化');

  sheet.setColumnWidth(10, 280); // J列

  Logger.log('Input_CashPlan 作成完了');
}

/**
 * Calendar シート（日付スパイン）
 * v5.0: 日付の連番を自動生成
 */
function setupCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Calendar');

  if (!sheet) {
    sheet = ss.insertSheet('Calendar');
  }

  if (sheet.getRange('A1').getValue() !== '') {
    Logger.log('Calendar は既に設定済み');
    return;
  }

  // ヘッダー
  sheet.getRange('A1').setValue('日付');
  sheet.getRange('A1').setBackground('#674ea7').setFontColor('#FFFFFF').setFontWeight('bold');

  // 開始日と終了日を設定（過去1年〜未来1年）
  sheet.getRange('C1').setValue('開始日:');
  sheet.getRange('D1').setValue(new Date(2024, 0, 1));
  sheet.getRange('C2').setValue('終了日:');
  sheet.getRange('D2').setValue(new Date(2025, 11, 31));

  // SEQUENCEで日付を自動生成（A2セル）
  const sequenceFormula = `=SEQUENCE(D2-D1+1, 1, D1, 1)`;
  sheet.getRange('A2').setFormula(sequenceFormula);

  // 日付フォーマット
  sheet.getRange('A2:A').setNumberFormat('yyyy-mm-dd');

  sheet.setColumnWidth(1, 120);

  // 説明欄
  sheet.getRange('F1').setValue('📆 日付スパイン（Calendar）');
  sheet.getRange('F1').setFontSize(14).setFontWeight('bold').setFontColor('#674ea7');
  sheet.getRange('F2').setValue('');
  sheet.getRange('F3').setValue('【原則】');
  sheet.getRange('F4').setValue('✅ 日付に欠番なし（連続保証）');
  sheet.getRange('F5').setValue('✅ Daily_Cashで残高を連続表示');

  sheet.setColumnWidth(6, 280);

  Logger.log('Calendar 作成完了');
}

/**
 * Daily_Cash シート（残高連続表示）
 * v5.0: 日次残高の連続表示
 */
function setupDaily_Cash() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Daily_Cash');

  if (!sheet) {
    sheet = ss.insertSheet('Daily_Cash');
  }

  if (sheet.getRange('A1').getValue() !== '') {
    Logger.log('Daily_Cash は既に設定済み');
    return;
  }

  // ヘッダー
  const headers = ['日付', '期首残高', '当日実績', '当日予定', '期末残高', '予定差異', '累計実績', '累計予定'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0b5394');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // A2: Calendarから日付を取得
  sheet.getRange('A2').setFormula('=Calendar!A2:A');

  // 説明欄
  sheet.getRange('J1').setValue('💵 日次残高（Daily_Cash）');
  sheet.getRange('J1').setFontSize(14).setFontWeight('bold').setFontColor('#0b5394');
  sheet.getRange('J2').setValue('');
  sheet.getRange('J3').setValue('【原則】');
  sheet.getRange('J4').setValue('✅ 日付は Calendar から自動取得');
  sheet.getRange('J5').setValue('✅ 実績は DB_Transactions から集計');
  sheet.getRange('J6').setValue('✅ 予定は Input_CashPlan から集計');
  sheet.getRange('J7').setValue('');
  sheet.getRange('J8').setValue('【Phase 2で実装予定】');
  sheet.getRange('J9').setValue('- SUMIF による日別集計');
  sheet.getRange('J10').setValue('- 残高の累積計算');

  sheet.setColumnWidth(10, 280);

  Logger.log('Daily_Cash 作成完了');
}

/**
 * Settings シート（対象月・期首残高）
 * v5.0: 月次表示の基準設定
 */
function setupSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Settings');

  if (!sheet) {
    sheet = ss.insertSheet('Settings');
  }

  if (sheet.getRange('A1').getValue() !== '') {
    Logger.log('Settings は既に設定済み');
    return;
  }

  // ヘッダー
  sheet.getRange('A1').setValue('⚙️ 設定');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#0b5394');

  // 対象月
  sheet.getRange('A3').setValue('対象月:');
  sheet.getRange('B3').setValue(new Date()); // 今月
  sheet.getRange('B3').setNumberFormat('yyyy-mm');

  // 期首残高
  sheet.getRange('A5').setValue('期首残高（全口座合算）:');
  sheet.getRange('B5').setValue(0);
  sheet.getRange('B5').setNumberFormat('#,##0');

  // 口座別期首残高（任意）
  sheet.getRange('A7').setValue('【口座別期首残高】');
  const accountHeaders = ['口座名', '期首残高'];
  sheet.getRange('A8:B8').setValues([accountHeaders]);
  sheet.getRange('A8:B8').setBackground('#0b5394').setFontColor('#FFFFFF').setFontWeight('bold');

  const sampleAccounts = [
    ['みずほ銀行', 1000000],
    ['SBI銀行', 500000],
    ['楽天銀行', 300000],
    ['UPSIDER', 200000]
  ];
  sheet.getRange(9, 1, sampleAccounts.length, 2).setValues(sampleAccounts);

  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);

  // 説明欄
  sheet.getRange('D1').setValue('⚙️ 設定シート');
  sheet.getRange('D1').setFontSize(14).setFontWeight('bold').setFontColor('#0b5394');
  sheet.getRange('D2').setValue('');
  sheet.getRange('D3').setValue('【使い方】');
  sheet.getRange('D4').setValue('1. 対象月を変更すると Month_View が自動更新');
  sheet.getRange('D5').setValue('2. 期首残高は月初の実残高を入力');
  sheet.getRange('D6').setValue('3. 口座別は任意（合算でもOK）');

  sheet.setColumnWidth(4, 280);

  Logger.log('Settings 作成完了');
}

/**
 * Month_View シート（月次資金予実表）
 * v5.0: 日次で実績と予定を表示
 */
function setupMonth_View() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Month_View');

  if (!sheet) {
    sheet = ss.insertSheet('Month_View');
  }

  sheet.clear();

  // ヘッダー
  const headers = [
    '日付',
    '期首残高',
    '実績入金',
    '実績出金',
    '実績純増減',
    '予定入金',
    '予定出金',
    '予定純増減',
    '差異',
    '期末残高',
    '予測残高',
    'メモ'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダースタイル
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0b5394');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setFontSize(11);

  // 数式エリア（A2から開始）
  // A2: 対象月の日付連番を取得（後でARRAYFORMULAで実装）
  sheet.getRange('A2').setFormula('=FILTER(Calendar!A:A, (MONTH(Calendar!A:A)=MONTH(Settings!B3))*(YEAR(Calendar!A:A)=YEAR(Settings!B3)))');

  // B2: 期首残高（1日目はSettings、2日目以降は前日の期末残高）
  sheet.getRange('B2').setFormula('=IF(ROW()=2, Settings!B5, J1)');

  // C2: 実績入金（その日の入金合計）
  sheet.getRange('C2').setFormula('=SUMIFS(DB_Transactions!D:D, DB_Transactions!A:A, A2, DB_Transactions!D:D, ">0", DB_Transactions!F:F, "<>振替")');

  // D2: 実績出金（その日の出金合計）
  sheet.getRange('D2').setFormula('=SUMIFS(DB_Transactions!D:D, DB_Transactions!A:A, A2, DB_Transactions!D:D, "<0", DB_Transactions!F:F, "<>振替")');

  // E2: 実績純増減
  sheet.getRange('E2').setFormula('=C2+D2');

  // F2: 予定入金
  sheet.getRange('F2').setFormula('=SUMIFS(Input_CashPlan!E:E, Input_CashPlan!A:A, A2, Input_CashPlan!E:E, ">0")');

  // G2: 予定出金
  sheet.getRange('G2').setFormula('=SUMIFS(Input_CashPlan!E:E, Input_CashPlan!A:A, A2, Input_CashPlan!E:E, "<0")');

  // H2: 予定純増減
  sheet.getRange('H2').setFormula('=F2+G2');

  // I2: 差異（実績がある場合のみ）
  sheet.getRange('I2').setFormula('=IF(OR(C2<>0, D2<>0), E2-H2, "")');

  // J2: 期末残高（実績ベース）
  sheet.getRange('J2').setFormula('=B2+E2');

  // K2: 予測残高（実績優先、なければ予定）
  sheet.getRange('K2').setFormula('=IF(OR(C2<>0, D2<>0), J2, B2+H2)');

  // 数式を下にコピー（最大31日分）
  const formulaRange = sheet.getRange('B2:K2');
  formulaRange.copyTo(sheet.getRange('B3:K32'), SpreadsheetApp.CopyPasteType.PASTE_FORMULA);

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // 日付
  sheet.setColumnWidth(2, 120);  // 期首残高
  sheet.setColumnWidth(3, 100);  // 実績入金
  sheet.setColumnWidth(4, 100);  // 実績出金
  sheet.setColumnWidth(5, 100);  // 実績純増減
  sheet.setColumnWidth(6, 100);  // 予定入金
  sheet.setColumnWidth(7, 100);  // 予定出金
  sheet.setColumnWidth(8, 100);  // 予定純増減
  sheet.setColumnWidth(9, 100);  // 差異
  sheet.setColumnWidth(10, 120); // 期末残高
  sheet.setColumnWidth(11, 120); // 予測残高
  sheet.setColumnWidth(12, 200); // メモ

  // 数値フォーマット
  sheet.getRange('B:K').setNumberFormat('#,##0');
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');

  // 条件付き書式（残高が0未満で赤）
  const balanceRange = sheet.getRange('J2:K32');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#f4c7c3')
    .setFontColor('#cc0000')
    .setRanges([balanceRange])
    .build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);

  // 説明欄
  sheet.getRange('N1').setValue('💰 月次資金予実表（Month_View）');
  sheet.getRange('N1').setFontSize(14).setFontWeight('bold').setFontColor('#0b5394');
  sheet.getRange('N2').setValue('');
  sheet.getRange('N3').setValue('【原則】');
  sheet.getRange('N4').setValue('✅ 日付は連番（欠番なし）');
  sheet.getRange('N5').setValue('✅ 実績が来たら予定を置き換え');
  sheet.getRange('N6').setValue('✅ 残高が日々繋がる');
  sheet.getRange('N7').setValue('');
  sheet.getRange('N8').setValue('【使い方】');
  sheet.getRange('N9').setValue('1. Settings で対象月を変更');
  sheet.getRange('N10').setValue('2. Source 貼付→統合更新');
  sheet.getRange('N11').setValue('3. 自動で実績が反映される');
  sheet.getRange('N12').setValue('');
  sheet.getRange('N13').setValue('【赤字】');
  sheet.getRange('N14').setValue('残高が0未満 = ショート警告');

  sheet.setColumnWidth(14, 280);

  Logger.log('Month_View 作成完了');
}

/**
 * 振替検出ロジック
 * 同日・同額・逆符号の取引を「振替」としてタグ付け
 */
function detectTransfers() {
  showToast('🔄 振替検出中...', '口座間移動を検出します', 2);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_Transactions');

  if (!sheet) {
    showToast('❌ エラー', 'DB_Transactionsシートがありません', 5);
    return { success: false, message: 'シートが見つかりません' };
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      showToast('⚠️ データなし', '取引データがありません', 5);
      return { success: false, message: 'データがありません' };
    }

    // A列〜F列のデータを取得
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 6);
    const values = dataRange.getValues();

    let transferCount = 0;

    // 各行をチェック
    for (let i = 0; i < values.length; i++) {
      const [date1, account1, desc1, amount1, category1, tag1] = values[i];

      // 既に振替タグが付いている場合はスキップ
      if (tag1 === '振替') continue;

      // 同じ日付で逆符号・同額の取引を探す
      for (let j = i + 1; j < values.length; j++) {
        const [date2, account2, desc2, amount2, category2, tag2] = values[j];

        // 同日、同額（絶対値）、逆符号、異なる口座
        if (
          date1.getTime() === date2.getTime() &&
          Math.abs(amount1) === Math.abs(amount2) &&
          amount1 + amount2 === 0 &&
          account1 !== account2
        ) {
          // 両方に「振替」タグを付ける
          sheet.getRange(i + 2, 6).setValue('振替'); // F列（タグ）
          sheet.getRange(j + 2, 6).setValue('振替');
          transferCount += 2;
          break; // 次の行へ
        }
      }
    }

    showToast('✅ 振替検出完了！', `${transferCount}件の振替を検出しました`, 5);
    Logger.log(`振替検出: ${transferCount}件`);

    return {
      success: true,
      message: '振替検出完了',
      count: transferCount
    };
  } catch (error) {
    showToast('❌ エラー', error.message, 10);
    Logger.log('振替検出エラー: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * 全シート状態確認（v5.0）
 */
function checkAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Source_1', 'Source_2', 'Source_3', 'Source_4', 'Source_5', 'Source_6', 'DB_Transactions', 'DB_Master', 'Input_CashPlan', 'Calendar', 'Settings', 'Month_View'];
  const existingSheets = ss.getSheets().map(sheet => sheet.getName());

  let existCount = 0;
  let missingSheets = [];

  requiredSheets.forEach(sheetName => {
    if (existingSheets.includes(sheetName)) {
      existCount++;
    } else {
      missingSheets.push(sheetName);
    }
  });

  if (missingSheets.length === 0) {
    showToast('✅ 全シート正常', `${existCount}/${requiredSheets.length}シート存在`, 3);
  } else {
    showToast('⚠️ 不足あり', `${missingSheets.length}シート未作成`, 5);
  }

  return {
    total: requiredSheets.length,
    existing: existCount,
    missing: missingSheets
  };
}

/**
 * Month_Viewシートを開く
 */
function openMonthView() {
  switchToSheet('Month_View');
}

/**
 * DB_Transactionsシートを開く
 */
function openTransactions() {
  switchToSheet('DB_Transactions');
}

/**
 * Settingsシートを開く
 */
function openSettings() {
  switchToSheet('Settings');
}

/**
 * シート切り替え
 */
function switchToSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    ss.setActiveSheet(sheet);
    showToast('📄 ' + sheetName, 'シートを切り替えました', 2);
    return { success: true };
  } else {
    showToast('❌ エラー', `${sheetName}が見つかりません`, 3);
    return { success: false };
  }
}

/**
 * ソースシートのメモ一覧を取得
 */
function getSourceMemos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memos = [];

  for (let i = 1; i <= 6; i++) {
    const sheetName = `Source_${i}`;
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const memo = sheet.getRange('K1').getValue() || `Source_${i}（未設定）`;
      memos.push({ number: i, memo: memo });
    }
  }

  return memos;
}

/**
 * DB_Transactionsシートを完全リセット
 * 数式が壊れた場合の緊急復旧用
 */
function resetTransactionsSheet() {
  showToast('🔄 リセット中...', 'DB_Transactionsを再構築します', 2);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DB_Transactions');

  // 既存シートを削除
  if (sheet) {
    ss.deleteSheet(sheet);
    Logger.log('既存のDB_Transactionsを削除しました');
  }

  // 新規作成
  setupDB_Transactions();

  showToast('✅ リセット完了！', '資金台帳が再稼働しました', 5);
  Logger.log('DB_Transactions完全リセット完了');

  return { success: true, message: 'リセット完了' };
}

/**
 * キーワードルールを登録
 * @param {string} keyword - 検索キーワード（正規表現可）
 * @param {string} category - 判定カテゴリ
 * @param {string} detail - 詳細タグ
 * @param {number} priority - 優先度（デフォルト: 10）
 */
function registerKeywordRule(keyword, category, detail, priority = 10) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_Master');

  if (!sheet) {
    showToast('❌ エラー', 'DB_Masterシートが見つかりません', 5);
    return { success: false, message: 'シートが見つかりません' };
  }

  try {
    // 新しい行を追加
    sheet.appendRow([keyword, category, detail, priority]);

    // 優先度でソート（優先度列がある場合）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, 4);
      dataRange.sort([{column: 4, ascending: true}, {column: 1, ascending: true}]);
    }

    showToast('✅ 登録完了！', `キーワード「${keyword}」を追加しました`, 3);
    Logger.log(`キーワードルール登録: ${keyword} → ${category}`);

    return {
      success: true,
      message: '登録完了',
      keyword: keyword,
      category: category
    };
  } catch (error) {
    showToast('❌ エラー', error.message, 5);
    Logger.log('キーワードルール登録エラー: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * プレビュー: キーワードに該当する件数を取得
 * @param {string} keyword - 検索キーワード
 */
function previewKeywordMatch(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_Transactions');

  if (!sheet) {
    return { success: false, count: 0, message: 'シートが見つかりません' };
  }

  try {
    // C列（摘要）のデータを取得
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, count: 0, message: 'データがありません' };
    }

    const descriptionRange = sheet.getRange(2, 3, lastRow - 1, 1); // C列（摘要）
    const descriptions = descriptionRange.getValues();

    // 正規表現でマッチング
    const regex = new RegExp(keyword, 'i');
    let matchCount = 0;

    descriptions.forEach(row => {
      if (row[0] && regex.test(row[0])) {
        matchCount++;
      }
    });

    return {
      success: true,
      count: matchCount,
      message: `${matchCount}件が該当します`
    };
  } catch (error) {
    Logger.log('プレビューエラー: ' + error);
    return { success: false, count: 0, message: error.message };
  }
}

/**
 * 選択中の行の摘要を取得
 */
function getSelectedDescription() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeRange = sheet.getActiveRange();

  if (!activeRange || sheet.getName() !== 'DB_Transactions') {
    return {
      success: false,
      description: '',
      message: 'DB_Transactionsシートで行を選択してください'
    };
  }

  const row = activeRange.getRow();
  if (row < 2) {
    return {
      success: false,
      description: '',
      message: 'データ行を選択してください'
    };
  }

  try {
    // C列（摘要）の値を取得
    const description = sheet.getRange(row, 3).getValue(); // C列

    return {
      success: true,
      description: description || '',
      row: row,
      message: '取得完了'
    };
  } catch (error) {
    Logger.log('摘要取得エラー: ' + error);
    return {
      success: false,
      description: '',
      message: error.message
    };
  }
}

/**
 * 未分類一覧を取得（グルーピング版）
 * サイドバーの「未分類バスター」タブ用
 */
function getUncategorizedTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_Transactions');

  if (!sheet) {
    return { success: false, data: [], message: 'シートが見つかりません' };
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [], message: 'データがありません' };
    }

    // A列〜H列のデータを取得
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 8);
    const values = dataRange.getValues();

    // 科目が「未分類」の行のみフィルタ
    const uncategorized = values
      .map((row, index) => ({
        rowNumber: index + 2,
        date: row[0],
        account: row[1],
        description: row[2],
        amount: row[3],
        category: row[4],
        tag: row[5],
        uid: row[6],
        source: row[7]
      }))
      .filter(item => item.category === '未分類');

    // 摘要でグルーピング
    const grouped = {};
    uncategorized.forEach(item => {
      const key = item.description;
      if (!grouped[key]) {
        grouped[key] = {
          description: key,
          count: 0,
          totalAmount: 0,
          accounts: new Set(),
          firstDate: item.date,
          sample: item
        };
      }
      grouped[key].count++;
      grouped[key].totalAmount += item.amount;
      grouped[key].accounts.add(item.account);
    });

    // 配列に変換して件数順にソート
    const groupedArray = Object.values(grouped)
      .map(g => ({
        description: g.description,
        count: g.count,
        totalAmount: g.totalAmount,
        accounts: Array.from(g.accounts).join(', '),
        firstDate: g.firstDate,
        sample: g.sample
      }))
      .sort((a, b) => b.count - a.count); // 件数が多い順

    return {
      success: true,
      data: groupedArray,
      totalCount: uncategorized.length,
      groupCount: groupedArray.length,
      message: `${uncategorized.length}件の未分類取引が${groupedArray.length}パターンにグルーピングされました`
    };
  } catch (error) {
    Logger.log('未分類取得エラー: ' + error);
    return { success: false, data: [], message: error.message };
  }
}
