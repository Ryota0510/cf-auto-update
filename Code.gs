/**
 * CF自動更新システム v5.0 - Cash Flow管理特化
 * Actual（実績）とPlan（予定）の完全分離
 * 日付スパイン + 残高連続表示
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 CF自動更新')
    .addItem('🚀 システム初期化', 'initializeDatabase')
    .addSeparator()
    .addSubMenu(ui.createMenu('🏦 データ管理')
      .addItem('資金台帳確認', 'refreshTransactions')
      .addItem('予算更新', 'updateBudget')
      .addItem('振替検出', 'detectTransfers')
      .addItem('DB_Transactions再構築', 'resetTransactionsSheet'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 表示切替')
      .addItem('CF表を開く', 'openCF')
      .addItem('CF_Snapshots（残高入力）を開く', 'openCF_Snapshots')
      .addItem('DB_Transactionsを開く', 'openTransactions')
      .addItem('DB_Budgetを開く', 'openBudget'))
    .addSeparator()
    .addItem('📋 全シート状態確認', 'checkAllSheets')
    .addToUi();

  showToast('💰 CF自動更新 v5.4', 'Cash Flow管理 稼働中', 5);
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
    // setupDB_Budget();     // 予算管理（削除：B案では不使用）
    setupInput_CashPlan();   // 予定取引
    setupCF_Snapshots();     // CF_Snapshots（週1残高入力）
    // setupCF();            // CF表（資金予実・日次）※別途手動で設定

    showToast('✅ 初期化完了！', 'Cash Flow管理システムが稼働しました', 5);

    return {
      success: true,
      message: '初期化完了',
      sheets: ['Source_1-6', 'DB_Transactions', 'DB_Master', 'Input_CashPlan', 'CF_Snapshots']
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

  // K1セルに銀行名を自動表示（F2から自動取得）
  const memoCell = sheet.getRange('K1');
  memoCell.setFormula('=IF(F2="", "", REGEXEXTRACT(F2, "(.+?銀行)"))');
  memoCell.setFontSize(14);
  memoCell.setFontWeight('bold');
  memoCell.setFontColor(colors[num - 1]);
  memoCell.setBackground('#fff3e0');
  memoCell.setBorder(true, true, true, true, true, true, '#ff6d00', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 使い方説明（K列以降）
  sheet.getRange('K2').setValue('上記は銀行名です。');
  sheet.getRange('K3').setValue(`💡 使い方`);
  sheet.getRange('K4').setValue('1. MoneyForwardで該当口座を絞り込み');
  sheet.getRange('K5').setValue('2. 全期間を選択してコピー');
  sheet.getRange('K6').setValue('3. A2セル（ヘッダーの下）に貼り付け');
  sheet.getRange('K7').setValue('4. A列は空欄でOK（チェックボックス用）');
  sheet.getRange('K8').setValue('5. 毎回全期間上書きでOK！');
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
  // Step 1: ヘッダー設定（A1:G1）
  // ═══════════════════════════════════════════════════
  const headers = ['日付', '口座', '摘要', '金額', '科目', 'UID', '転記元'];
  sheet.getRange('A1:G1').setValues([headers]);

  // ヘッダースタイル
  const headerRange = sheet.getRange('A1:G1');
  headerRange.setBackground('#0b5394');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setFontSize(11);

  // ═══════════════════════════════════════════════════
  // Step 2: ARRAYFORMULA列の構築（A2-G列）
  // ═══════════════════════════════════════════════════

  // Source_1-6から統合データを取得するための内部シート参照用
  // J2からデータ開始（ヘッダーなし）
  const rawQueryFormula = `=QUERY({Source_1!A2:H; Source_2!A2:H; Source_3!A2:H; Source_4!A2:H; Source_5!A2:H; Source_6!A2:H}, "where Col2 is not null", 0)`;
  sheet.getRange('J2').setFormula(rawQueryFormula);

  // MoneyForwardフォーマット（J列以降）:
  // J列=☑, K列=日付, L列=内容, M列=金額, N列=残高, O列=連携サービス, P列=ステータス, Q列=取引No

  // A2: 日付(整形) - K列（日付）から
  sheet.getRange('A2').setFormula('=ARRAYFORMULA(IF(K2:K="", "", DATEVALUE(LEFT(K2:K, 10))))');
  sheet.getRange('A2:A').setNumberFormat('yyyy/mm/dd');

  // B2: 口座 - O列（連携サービス）から
  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(O2:O="", "", O2:O))');

  // C2: 摘要 - L列（内容）から
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(L2:L="", "", L2:L))');

  // D2: 金額(+/-) - M列（金額）を数値化（入金+/出金-）
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(M2:M="", "", VALUE(REGEXREPLACE(TO_TEXT(M2:M), "[^0-9-]", ""))))');

  // E2: 科目 - 正数（入金）は「入金」、それ以外は自動分類
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(C2:C="", "", IF(D2:D>0, "入金", AUTO_CATEGORIZE(C2:C))))');

  // F2: UID - 口座+取引Noで一意キー生成
  sheet.getRange('F2').setFormula('=ARRAYFORMULA(IF(B2:B="", "", B2:B & "-" & Q2:Q))');

  // G2: 転記元 - 固定値「MF連携」
  sheet.getRange('G2').setFormula('=ARRAYFORMULA(IF(K2:K="", "", "MF連携"))');

  // ═══════════════════════════════════════════════════
  // Step 3: 列幅調整
  // ═══════════════════════════════════════════════════
  sheet.setColumnWidth(1, 100);  // 日付
  sheet.setColumnWidth(2, 150);  // 口座
  sheet.setColumnWidth(3, 250);  // 摘要
  sheet.setColumnWidth(4, 120);  // 金額
  sheet.setColumnWidth(5, 150);  // 科目
  sheet.setColumnWidth(6, 200);  // UID
  sheet.setColumnWidth(7, 100);  // 転記元

  // J列以降は非表示（内部データ）
  sheet.hideColumns(10, 10);

  // ═══════════════════════════════════════════════════
  // Step 4: 条件付き書式（未分類は赤背景）
  // ═══════════════════════════════════════════════════
  const uncategorizedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('未分類')
    .setBackground('#f4c7c3')
    .setFontColor('#cc0000')
    .setRanges([sheet.getRange('E2:E')])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(uncategorizedRule);
  sheet.setConditionalFormatRules(rules);

  // ═══════════════════════════════════════════════════
  // Step 5: 未分類カウンター（T1セル）
  // ═══════════════════════════════════════════════════
  sheet.getRange('T1').setFormula('=IF(COUNTIF(E:E, "未分類")=0, "✅ 全て分類済み", "⚠️ 未分類: " & COUNTIF(E:E, "未分類") & "件")');
  sheet.getRange('T1').setFontSize(14).setFontWeight('bold');

  // 条件付き書式でカウンターの色を変更
  const counterRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('未分類')
    .setFontColor('#cc0000')
    .setRanges([sheet.getRange('T1')])
    .build();

  const counterRuleGreen = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('全て分類済み')
    .setFontColor('#2e7d32')
    .setRanges([sheet.getRange('T1')])
    .build();

  const existingRules = sheet.getConditionalFormatRules();
  existingRules.push(counterRule);
  existingRules.push(counterRuleGreen);
  sheet.setConditionalFormatRules(existingRules);

  sheet.setColumnWidth(20, 200); // T列

  // ═══════════════════════════════════════════════════
  // Step 6: 説明欄
  // ═══════════════════════════════════════════════════
  sheet.getRange('T3').setValue('💰 資金台帳（DB_Transactions）');
  sheet.getRange('T3').setFontSize(14).setFontWeight('bold').setFontColor('#0b5394');
  sheet.getRange('T4').setValue('');
  sheet.getRange('T5').setValue('【原則】');
  sheet.getRange('T6').setValue('✅ 真実は「実際に口座残高が動いた取引」だけ');
  sheet.getRange('T7').setValue('✅ 入金はプラス、出金はマイナスで統一');
  sheet.getRange('T8').setValue('✅ UPSIDERも銀行口座と同格');
  sheet.getRange('T9').setValue('');
  sheet.getRange('T10').setValue('【列の意味】');
  sheet.getRange('T11').setValue('日付: 取引発生日');
  sheet.getRange('T12').setValue('口座: 資金が動いた口座・サービス名');
  sheet.getRange('T13').setValue('摘要: 取引内容');
  sheet.getRange('T14').setValue('金額: 入金+/出金-');
  sheet.getRange('T15').setValue('科目: 自動仕訳（AUTO_CATEGORIZE）');
  sheet.getRange('T16').setValue('UID: 一意キー（重複検知用）');
  sheet.getRange('T17').setValue('転記元: データソース');
  sheet.getRange('T18').setValue('');
  sheet.getRange('T19').setValue('【禁止事項】');
  sheet.getRange('T20').setValue('❌ このシートに直接入力しない');
  sheet.getRange('T21').setValue('❌ 数式を変更しない');

  sheet.setColumnWidth(20, 280); // T列

  Logger.log('DB_Transactions 作成完了（資金台帳 v5.0）');
}

/**
 * 資金台帳データを確認 & 科目を自動更新
 * v5.1: Apps Scriptで科目を一括更新
 */
function refreshTransactions() {
  showToast('🔄 更新中...', '資金台帳を更新します', 2);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transSheet = ss.getSheetByName('DB_Transactions');
  const masterSheet = ss.getSheetByName('DB_Master');

  if (!transSheet) {
    showToast('❌ エラー', 'DB_Transactionsシートがありません', 5);
    return { success: false, message: 'シートが見つかりません' };
  }

  if (!masterSheet) {
    showToast('❌ エラー', 'DB_Masterシートがありません', 5);
    return { success: false, message: 'DB_Masterが見つかりません' };
  }

  try {
    const lastRow = transSheet.getLastRow();

    if (lastRow < 2) {
      showToast('⚠️ データなし', 'Source_1〜6にデータを貼り付けてください', 5);
      return { success: false, message: 'データがありません' };
    }

    // DB_Masterからルールを取得（A列:キーワード, B列:科目）
    const masterLastRow = masterSheet.getLastRow();
    const masterData = masterSheet.getRange(2, 1, masterLastRow - 1, 2).getValues();
    const rules = masterData.filter(row => row[0]); // キーワードがある行のみ

    Logger.log(`ルール数: ${rules.length}`);

    // C列（摘要）を取得
    const descriptions = transSheet.getRange(2, 3, lastRow - 1, 1).getValues();

    // 各行の科目を決定
    const results = descriptions.map(row => {
      const desc = row[0];
      if (!desc) return [''];

      // ルールを上から順にチェック（行順 = 優先度）
      for (let i = 0; i < rules.length; i++) {
        const keyword = rules[i][0];
        const category = rules[i][1];

        if (desc.includes(keyword)) {
          return [category];
        }
      }

      return ['未分類'];
    });

    // E列（科目）に一括書き込み
    transSheet.getRange(2, 5, results.length, 1).setValues(results);

    showToast('✅ 更新完了！', `${lastRow - 1}行の科目を更新しました`, 5);
    Logger.log(`DB_Transactions更新: ${lastRow - 1}行`);

    return {
      success: true,
      message: `${lastRow - 1}行処理完了`,
      rowCount: lastRow - 1
    };
  } catch (error) {
    showToast('❌ エラー', error.message, 10);
    Logger.log('資金台帳更新エラー: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * 予算更新（残日数・1日使用可能額を計算）
 * v5.3: 週1回実行想定
 */
function updateBudget() {
  showToast('🔄 予算更新中...', '残日数と1日使用可能額を計算します', 2);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_Budget');

  if (!sheet) {
    showToast('❌ エラー', 'DB_Budgetシートがありません', 5);
    return { success: false, message: 'シートが見つかりません' };
  }

  try {
    // 今日の日付を取得
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth(); // 0-11

    // 月末日を取得
    const lastDayOfMonth = new Date(year, month + 1, 0);
    const lastDay = lastDayOfMonth.getDate();
    const currentDay = today.getDate();

    // 残日数を計算（今日を含む）
    const remainingDays = lastDay - currentDay + 1;

    Logger.log(`今日: ${year}/${month + 1}/${currentDay}, 月末: ${lastDay}, 残日数: ${remainingDays}`);

    // 各行をチェック（2行目から）
    const lastRow = sheet.getLastRow();
    for (let row = 2; row <= lastRow; row++) {
      const target = sheet.getRange(row, 1).getValue(); // A列（対象）
      const monthlyBudget = sheet.getRange(row, 2).getValue(); // B列（月間予算）

      // 月間予算がある場合のみ計算（UPSIDER・現金）
      if (monthlyBudget && monthlyBudget > 0) {
        // D列: 残日数
        sheet.getRange(row, 4).setValue(remainingDays);

        // E列: 1日使用可能額 = 月間予算 ÷ 残日数
        const dailyBudget = Math.floor(monthlyBudget / remainingDays);
        sheet.getRange(row, 5).setValue(dailyBudget);

        Logger.log(`${target}: 月間予算=${monthlyBudget}, 残日数=${remainingDays}, 1日使用可=${dailyBudget}`);
      } else {
        // 月間予算がない場合はクリア
        sheet.getRange(row, 4).setValue('');
        sheet.getRange(row, 5).setValue('');
      }
    }

    showToast('✅ 予算更新完了！', `残り${remainingDays}日`, 5);
    Logger.log(`予算更新完了: 残日数=${remainingDays}`);

    return {
      success: true,
      message: '予算更新完了',
      remainingDays: remainingDays
    };
  } catch (error) {
    showToast('❌ エラー', error.message, 10);
    Logger.log('予算更新エラー: ' + error);
    return { success: false, message: error.message };
  }
}

/**
 * DB_Master シート（脳みそ）
 * 自動仕訳のルールを管理
 * v5.2: タグ削除、キーワードと科目のみ
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

  // ヘッダー（キーワードと科目のみ）
  const headers = ['検索キーワード', '科目'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#674ea7');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // サンプルデータ（キーワードと科目のみ）
  const sampleData = [
    ['振込手数料', '支払手数料'],
    ['UnivaPay', '売上'],
    ['UPSIDER', '立替金'],
    ['GOOGLE', '広告宣伝費'],
    ['カ）オールエーアイ', '外注費'],
    ['振込＊モカ', '役員報酬'],
    ['PayPay', '売上'],
    ['Amazon', '消耗品費'],
    ['さくら', '通信費'],
    ['Adobe', '通信費'],
    ['みずほ', '支払手数料'],
    ['SBI', '支払手数料'],
    ['楽天', '支払手数料'],
    ['Notion', '通信費'],
    ['GitHub', '通信費'],
    ['AWS', '通信費']
  ];

  sheet.getRange(2, 1, sampleData.length, 2).setValues(sampleData);

  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);

  // 使い方説明
  sheet.getRange('E1').setValue('🧠 自動仕訳の脳みそ');
  sheet.getRange('E1').setFontSize(14).setFontWeight('bold').setFontColor('#674ea7');
  sheet.getRange('E2').setValue('');
  sheet.getRange('E3').setValue('【仕組み】');
  sheet.getRange('E4').setValue('DB_Transactionsの「摘要」列に');
  sheet.getRange('E5').setValue('A列のキーワードが含まれるか検索');
  sheet.getRange('E6').setValue('→ 該当したらB列・C列を自動入力');
  sheet.getRange('E7').setValue('');
  sheet.getRange('E8').setValue('【運用ルール】');
  sheet.getRange('E9').setValue('✅ 上の行ほど優先（行順 = 優先度）');
  sheet.getRange('E10').setValue('✅ 部分一致で検索（前方一致不要）');
  sheet.getRange('E11').setValue('✅ 「未分類」が出たらここに追加');
  sheet.getRange('E12').setValue('✅ 追加した瞬間、自動で反映される');
  sheet.getRange('E13').setValue('');
  sheet.getRange('E14').setValue('⚠️ A列は大文字小文字を区別します');

  // 列幅調整
  sheet.setColumnWidth(5, 280); // E列

  Logger.log('DB_Master 作成完了（脳みそ v5.1）');
}

/**
 * DB_Budget シート（予算管理）
 * v5.3: UPSIDER・現金の月間予算管理
 */
function setupDB_Budget() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DB_Budget');

  if (!sheet) {
    sheet = ss.insertSheet('DB_Budget');
  }

  if (sheet.getRange('A1').getValue() !== '') {
    Logger.log('DB_Budget は既に設定済み');
    return;
  }

  // ヘッダー
  const headers = ['科目', '月間予算', '実残高（MF転記）', '残日数', '1日使用可能額'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#27ae60');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // サンプルデータ
  const sampleData = [
    ['UPSIDER', 500000, 450000, '', ''], // 残日数・1日使用可能額は自動計算
    ['現金', 300000, 280000, '', ''],
    ['みずほ銀行', '', 1200000, '', ''],  // 月間予算なし、実残高のみ
    ['SBI銀行', '', 800000, '', ''],
    ['楽天銀行', '', 500000, '', '']
  ];

  sheet.getRange(2, 1, sampleData.length, 5).setValues(sampleData);

  // 列幅調整
  sheet.setColumnWidth(1, 150);  // 科目
  sheet.setColumnWidth(2, 120);  // 月間予算
  sheet.setColumnWidth(3, 150);  // 実残高
  sheet.setColumnWidth(4, 100);  // 残日数
  sheet.setColumnWidth(5, 150);  // 1日使用可能額

  // 数値フォーマット
  sheet.getRange('B:C').setNumberFormat('#,##0');
  sheet.getRange('E:E').setNumberFormat('#,##0');

  // 条件付き書式（1日使用可能額が1万円未満で警告）
  const warningRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(10000)
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([sheet.getRange('E2:E6')])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(warningRule);
  sheet.setConditionalFormatRules(rules);

  // 使い方説明
  sheet.getRange('G1').setValue('💰 予算管理（DB_Budget）');
  sheet.getRange('G1').setFontSize(14).setFontWeight('bold').setFontColor('#27ae60');
  sheet.getRange('G2').setValue('');
  sheet.getRange('G3').setValue('【原則】');
  sheet.getRange('G4').setValue('✅ UPSIDERと現金は月間予算で管理');
  sheet.getRange('G5').setValue('✅ 銀行口座は実残高のみ記録');
  sheet.getRange('G6').setValue('✅ 実残高はMFから週1回転記');
  sheet.getRange('G7').setValue('');
  sheet.getRange('G8').setValue('【運用ルール】');
  sheet.getRange('G9').setValue('1. B列（月間予算）: UPSIDER・現金のみ入力');
  sheet.getRange('G10').setValue('2. C列（実残高）: 全口座、MFから転記');
  sheet.getRange('G11').setValue('3. メニューから「予算更新」実行');
  sheet.getRange('G12').setValue('4. D列（残日数）・E列（1日使用可）自動計算');
  sheet.getRange('G13').setValue('');
  sheet.getRange('G14').setValue('【計算式】');
  sheet.getRange('G15').setValue('残日数 = 月末日 - 今日 + 1');
  sheet.getRange('G16').setValue('1日使用可能額 = 月間予算 ÷ 残日数');

  sheet.setColumnWidth(7, 300); // G列

  Logger.log('DB_Budget 作成完了');
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

  // ヘッダー（5列に簡素化）
  const headers = ['予定日', '科目', '予定金額', '種別', 'メモ'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#e67e22');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // サンプルデータ（月次枠と単発の例）
  const sampleData = [
    [new Date(2025, 0, 1), 'UPSIDER枠', 500000, '月次枠', '月間予算'],
    [new Date(2025, 0, 1), '現金経費', 300000, '月次枠', '月間予算'],
    [new Date(2025, 0, 25), '家賃', 200000, '単発', ''],
    [new Date(2025, 0, 31), '人件費', 300000, '単発', '給与振込']
  ];

  sheet.getRange(2, 1, sampleData.length, 5).setValues(sampleData);
  sheet.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('C2:C').setNumberFormat('#,##0');

  // B列（科目）：CF_Snapshots!K4:K からドロップダウン
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRange('CF_Snapshots!K4:K'), true)
    .setAllowInvalid(false)
    .setHelpText('科目一覧から選択してください')
    .build();
  sheet.getRange('B2:B').setDataValidation(categoryRule);

  // D列（種別）：単発 or 月次枠 のみ
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['単発', '月次枠'], true)
    .setAllowInvalid(false)
    .setHelpText('「単発」または「月次枠」を選択')
    .build();
  sheet.getRange('D2:D').setDataValidation(typeRule);

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // 予定日
  sheet.setColumnWidth(2, 150);  // 科目
  sheet.setColumnWidth(3, 120);  // 予定金額
  sheet.setColumnWidth(4, 100);  // 種別
  sheet.setColumnWidth(5, 200);  // メモ

  // 説明欄
  sheet.getRange('J1').setValue('📅 予定取引（Input_CashPlan）');
  sheet.getRange('J1').setFontSize(14).setFontWeight('bold').setFontColor('#e67e22');
  sheet.getRange('J2').setValue('');
  sheet.getRange('J3').setValue('【種別：月次枠】');
  sheet.getRange('J4').setValue('・予定日=その月の1日（例：2025/11/01）');
  sheet.getRange('J5').setValue('・予定金額=月間予算（正数）');
  sheet.getRange('J6').setValue('・科目=枠の科目（UPSIDER枠、現金経費など）');
  sheet.getRange('J7').setValue('→ CF表で日割り展開され、端数は月末に寄せられます');
  sheet.getRange('J8').setValue('');
  sheet.getRange('J9').setValue('【種別：単発】');
  sheet.getRange('J10').setValue('・特定日の支出/入金（家賃、人件費など）');
  sheet.getRange('J11').setValue('・予定金額は正数=出金、負数=入金');

  sheet.setColumnWidth(10, 320); // J列

  Logger.log('Input_CashPlan 作成完了');
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
 * 全シート状態確認（v5.4）
 */
function checkAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Source_1', 'Source_2', 'Source_3', 'Source_4', 'Source_5', 'Source_6', 'DB_Transactions', 'DB_Master', 'Input_CashPlan', 'CF_Snapshots'];
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
 * DB_Transactionsシートを開く
 */
function openTransactions() {
  switchToSheet('DB_Transactions');
}

/**
 * DB_Budgetシートを開く
 */
function openBudget() {
  switchToSheet('DB_Budget');
}

/**
 * CF表を開く
 */
function openCF() {
  switchToSheet('CF');
}

/**
 * CF_Snapshotsを開く
 */
function openCF_Snapshots() {
  switchToSheet('CF_Snapshots');
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
 * @param {string} category - 科目
 * @param {string} detail - 詳細タグ
 */
function registerKeywordRule(keyword, category, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_Master');

  if (!sheet) {
    showToast('❌ エラー', 'DB_Masterシートが見つかりません', 5);
    return { success: false, message: 'シートが見つかりません' };
  }

  try {
    // 新しい行を最後に追加（行順 = 優先度）
    sheet.appendRow([keyword, category, detail]);

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
 * カスタム関数: 自動分類
 * E2セルに =AUTO_CATEGORIZE(C2:C) と入力
 *
 * @param {Array} descriptionRange - C列（摘要）の範囲
 * @return {Array} 科目の1次元配列
 * @customfunction
 */
function AUTO_CATEGORIZE(descriptionRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('DB_Master');

  if (!masterSheet) {
    return descriptionRange.map(() => ['エラー: DB_Masterなし']);
  }

  // DB_Masterからルールを取得（A列:キーワード, B列:科目）
  const masterLastRow = masterSheet.getLastRow();
  if (masterLastRow < 2) {
    return descriptionRange.map(() => ['未分類']);
  }

  const rules = masterSheet.getRange(2, 1, masterLastRow - 1, 2).getValues()
    .filter(row => row[0]); // キーワードがある行のみ

  // 各摘要を処理
  return descriptionRange.map(row => {
    const desc = row[0];
    if (!desc) return [''];

    // ルールを上から順にチェック（行順 = 優先度）
    for (const [keyword, category] of rules) {
      if (desc.toString().includes(keyword.toString())) {
        return [category || '未分類'];
      }
    }

    return ['未分類'];
  });
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
 * DB_Masterから科目一覧を取得
 * v5.1: 未分類取引のドロップダウン表示用
 */
function getAllCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DB_Master');

  if (!sheet) {
    return { success: false, categories: [] };
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, categories: [] };
    }

    // B列（科目）のデータを取得
    const categoryRange = sheet.getRange(2, 2, lastRow - 1, 1);
    const categories = categoryRange.getValues()
      .map(row => row[0])
      .filter(cat => cat !== '') // 空白除外
      .filter((cat, index, self) => self.indexOf(cat) === index); // 重複除外

    return {
      success: true,
      categories: categories.sort() // アルファベット順ソート
    };
  } catch (error) {
    Logger.log('科目一覧取得エラー: ' + error);
    return { success: false, categories: [] };
  }
}

/**
 * CF_Snapshots シート作成（週1残高入力専用）
 */
function setupCF_Snapshots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('CF_Snapshots');

  if (!sheet) {
    sheet = ss.insertSheet('CF_Snapshots');
  }

  // 既に設定済みの場合はスキップ
  if (sheet.getRange('A1').getValue() === '💰 週1残高入力') {
    Logger.log('CF_Snapshots は既に設定済み');
    return;
  }

  sheet.clear();

  sheet.getRange('A1').setValue('💰 週1残高入力（6口座）');
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor('#0b5394');

  // ヘッダー行（A3:I3）- B列に合計を追加
  const snapshotHeaders = ['入力日', '合計', 'Source_1', 'Source_2', 'Source_3', 'Source_4', 'Source_5', 'Source_6', 'メモ'];
  sheet.getRange(3, 1, 1, 9).setValues([snapshotHeaders]);

  const snapshotHeaderRange = sheet.getRange(3, 1, 1, 9);
  snapshotHeaderRange.setBackground('#34a853');
  snapshotHeaderRange.setFontColor('#FFFFFF');
  snapshotHeaderRange.setFontWeight('bold');
  snapshotHeaderRange.setHorizontalAlignment('center');

  // C〜H列のヘッダは Source_1〜6!K1 を参照（口座名を自動表示）
  for (let i = 1; i <= 6; i++) {
    sheet.getRange(3, i + 2).setFormula(`=IFERROR(Source_${i}!K1, "Source_${i}")`);
  }

  // サンプルデータ（1行）
  const sampleSnapshot = [
    [new Date(), '', 1200000, 800000, 500000, 0, 0, 0, '初期残高']
  ];
  sheet.getRange(4, 1, 1, 9).setValues(sampleSnapshot);
  sheet.getRange('A4').setNumberFormat('yyyy/mm/dd');

  // B列: ARRAYFORMULAで合計を自動計算（A列に日付があれば C〜H列を合計）
  sheet.getRange('B4').setFormula('=ARRAYFORMULA(IF(A4:A="", "", C4:C+D4:D+E4:E+F4:F+G4:G+H4:H))');
  sheet.getRange('B:B').setNumberFormat('#,##0');
  sheet.getRange('C4:H4').setNumberFormat('#,##0');

  // K列：科目一覧（DB_Master, Input_CashPlanから統合）
  sheet.getRange('K3').setValue('科目一覧');
  sheet.getRange('K3').setFontWeight('bold').setBackground('#34a853').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  // K4: 全シートから科目を取得してソート・ユニーク化（DB_MasterとInput_CashPlanのみ）
  const categoryFormula = '=SORT(UNIQUE(FILTER({DB_Master!B2:B; Input_CashPlan!B2:B}, {DB_Master!B2:B; Input_CashPlan!B2:B}<>"" )))';
  sheet.getRange('K4').setFormula(categoryFormula);

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // A列：入力日
  sheet.setColumnWidth(2, 120);  // B列：合計
  for (let i = 3; i <= 8; i++) {
    sheet.setColumnWidth(i, 100); // C〜H列：Source_1〜6残高
  }
  sheet.setColumnWidth(9, 150);  // I列：メモ
  sheet.setColumnWidth(11, 120); // K列：科目一覧

  Logger.log('CF_Snapshots シート作成完了');
}
