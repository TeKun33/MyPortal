/**
 * ==========================================================================
 * マイポータル (MyPortal) - Google Apps Script サーバーサイドバックエンド
 * ==========================================================================
 *
 * 【システム概要】
 * 本スクリプトは、個人・組織向けの統合ポータルWebアプリケーション「マイポータル」の
 * サーバー側バックエンドロジック (Google Apps Script) を担当します。
 * スプレッドシートをデータベースとして活用し、Google Workspace API (Gmail, Calendar, Tasks, Drive)
 * と連携して各種機能を提供します。
 *
 * 【主な機能モジュール】
 * 1. Webアプリの配信 (doGet による HTML/SPA のレンダリング)
 * 2. ログ管理・監査機能 (__LOG__ シートへの操作履歴記録)
 * 3. ユーザー設定・個人シート管理 (テーマ、ダークモード等の永続化)
 * 4. 利用規約同意管理 (__TERMS__ シートおよび各ユーザーシートによる規約同意制御)
 * 5. ブックマーク管理 (CRUD処理、アイコン設定、ドラッグ＆ドロップ並び替え)
 * 6. Gmail 連携 (受信トレイの最新メール一覧取得・プレビュー)
 * 7. Google カレンダー連携 (週単位の予定取得、終日・時間指定予定の追加/編集/削除)
 * 8. Google Tasks 連携 (タスク一覧取得、完了チェック、タスクの追加/編集/削除/リスト間移動)
 * 9. QRコード Drive 保存 (生成したQRコード画像のGoogle Driveへの保存)
 *
 * 【スプレッドシート構成】
 * - [__LOG__]   : システム全体の操作ログ（日時, ユーザー, 操作, 詳細）
 * - [__TERMS__] : 全ユーザーの利用規約同意ステータス一覧（ユーザー, 同意ステータス, 同意日時, 最終更新日時）
 * - [メールアドレス名シート] : 各ユーザーの専用シート
 *     - Row 1   : 設定情報 (__SETTINGS__, theme, ocean, darkMode, false, termsAgreed, false, termsAgreedAt, '')
 *     - Row 2   : ブックマークヘッダー (__BOOKMARKS__, タイトル, URL, カテゴリ, アイコン, 追加日時)
 *     - Row 3〜 : 登録済みブックマークデータ
 * ==========================================================================
 */


/* ==========================================================================
 * 1. Webアプリケーション初期化・エントリーポイント
 * ========================================================================== */

/**
 * Webアプリケーションアクセス時の初期表示処理 (HTTP GET エントリーポイント)
 * - HTMLテンプレート 'index.html' を読み込んで評価し、Webページとして出力します。
 * - iframe内での表示を許可 (setXFrameOptionsMode.ALLOWALL) しています。
 *
 * @param {Object} e - HTTP GETリクエストのイベントオブジェクト
 * @returns {HtmlOutput} レンダリングされたHTMLページ、またはエラーメッセージ
 */
function doGet(e) {
  try {
    // 1. プロジェクト内の index.html テンプレートを読み込み
    const template = HtmlService.createTemplateFromFile('index');
    
    // 2. テンプレートを評価してHtmlOutputを生成し、ページタイトルとiframe表示許可を設定
    return template.evaluate()
      .setTitle('マイポータル')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    // テンプレート評価時などの例外発生時はエラー画面を返却
    return HtmlService.createHtmlOutput('エラーが発生しました: ' + error.message);
  }
}


/* ==========================================================================
 * 2. ログ管理・監査機能 (__LOG__ シート) & 管理用シート初期化
 * ========================================================================== */

/**
 * 監査ログ用のシート (__LOG__) を取得する。存在しない場合は新規作成して初期化する。
 * - カラム構成: A=日時, B=ユーザー, C=操作, D=詳細
 *
 * @returns {Sheet} ログシートオブジェクト
 */
function getOrCreateLogSheet() {
  // スクリプトのプロパティストアからスプレッドシートID (SSID) を取得
  const SSID = PropertiesService.getScriptProperties().getProperty('SSID');
  const SS = SpreadsheetApp.openById(SSID);
  
  // '__LOG__' という名称のシートを検索
  let sheet = SS.getSheetByName('__LOG__');
  
  // シートが存在しない場合は新規作成し、ヘッダーと見やすい列幅を設定
  if (!sheet) {
    sheet = SS.insertSheet('__LOG__');
    // 1行目にヘッダーラベルを設定
    sheet.getRange(1, 1, 1, 4).setValues([['日時', 'ユーザー', '操作', '詳細']]);
    sheet.setFrozenRows(1); // スクロール時も見出しが見えるよう1行目を固定
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold'); // ヘッダーを太字に設定
    sheet.setColumnWidth(1, 180); // 日時列幅
    sheet.setColumnWidth(2, 200); // ユーザー(メールアドレス)列幅
    sheet.setColumnWidth(3, 150); // 操作カテゴリ列幅
    sheet.setColumnWidth(4, 400); // 詳細メッセージ列幅
  }
  return sheet;
}

/**
 * 利用規約同意状況シート (__TERMS__) を取得する。存在しない場合は新規作成する。
 * - 全ユーザーの規約同意状態を一元管理するためのシートです。
 * - カラム構成: A=ユーザー, B=同意ステータス, C=同意日時, D=最終更新日時
 *
 * @returns {Sheet} 利用規約シートオブジェクト
 */
function getOrCreateTermsSheet() {
  const SSID = PropertiesService.getScriptProperties().getProperty('SSID');
  const SS = SpreadsheetApp.openById(SSID);
  
  let sheet = SS.getSheetByName('__TERMS__');
  if (!sheet) {
    sheet = SS.insertSheet('__TERMS__');
    sheet.getRange(1, 1, 1, 4).setValues([['ユーザー', '同意ステータス', '同意日時', '最終更新日時']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setColumnWidth(1, 220); // ユーザー列
    sheet.setColumnWidth(2, 120); // 同意ステータス列
    sheet.setColumnWidth(3, 180); // 同意日時列
    sheet.setColumnWidth(4, 180); // 最終更新日時列
  }
  return sheet;
}

/**
 * 操作履歴を監査ログ (__LOG__ シート) に1行追加する
 *
 * @param {string} email - 操作を行ったユーザーのメールアドレス
 * @param {string} action - 操作のカテゴリ名 (例: 'ブックマーク', 'カレンダー', '設定')
 * @param {string} detail - 操作の詳細内容
 */
function writeLog(email, action, detail) {
  try {
    const sheet = getOrCreateLogSheet();
    // スクリプトのタイムゾーンに合わせて現在日時をフォーマット (yyyy/MM/dd HH:mm:ss)
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    // シートの末尾に行を追加
    sheet.appendRow([timestamp, email || '', action || '', detail || '']);
  } catch (e) {
    // ログ記録の失敗によりユーザーのメイン処理が中断しないよう、コンソール出力に留める
    console.error('writeLog Error:', e);
  }
}


/* ==========================================================================
 * 3. ユーザー情報・個別シート & 設定管理
 * ========================================================================== */

/**
 * 現在ログインしているユーザーのメールアドレスを取得する
 *
 * @returns {string} ユーザーのメールアドレス
 */
function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    console.error('getUserEmail Error:', error);
    throw error;
  }
}

/**
 * ユーザー専用のシートを取得する。存在しない場合は新規作成してヘッダー行をセットアップする。
 * - シート名: ユーザーのメールアドレス
 * - Row 1 : 設定情報 (__SETTINGS__, theme, ocean, darkMode, false, termsAgreed, false, termsAgreedAt, '')
 * - Row 2 : ブックマークヘッダー (__BOOKMARKS__, タイトル, URL, カテゴリ, アイコン, 追加日時)
 * - Row 3〜: 個別ブックマークデータ
 *
 * @param {string} email - ユーザーのメールアドレス
 * @returns {Sheet} ユーザー用シートオブジェクト
 */
function getOrCreateUserSheet(email) {
  try {
    const SSID = PropertiesService.getScriptProperties().getProperty('SSID');
    const SS = SpreadsheetApp.openById(SSID);
    let sheet = SS.getSheetByName(email);
    
    // シートがまだ存在しない場合は作成し、初期フォーマットをセットアップ
    if (!sheet) {
      sheet = SS.insertSheet(email);
      // Row 1: 設定情報（テーマ名、ダークモードフラグ、規約同意フラグ、同意日時）
      sheet.getRange(1, 1, 1, 9).setValues([['__SETTINGS__', 'theme', 'ocean', 'darkMode', 'false', 'termsAgreed', 'false', 'termsAgreedAt', '']]);
      // Row 2: ブックマークテーブルの見出し行
      sheet.getRange(2, 1, 1, 6).setValues([['__BOOKMARKS__', 'タイトル', 'URL', 'カテゴリ', 'アイコン', '追加日時']]);
    }
    
    return sheet;
  } catch (error) {
    console.error('getOrCreateUserSheet Error:', error);
    throw error;
  }
}

/**
 * ユーザーの各種設定 (テーマ, ダークモード, 利用規約同意状況) を取得する
 *
 * @param {string} email - ユーザーのメールアドレス
 * @returns {Object} { theme: string, darkMode: boolean, termsAgreed: boolean, termsAgreedAt: string }
 */
function getUserSettings(email) {
  try {
    const sheet = getOrCreateUserSheet(email);
    // Row 1 の各セルから設定値を一括取得
    const settingsRange = sheet.getRange(1, 1, 1, 9).getValues()[0];
    
    return {
      theme: settingsRange[2] || 'ocean',                                   // 列C: テーマ名 (デフォルト: 'ocean')
      darkMode: settingsRange[4] === 'true' || settingsRange[4] === true,   // 列E: ダークモードフラグ
      termsAgreed: settingsRange[6] === 'true' || settingsRange[6] === true,// 列G: 規約同意フラグ
      termsAgreedAt: settingsRange[8] ? String(settingsRange[8]) : ''       // 列I: 規約同意日時
    };
  } catch (error) {
    console.error('getUserSettings Error:', error);
    // エラー時は安全なデフォルト設定を返す
    return { theme: 'ocean', darkMode: false, termsAgreed: false, termsAgreedAt: '' };
  }
}

/**
 * ユーザーの設定情報 (テーマ, ダークモード, 規約同意) を保存する
 *
 * @param {string} email - ユーザーのメールアドレス
 * @param {Object} settings - 保存する設定オブジェクト { theme, darkMode, termsAgreed }
 * @returns {Object} { success: boolean, error?: string }
 */
function saveUserSettings(email, settings) {
  try {
    const sheet = getOrCreateUserSheet(email);
    
    // テーマ設定の更新 (セル C1)
    if (settings.theme !== undefined) {
      sheet.getRange('C1').setValue(settings.theme);
    }
    // ダークモード設定の更新 (セル E1)
    if (settings.darkMode !== undefined) {
      sheet.getRange('E1').setValue(settings.darkMode.toString());
    }
    // 利用規約同意フラグの更新 (セル F1〜I1)
    if (settings.termsAgreed !== undefined) {
      sheet.getRange('F1').setValue('termsAgreed');
      sheet.getRange('G1').setValue(settings.termsAgreed ? 'true' : 'false');
      sheet.getRange('H1').setValue('termsAgreedAt');
      if (settings.termsAgreed) {
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
        sheet.getRange('I1').setValue(timestamp);
      }
    }
    
    // 変更履歴をログに記録
    writeLog(email, '設定', '設定の変更をしました');
    return { success: true };
  } catch (error) {
    console.error('saveUserSettings Error:', error);
    return { success: false, error: error.message };
  }
}


/* ==========================================================================
 * 4. 利用規約同意管理機能 (__TERMS__ シート & ガード処理)
 * ========================================================================== */

/**
 * 利用規約の同意状態をスプレッドシートに保存する
 * 1. ユーザー専用シートの Row 1 (設定行) を更新
 * 2. __TERMS__ シートの該当ユーザー行を更新または追加
 * 3. __LOG__ シートに監査ログを記録
 *
 * @param {string} email - ユーザーのメールアドレス (空の場合は現在のアカウント)
 * @param {boolean|string} agreed - 同意フラグ (true / false)
 * @returns {Object} { success: boolean, termsAgreed: boolean, termsAgreedAt: string }
 */
function saveTermsAgreement(email, agreed) {
  try {
    if (!email) email = getUserEmail();
    const isAgreed = agreed === true || agreed === 'true';
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    
    // 1. ユーザー専用シートの Row 1 を更新
    const userSheet = getOrCreateUserSheet(email);
    userSheet.getRange('F1').setValue('termsAgreed');
    userSheet.getRange('G1').setValue(isAgreed ? 'true' : 'false');
    userSheet.getRange('H1').setValue('termsAgreedAt');
    if (isAgreed) {
      userSheet.getRange('I1').setValue(timestamp);
    }
    
    // 2. __TERMS__ シートの該当ユーザー行を更新または新規追加
    const termsSheet = getOrCreateTermsSheet();
    const data = termsSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        rowIndex = i + 1; // 1-based の行番号
        break;
      }
    }
    
    const statusText = isAgreed ? '同意済み' : '未同意';
    if (rowIndex > 0) {
      // 既存行を更新
      termsSheet.getRange(rowIndex, 2).setValue(statusText);
      if (isAgreed) {
        termsSheet.getRange(rowIndex, 3).setValue(timestamp);
      }
      termsSheet.getRange(rowIndex, 4).setValue(timestamp);
    } else {
      // 新規行を追加
      termsSheet.appendRow([email, statusText, isAgreed ? timestamp : '', timestamp]);
    }
    
    // 3. __LOG__ シートに監査ログを記録
    writeLog(email, '利用規約', isAgreed ? '初回利用規約に同意しました' : '利用規約の同意状態を更新しました');
    
    return { success: true, termsAgreed: isAgreed, termsAgreedAt: isAgreed ? timestamp : '' };
  } catch (error) {
    console.error('saveTermsAgreement Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * ユーザーの利用規約同意状態を取得する
 *
 * @param {string} [email] - ユーザーのメールアドレス (省略時は現在のアカウント)
 * @returns {Object} { termsAgreed: boolean, termsAgreedAt: string }
 */
function getTermsAgreement(email) {
  try {
    if (!email) email = getUserEmail();
    const settings = getUserSettings(email);
    return {
      termsAgreed: settings.termsAgreed || false,
      termsAgreedAt: settings.termsAgreedAt || ''
    };
  } catch (error) {
    console.error('getTermsAgreement Error:', error);
    return { termsAgreed: false, termsAgreedAt: '' };
  }
}

/**
 * 【セキュリティガード】ユーザーが利用規約に同意しているかを検証する
 * 未同意の場合はエラーをスローし、各APIへの不正・未承認アクセスを遮断します。
 *
 * @param {string} [email] - ユーザーのメールアドレス (省略時は実行中アカウント)
 * @throws {Error} 利用規約に未同意の場合にエラーをスロー
 */
function assertTermsAgreed(email) {
  if (!email) email = getUserEmail();
  const agreement = getTermsAgreement(email);
  if (!agreement.termsAgreed) {
    throw new Error('利用規約に同意していないため、この操作は許可されていません。');
  }
}


/* ==========================================================================
 * 5. ブックマーク管理機能 (CRUD & 並び替え)
 * ========================================================================== */

/**
 * ユーザーの登録済みブックマーク一覧を取得する (3行目以降)
 *
 * @param {string} email - ユーザーのメールアドレス
 * @returns {Array<Object>} ブックマークオブジェクトの配列 [{ index, title, url, category, icon, addedAt }]
 */
function getBookmarks(email) {
  try {
    assertTermsAgreed(email); // 利用規約チェック
    const sheet = getOrCreateUserSheet(email);
    const lastRow = sheet.getLastRow();
    
    // 2行以下（ヘッダーのみまたは空）の場合はデータなし
    if (lastRow <= 2) {
      return [];
    }
    
    // Row 3 以降の全データを一括取得 (B列:タイトル 〜 F列:追加日時 の5列)
    const dataRange = sheet.getRange(3, 2, lastRow - 2, 5).getValues();
    const bookmarks = [];
    
    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      // タイトルもURLもない空行はスキップ
      if (!row[0] && !row[1]) continue;
      
      bookmarks.push({
        index: i,
        title: row[0],
        url: row[1],
        category: row[2],
        icon: row[3],
        addedAt: row[4]
      });
    }
    
    return bookmarks;
  } catch (error) {
    console.error('getBookmarks Error:', error);
    return [];
  }
}

/**
 * 新しいブックマークを追加する (シート末尾に行を追加)
 *
 * @param {string} email - ユーザーのメールアドレス
 * @param {string} title - タイトル
 * @param {string} url - リンク先URL
 * @param {string} category - カテゴリ名
 * @param {string} [icon] - Material Symbols アイコン名
 * @returns {Object} { success: boolean, error?: string }
 */
function addBookmark(email, title, url, category, icon) {
  try {
    assertTermsAgreed(email); // 利用規約チェック
    const sheet = getOrCreateUserSheet(email);
    
    // アイコン名の補正処理（指定がない場合はカテゴリ値またはデフォルトアイコン 'language'）
    const iconName = icon || (category && !category.includes('/') ? category : 'language');
    const cat = (icon && category) ? category : '';
    const addedAt = new Date().toISOString();
    
    // シート末尾に行を追加 (A列空欄, B:タイトル, C:URL, D:カテゴリ, E:アイコン, F:追加日時)
    sheet.appendRow(['', title, url, cat, iconName, addedAt]);
    
    writeLog(email, 'ブックマーク', '新しいブックマークを追加しました');
    return { success: true };
  } catch (error) {
    console.error('addBookmark Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 既存のブックマークを更新（編集）する
 *
 * @param {string} email - ユーザーのメールアドレス
 * @param {number} index - 更新対象のインデックス (0始まり)
 * @param {string} title - タイトル
 * @param {string} url - リンク先URL
 * @param {string} [category] - カテゴリ名
 * @param {string} [icon] - アイコン名
 * @returns {Object} { success: boolean, error?: string }
 */
function updateBookmark(email, index, title, url, category, icon) {
  try {
    assertTermsAgreed(email); // 利用規約チェック
    const sheet = getOrCreateUserSheet(email);
    const rowToUpdate = index + 3; // 3行目からデータ開始のため +3
    
    // 更新対象行が有効範囲内か検証
    if (rowToUpdate < 3 || rowToUpdate > sheet.getLastRow()) {
      return { success: false, error: '指定されたブックマークが見つかりません' };
    }

    const iconName = icon || (category && !category.includes('/') ? category : 'language');
    const cat = (icon && category) ? category : '';

    // B列〜E列 (タイトル, URL, カテゴリ, アイコン) を上書き更新
    sheet.getRange(rowToUpdate, 2, 1, 4).setValues([[
      title,
      url,
      cat,
      iconName
    ]]);
    
    writeLog(email, 'ブックマーク', 'ブックマークを更新しました');
    return { success: true };
  } catch (error) {
    console.error('updateBookmark Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * ブックマークを削除する
 *
 * @param {string} email - ユーザーのメールアドレス
 * @param {number} index - 削除対象のインデックス (0始まり)
 * @returns {Object} { success: boolean, error?: string }
 */
function deleteBookmark(email, index) {
  try {
    assertTermsAgreed(email); // 利用規約チェック
    const sheet = getOrCreateUserSheet(email);
    const rowToDelete = index + 3;
    
    // 対象行を削除
    if (rowToDelete > 2 && rowToDelete <= sheet.getLastRow()) {
      sheet.deleteRow(rowToDelete);
    }
    
    writeLog(email, 'ブックマーク', 'ブックマークを削除しました');
    return { success: true };
  } catch (error) {
    console.error('deleteBookmark Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 全ブックマークを並び順通りに一括上書き保存する (ドラッグ＆ドロップ並び替え用)
 *
 * @param {string} email - ユーザーのメールアドレス
 * @param {Array<Object>} bookmarks - 新しい順序に並んだブックマークオブジェクト配列
 * @returns {Object} { success: boolean, error?: string }
 */
function saveAllBookmarks(email, bookmarks) {
  try {
    assertTermsAgreed(email); // 利用規約チェック
    const sheet = getOrCreateUserSheet(email);
    const lastRow = sheet.getLastRow();
    
    // 3行目以降の既存ブックマークデータをすべてクリア
    if (lastRow >= 3) {
      sheet.getRange(3, 1, lastRow - 2, 6).clearContent();
    }
    
    // データが空の場合はクリアのみで正常終了
    if (!bookmarks || bookmarks.length === 0) {
      return { success: true };
    }
    
    // 新しい並び順の2次元配列データを構築
    const rows = bookmarks.map(bm => {
      const icon = bm.icon || 'language';
      const addedAt = bm.addedAt || new Date().toISOString();
      return ['', bm.title || '', bm.url || '', bm.category || '', icon, addedAt];
    });
    
    // 3行目から一括書き込み
    sheet.getRange(3, 1, rows.length, 6).setValues(rows);
    writeLog(email, 'ブックマーク', 'ブックマークの並び順を更新しました');
    return { success: true };
  } catch (error) {
    console.error('saveAllBookmarks Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 互換用: インデックス配列に基づいてブックマークを並び替える
 *
 * @param {string} email - ユーザーのメールアドレス
 * @param {Array<number>} newOrder - 新しい並び順のインデックス配列
 * @returns {Object} { success: boolean, error?: string }
 */
function reorderBookmarks(email, newOrder) {
  try {
    assertTermsAgreed(email);
    const bookmarks = getBookmarks(email);
    if (bookmarks.length === 0) return { success: true };
    
    const reordered = [];
    for (let i = 0; i < newOrder.length; i++) {
      const idx = newOrder[i];
      if (bookmarks[idx]) {
        reordered.push(bookmarks[idx]);
      }
    }
    bookmarks.forEach((bm, i) => {
      if (newOrder.indexOf(i) === -1) reordered.push(bm);
    });
    
    return saveAllBookmarks(email, reordered);
  } catch (error) {
    console.error('reorderBookmarks Error:', error);
    return { success: false, error: error.message };
  }
}


/* ==========================================================================
 * 6. Gmail 連携機能
 * ========================================================================== */

/**
 * Gmailの受信トレイから最新メッセージ (最大50件) を取得する
 *
 * @returns {Array<Object>} メールメッセージオブジェクトの配列
 *                          [{ id, subject, from, date, snippet, isUnread, permalink }]
 */
function getGmailMessages() {
  try {
    assertTermsAgreed(); // 利用規約チェック
    // 受信トレイのスレッドを先頭から最大50件取得
    const threads = GmailApp.getInboxThreads(0, 50);
    const messages = [];
    
    if (!threads || threads.length === 0) {
      return [];
    }
    
    for (const thread of threads) {
      const msgs = thread.getMessages();
      // スレッド内の最後のメッセージ（最新のやり取り）を取得
      const latestMsg = msgs[msgs.length - 1];
      
      // メール本文のプレーンテキストから先頭100文字をスニペットとして抽出
      let snippet = latestMsg.getPlainBody() || '';
      snippet = snippet.substring(0, 100);
      
      messages.push({
        id: latestMsg.getId(),
        subject: latestMsg.getSubject(),
        from: latestMsg.getFrom(),
        date: latestMsg.getDate().toISOString(),
        snippet: snippet,
        isUnread: thread.isUnread(),
        permalink: thread.getPermalink()
      });
    }
    
    return messages;
  } catch (error) {
    console.error('getGmailMessages Error:', error);
    return [];
  }
}


/* ==========================================================================
 * 7. Google カレンダー連携機能
 * ========================================================================== */

/**
 * 指定週 (7日間) のGoogleカレンダーイベントを取得する
 *
 * @param {number|string} weekOffset - 今週を基準とした週オフセット (0:今週, -1:前週, 1:次週)
 * @returns {Array<Object>} イベントオブジェクトの配列
 *                          [{ id, title, startTime, endTime, description, isAllDay, color }]
 */
function getCalendarEvents(weekOffset) {
  try {
    assertTermsAgreed(); // 利用規約チェック
    const calendar = CalendarApp.getDefaultCalendar();
    if (!calendar) return [];
    
    const offset = parseInt(weekOffset, 10) || 0;
    // 週の開始日（今日の日付の00:00:00）を週オフセットを加味して算出
    const startDate = new Date();
    startDate.setHours(0, 0, 0, 0);
    startDate.setDate(startDate.getDate() + (offset * 7));
    
    // 7日後を終了日とする
    const endDate = new Date(startDate.getTime() + (7 * 24 * 60 * 60 * 1000));
    
    // 期間内のカレンダーイベントを取得
    const events = calendar.getEvents(startDate, endDate);
    const result = [];
    
    if (!events || events.length === 0) {
      return [];
    }
    
    for (const event of events) {
      const isAllDay = event.isAllDayEvent();
      let startTimeISO = event.getStartTime().toISOString();
      let endTimeISO = event.getEndTime().toISOString();
      
      // 【終日イベントの終了日補正】
      // Google Calendar API の仕様上、終日イベントの endTime は翌日の 00:00:00 (exclusive) が返るため、
      // 1秒減算して実質的な最終日 23:59:59 に補正します
      if (isAllDay) {
        const adjustedEnd = new Date(event.getEndTime().getTime() - 1000);
        endTimeISO = adjustedEnd.toISOString();
      }

      result.push({
        id: event.getId(),
        title: event.getTitle(),
        startTime: startTimeISO,
        endTime: endTimeISO,
        description: event.getDescription() || '',
        isAllDay: isAllDay,
        color: event.getColor() || ''
      });
    }
    
    return result;
  } catch (error) {
    console.error('getCalendarEvents Error:', error);
    return [];
  }
}

/**
 * Googleカレンダーに予定を新規追加する (終日/時間指定対応)
 *
 * @param {string} title - 予定のタイトル
 * @param {string} startTime - 開始日時 (ISO文字列)
 * @param {string} endTime - 終了日時 (ISO文字列)
 * @param {string} description - 予定の説明・詳細
 * @param {string|number} color - Googleカレンダーのカラー番号 (1〜11)
 * @param {boolean} isAllDay - 終日イベントかどうかのフラグ
 * @returns {Object} { success: boolean, id?: string, error?: string }
 */
function addCalendarEvent(title, startTime, endTime, description, color, isAllDay) {
  try {
    assertTermsAgreed(); // 利用規約チェック
    const calendar = CalendarApp.getDefaultCalendar();
    const start = new Date(startTime);
    const end = new Date(endTime);
    let event;
    
    if (isAllDay) {
      // 終日イベントの場合
      const startDate = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const endDate = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      
      if (startDate.getTime() === endDate.getTime()) {
        // 単日の終日予定
        event = calendar.createAllDayEvent(title, startDate, { description: description || '' });
      } else {
        // 複数日にまたがる終日予定（Google仕様に合わせて終了日の翌日0時を指定）
        const nextDayOfEnd = new Date(endDate.getTime() + (24 * 60 * 60 * 1000));
        event = calendar.createAllDayEvent(title, startDate, nextDayOfEnd, { description: description || '' });
      }
    } else {
      // 時間指定の予定
      event = calendar.createEvent(title, start, end, {
        description: description || ''
      });
    }

    // イベントカラーの設定 (1〜11のカラーコード)
    if (color) {
      try {
        const colorVal = parseInt(color, 10);
        if (!isNaN(colorVal) && colorVal >= 1 && colorVal <= 11) {
          event.setColor(colorVal.toString());
        }
      } catch (ce) {
        console.warn('Set color error:', ce);
      }
    }
    
    const email = Session.getActiveUser().getEmail();
    writeLog(email, 'カレンダー', 'カレンダー予定を追加しました');
    return { success: true, id: event.getId() };
  } catch (error) {
    console.error('addCalendarEvent Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Googleカレンダーから予定を削除する
 *
 * @param {string} eventId - 削除対象のイベントID
 * @returns {Object} { success: boolean, message?: string }
 */
function deleteCalendarEvent(eventId) {
  try {
    assertTermsAgreed(); // 利用規約チェック
    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId);
    
    if (event) {
      const title = event.getTitle();
      event.deleteEvent();
      const email = Session.getActiveUser().getEmail();
      writeLog(email, 'カレンダー', 'カレンダー予定を削除しました');
      return { success: true };
    } else {
      return { success: false, message: 'イベントが見つかりませんでした' };
    }
  } catch (error) {
    console.error('deleteCalendarEvent Error:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Googleカレンダーの既存予定を更新（編集）する
 *
 * @param {string} eventId - 更新対象のイベントID
 * @param {string} title - タイトル
 * @param {string} startTime - 開始日時 (ISO文字列)
 * @param {string} endTime - 終了日時 (ISO文字列)
 * @param {string} description - 説明
 * @param {string|number} color - カラー番号 (1〜11)
 * @param {boolean} isAllDay - 終日フラグ
 * @returns {Object} { success: boolean, error?: string, message?: string }
 */
function updateCalendarEvent(eventId, title, startTime, endTime, description, color, isAllDay) {
  try {
    assertTermsAgreed(); // 利用規約チェック
    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId);
    if (!event) {
      return { success: false, message: 'イベントが見つかりませんでした' };
    }
    
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    // タイトルと詳細説明を更新
    event.setTitle(title);
    event.setDescription(description || '');
    
    // 日時・終日の更新
    if (isAllDay) {
      const startDate = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const endDate = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      if (startDate.getTime() === endDate.getTime()) {
        event.setAllDayDate(startDate);
      } else {
        const nextDayOfEnd = new Date(endDate.getTime() + (24 * 60 * 60 * 1000));
        event.setAllDayDates(startDate, nextDayOfEnd);
      }
    } else {
      event.setTime(start, end);
    }
    
    // イベントカラーの更新
    if (color) {
      try {
        const colorVal = parseInt(color, 10);
        if (!isNaN(colorVal) && colorVal >= 1 && colorVal <= 11) {
          event.setColor(colorVal.toString());
        }
      } catch (ce) {}
    }
    
    const email = Session.getActiveUser().getEmail();
    writeLog(email, 'カレンダー', 'カレンダー予定を更新しました');
    return { success: true };
  } catch (error) {
    console.error('updateCalendarEvent Error:', error);
    return { success: false, error: error.message };
  }
}


/* ==========================================================================
 * 8. Google Tasks 連携機能
 * ========================================================================== */

/**
 * Google Tasksの全タスクリストと、各リストに属するタスクを取得する
 *
 * @returns {Array<Object>|Object} タスクリスト一覧 [{ listId, listTitle, tasks: [...] }]、またはエラーオブジェクト
 */
function getTasks() {
  try {
    assertTermsAgreed(); // 利用規約チェック
    const taskLists = Tasks.Tasklists.list();
    if (!taskLists.items) {
      return [];
    }
    
    const result = [];
    
    // 各タスクリストを走査してタスク一覧を取得
    for (const list of taskLists.items) {
      const listId = list.id;
      const tasksResult = Tasks.Tasks.list(listId, {
        showCompleted: true, // 完了済みタスクも含める
        showHidden: true     // 非表示タスクも含める
      });
      
      const tasks = [];
      if (tasksResult.items) {
        for (const task of tasksResult.items) {
          tasks.push({
            id: task.id,
            title: task.title || '',
            notes: task.notes || '',
            due: task.due ? new Date(task.due).toISOString() : null,
            status: task.status,
            completed: task.status === 'completed'
          });
        }
      }
      
      result.push({
        listId: listId,
        listTitle: list.title,
        tasks: tasks
      });
    }
    
    return result;
  } catch (error) {
    console.error('getTasks Error:', error);
    return { error: error.message };
  }
}

/**
 * Google Tasksに新しいタスクを追加する
 *
 * @param {string} taskListId - 対象のタスクリストID
 * @param {string} title - タスクのタイトル
 * @param {string} notes - メモ・詳細
 * @param {string} dueDate - 期限日時 (ISO文字列)
 * @returns {Object} { success: boolean, id?: string, error?: string }
 */
function addTask(taskListId, title, notes, dueDate) {
  try {
    assertTermsAgreed(); // 利用規約チェック
    const newTask = {
      title: title,
      notes: notes || ''
    };
    
    if (dueDate) {
      newTask.due = new Date(dueDate).toISOString();
    }
    
    const task = Tasks.Tasks.insert(newTask, taskListId);
    const email = Session.getActiveUser().getEmail();
    writeLog(email, 'タスク', '新しいタスクを追加しました');
    return { success: true, id: task.id };
  } catch (error) {
    console.error('addTask Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Google Tasksのタスクを削除する
 *
 * @param {string} taskListId - 対象のタスクリストID
 * @param {string} taskId - 削除するタスクのID
 * @returns {Object} { success: boolean, error?: string }
 */
function deleteTask(taskListId, taskId) {
  try {
    assertTermsAgreed(); // 利用規約チェック
    Tasks.Tasks.remove(taskListId, taskId);
    const email = Session.getActiveUser().getEmail();
    writeLog(email, 'タスク', 'タスクを削除しました');
    return { success: true };
  } catch (error) {
    console.error('deleteTask Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Google Tasksのタスクを完了済みに更新する
 *
 * @param {string} taskListId - 対象のタスクリストID
 * @param {string} taskId - 完了にするタスクのID
 * @returns {Object} { success: boolean, error?: string }
 */
function completeTask(taskListId, taskId) {
  try {
    assertTermsAgreed(); // 利用規約チェック
    const task = Tasks.Tasks.get(taskListId, taskId);
    task.status = 'completed';
    Tasks.Tasks.patch(task, taskListId, taskId);
    const email = Session.getActiveUser().getEmail();
    writeLog(email, 'タスク', 'タスクを完了しました');
    return { success: true };
  } catch (error) {
    console.error('completeTask Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Google Tasksのタスクを更新する (リスト間移動もサポート)
 *
 * @param {string} taskListId - 現在のタスクリストID
 * @param {string} taskId - 更新対象のタスクID
 * @param {string} title - タイトル
 * @param {string} notes - メモ
 * @param {string|null} dueDate - 期限日時 (ISO文字列またはnull)
 * @param {string} [newTaskListId] - 移動先タスクリストID (省略時は同一リスト内で更新)
 * @returns {Object} { success: boolean, id?: string, error?: string }
 */
function updateTask(taskListId, taskId, title, notes, dueDate, newTaskListId) {
  try {
    assertTermsAgreed(); // 利用規約チェック
    const targetListId = newTaskListId || taskListId;

    if (targetListId !== taskListId) {
      // 【異なるタスクリストへの移動】
      // 元のタスク情報を取得して新リストにタスクを新規作成した後、旧タスクを削除する
      const originalTask = Tasks.Tasks.get(taskListId, taskId);
      const newTask = {
        title: title !== undefined ? title : originalTask.title,
        notes: notes !== undefined ? notes : (originalTask.notes || ''),
        status: originalTask.status || 'needsAction'
      };
      if (dueDate) {
        newTask.due = new Date(dueDate).toISOString();
      }
      const inserted = Tasks.Tasks.insert(newTask, targetListId);
      Tasks.Tasks.remove(taskListId, taskId);

      const email = Session.getActiveUser().getEmail();
      writeLog(email, 'タスク', 'タスクを更新・移動しました');
      return { success: true, id: inserted.id };
    } else {
      // 【同一リスト内での通常更新】
      const taskPatch = {
        title: title,
        notes: notes || '',
        due: dueDate ? new Date(dueDate).toISOString() : null
      };
      const updated = Tasks.Tasks.patch(taskPatch, taskListId, taskId);

      const email = Session.getActiveUser().getEmail();
      writeLog(email, 'タスク', 'タスクを更新しました');
      return { success: true, id: updated.id };
    }
  } catch (error) {
    console.error('updateTask Error:', error);
    return { success: false, error: error.message };
  }
}


/* ==========================================================================
 * 9. QRコード画像 Google Drive 保存機能
 * ========================================================================== */

/**
 * クライアント側で生成されたQRコード画像を Google Drive に保存する
 * - 「マイポータル - QRコード」フォルダを自動作成または取得して格納します。
 *
 * @param {string} email - ユーザーのメールアドレス
 * @param {string} base64Data - Base64エンコードされた画像データ文字列
 * @param {string} fileName - 保存するファイル名 (例: 'qr_2026-09-12T12-00-00.png')
 * @param {string} mimeType - MIMEタイプ ('image/png', 'image/jpeg', 'image/svg+xml')
 * @returns {Object} { success: boolean, url?: string, fileName?: string, error?: string }
 */
function saveQRCodeToDrive(email, base64Data, fileName, mimeType) {
  try {
    assertTermsAgreed(); // 利用規約チェック

    // 1. 「マイポータル - QRコード」保存先フォルダを検索または作成
    var folderName = 'マイポータル - QRコード';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    // 2. Base64文字列をバイナリBlobオブジェクトに復元
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType, fileName);

    // 3. フォルダ内にファイルを作成
    var file = folder.createFile(blob);
    file.setDescription('マイポータルから生成されたQRコード');

    var fileUrl = file.getUrl();

    // 4. 監査ログに保存履歴を記録
    writeLog(email, 'QRコード', 'QRコードをGoogle Driveに保存しました: ' + fileName);

    return {
      success: true,
      url: fileUrl,
      fileName: fileName
    };
  } catch (error) {
    console.error('saveQRCodeToDrive Error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}
