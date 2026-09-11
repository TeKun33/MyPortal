/**
 * マイポータル - Google Apps Script Server Side
 */

/**
 * アプリの初期表示処理 (Entry Point)
 * @param {Object} e イベントオブジェクト
 * @returns {HtmlOutput} HTML出力
 */
function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    return template.evaluate()
      .setTitle('マイポータル')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput('エラーが発生しました: ' + error.message);
  }
}

/**
 * ログ用シートを取得または作成する
 * @returns {Sheet} ログシート
 */
function getOrCreateLogSheet() {
  const SSID = PropertiesService.getScriptProperties().getProperty('SSID');
  const SS = SpreadsheetApp.openById(SSID);
  let sheet = SS.getSheetByName('__LOG__');
  if (!sheet) {
    sheet = SS.insertSheet('__LOG__');
    sheet.getRange(1, 1, 1, 4).setValues([['日時', 'ユーザー', '操作', '詳細']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 400);
  }
  return sheet;
}

/**
 * 利用規約同意状況シートを取得または作成する
 * @returns {Sheet} 利用規約シート
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
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 180);
    sheet.setColumnWidth(4, 180);
  }
  return sheet;
}

/**
 * スプレッドシートにログを記録する
 * @param {string} email ユーザーのメールアドレス
 * @param {string} action 操作名
 * @param {string} detail 詳細情報
 */
function writeLog(email, action, detail) {
  try {
    const sheet = getOrCreateLogSheet();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    sheet.appendRow([timestamp, email || '', action || '', detail || '']);
  } catch (e) {
    console.error('writeLog Error:', e);
  }
}

/**
 * 現在のユーザーのメールアドレスを取得する
 * @returns {string} メールアドレス
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
 * ユーザー専用のシートを取得または作成する
 * @param {string} email ユーザーのメールアドレス
 * @returns {Sheet} ユーザー用シートオブジェクト
 */
function getOrCreateUserSheet(email) {
  try {
    const SSID = PropertiesService.getScriptProperties().getProperty('SSID');
    const SS = SpreadsheetApp.openById(SSID);
    let sheet = SS.getSheetByName(email);
    
    if (!sheet) {
      sheet = SS.insertSheet(email);
      // Row 1: Settings (theme, darkMode, termsAgreed, termsAgreedAt)
      sheet.getRange(1, 1, 1, 9).setValues([['__SETTINGS__', 'theme', 'ocean', 'darkMode', 'false', 'termsAgreed', 'false', 'termsAgreedAt', '']]);
      // Row 2: Bookmarks Header
      sheet.getRange(2, 1, 1, 6).setValues([['__BOOKMARKS__', 'タイトル', 'URL', 'カテゴリ', 'アイコン', '追加日時']]);
    }
    
    return sheet;
  } catch (error) {
    console.error('getOrCreateUserSheet Error:', error);
    throw error;
  }
}

/**
 * ユーザーの設定を取得する
 * @param {string} email ユーザーのメールアドレス
 * @returns {Object} {theme, darkMode, termsAgreed, termsAgreedAt} 
 */
function getUserSettings(email) {
  try {
    const sheet = getOrCreateUserSheet(email);
    const settingsRange = sheet.getRange(1, 1, 1, 9).getValues()[0];
    
    return {
      theme: settingsRange[2] || 'ocean',
      darkMode: settingsRange[4] === 'true' || settingsRange[4] === true,
      termsAgreed: settingsRange[6] === 'true' || settingsRange[6] === true,
      termsAgreedAt: settingsRange[8] ? String(settingsRange[8]) : ''
    };
  } catch (error) {
    console.error('getUserSettings Error:', error);
    return { theme: 'ocean', darkMode: false, termsAgreed: false, termsAgreedAt: '' };
  }
}

/**
 * ユーザーの設定を保存する
 * @param {string} email ユーザーのメールアドレス
 * @param {Object} settings {theme, darkMode, termsAgreed}
 * @returns {Object} 処理結果
 */
function saveUserSettings(email, settings) {
  try {
    const sheet = getOrCreateUserSheet(email);
    
    if (settings.theme !== undefined) {
      sheet.getRange('C1').setValue(settings.theme);
    }
    if (settings.darkMode !== undefined) {
      sheet.getRange('E1').setValue(settings.darkMode.toString());
    }
    if (settings.termsAgreed !== undefined) {
      sheet.getRange('F1').setValue('termsAgreed');
      sheet.getRange('G1').setValue(settings.termsAgreed ? 'true' : 'false');
      sheet.getRange('H1').setValue('termsAgreedAt');
      if (settings.termsAgreed) {
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
        sheet.getRange('I1').setValue(timestamp);
      }
    }
    
    writeLog(email, '設定', '設定の変更をしました');
    return { success: true };
  } catch (error) {
    console.error('saveUserSettings Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 利用規約の同意状態をスプレッドシートに保存する
 * @param {string} email ユーザーのメールアドレス
 * @param {boolean} agreed 同意したかどうか
 * @returns {Object} 処理結果
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
    
    // 2. __TERMS__ シートの該当ユーザー行を更新または追加
    const termsSheet = getOrCreateTermsSheet();
    const data = termsSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        rowIndex = i + 1; // 1-based row index
        break;
      }
    }
    
    const statusText = isAgreed ? '同意済み' : '未同意';
    if (rowIndex > 0) {
      termsSheet.getRange(rowIndex, 2).setValue(statusText);
      if (isAgreed) {
        termsSheet.getRange(rowIndex, 3).setValue(timestamp);
      }
      termsSheet.getRange(rowIndex, 4).setValue(timestamp);
    } else {
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
 * 利用規約の同意状態を取得する
 * @param {string} email ユーザーのメールアドレス
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
 * ユーザーが利用規約に同意しているかを検証する
 * 未同意の場合はエラーをスローし、APIアクセスを遮断する
 * @param {string} [email] ユーザーのメールアドレス (省略時は実行中アカウント)
 */
function assertTermsAgreed(email) {
  if (!email) email = getUserEmail();
  const agreement = getTermsAgreement(email);
  if (!agreement.termsAgreed) {
    throw new Error('利用規約に同意していないため、この操作は許可されていません。');
  }
}


/**
 * ブックマーク一覧を取得する
 * @param {string} email ユーザーのメールアドレス
 * @returns {Array} ブックマークオブジェクトの配列
 */
function getBookmarks(email) {
  try {
    assertTermsAgreed(email);
    const sheet = getOrCreateUserSheet(email);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 2) {
      return [];
    }
    
    // Row 3 onwards
    const dataRange = sheet.getRange(3, 2, lastRow - 2, 5).getValues();
    const bookmarks = [];
    
    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      // 空行をスキップ
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
 * ブックマークを追加する
 * @param {string} email ユーザーのメールアドレス
 * @param {string} title タイトル
 * @param {string} url URL
 * @param {string} category カテゴリ
 * @returns {Object} 処理結果
 */
function addBookmark(email, title, url, category, icon) {
  try {
    assertTermsAgreed(email);
    const sheet = getOrCreateUserSheet(email);
    
    const iconName = icon || (category && !category.includes('/') ? category : 'language');
    const cat = (icon && category) ? category : '';
    const addedAt = new Date().toISOString();
    
    sheet.appendRow(['', title, url, cat, iconName, addedAt]);
    
    writeLog(email, 'ブックマーク', '新しいブックマークを追加しました');
    return { success: true };
  } catch (error) {
    console.error('addBookmark Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * ブックマークを削除する
 * @param {string} email ユーザーのメールアドレス
 * @param {number} index 削除するブックマークのインデックス (0-based)
 * @returns {Object} 処理結果
 */
/**
 * ブックマークを更新する（編集）
 * @param {string} email ユーザーのメールアドレス
 * @param {number} index 更新するブックマークのインデックス (0-based)
 * @param {string} title タイトル
 * @param {string} url URL
 * @param {string} [category] カテゴリ
 * @returns {Object} 処理結果
 */
function updateBookmark(email, index, title, url, category, icon) {
  try {
    assertTermsAgreed(email);
    const sheet = getOrCreateUserSheet(email);
    const rowToUpdate = index + 3;
    
    if (rowToUpdate < 3 || rowToUpdate > sheet.getLastRow()) {
      return { success: false, error: '指定されたブックマークが見つかりません' };
    }

    const iconName = icon || (category && !category.includes('/') ? category : 'language');
    const cat = (icon && category) ? category : '';

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

function deleteBookmark(email, index) {
  try {
    assertTermsAgreed(email);
    const sheet = getOrCreateUserSheet(email);
    const rowToDelete = index + 3;
    
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
 * 全ブックマークを並び順通りに上書き保存する
 * @param {string} email ユーザーのメールアドレス
 * @param {Array<Object>} bookmarks 並び替え後のブックマークオブジェクト配列
 * @returns {Object} 処理結果
 */
function saveAllBookmarks(email, bookmarks) {
  try {
    assertTermsAgreed(email);
    const sheet = getOrCreateUserSheet(email);
    const lastRow = sheet.getLastRow();
    
    // 3行目以降の既存ブックマークデータをクリア
    if (lastRow >= 3) {
      sheet.getRange(3, 1, lastRow - 2, 6).clearContent();
    }
    
    if (!bookmarks || bookmarks.length === 0) {
      return { success: true };
    }
    
    const rows = bookmarks.map(bm => {
      const icon = bm.icon || 'language';
      const addedAt = bm.addedAt || new Date().toISOString();
      return ['', bm.title || '', bm.url || '', bm.category || '', icon, addedAt];
    });
    
    sheet.getRange(3, 1, rows.length, 6).setValues(rows);
    writeLog(email, 'ブックマーク', 'ブックマークの並び順を更新しました');
    return { success: true };
  } catch (error) {
    console.error('saveAllBookmarks Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 互換用: インデックス配列で並び替え
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

/**
 * Gmailの最新メッセージを取得する
 * @returns {Array} メッセージオブジェクトの配列
 */
function getGmailMessages() {
  try {
    assertTermsAgreed();
    const threads = GmailApp.getInboxThreads(0, 50);
    const messages = [];
    
    if (!threads || threads.length === 0) {
      return [];
    }
    
    for (const thread of threads) {
      const msgs = thread.getMessages();
      const latestMsg = msgs[msgs.length - 1];
      
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

/**
 * 指定された週（デフォルトは今週）の7日間のGoogleカレンダーイベントを取得する
 * @param {number|string} weekOffset 今週からの週オフセット (0: 今週, -1: 前週, 1: 次週)
 * @returns {Array} イベントオブジェクトの配列
 */
function getCalendarEvents(weekOffset) {
  try {
    assertTermsAgreed();
    const calendar = CalendarApp.getDefaultCalendar();
    if (!calendar) return [];
    
    const offset = parseInt(weekOffset, 10) || 0;
    const startDate = new Date();
    startDate.setHours(0, 0, 0, 0);
    startDate.setDate(startDate.getDate() + (offset * 7));
    
    const endDate = new Date(startDate.getTime() + (7 * 24 * 60 * 60 * 1000));
    
    const events = calendar.getEvents(startDate, endDate);
    const result = [];
    
    if (!events || events.length === 0) {
      return [];
    }
    
    for (const event of events) {
      const isAllDay = event.isAllDayEvent();
      let startTimeISO = event.getStartTime().toISOString();
      let endTimeISO = event.getEndTime().toISOString();
      
      // 終日イベントの場合、Google Calendar APIは endTime に翌日00:00:00 (exclusive) を返すため1秒減算して実質最終日23:59:59にする
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
 * Googleカレンダーにイベントを追加する
 * @param {string} title タイトル
 * @param {string} startTime 開始日時/開始日 (ISO文字列)
 * @param {string} endTime 終了日時/終了日 (ISO文字列)
 * @param {string} description 説明
 * @param {string|number} color イベントのカラーID (1-11)
 * @param {boolean} isAllDay 終日フラグ
 * @returns {Object} 処理結果
 */
function addCalendarEvent(title, startTime, endTime, description, color, isAllDay) {
  try {
    assertTermsAgreed();
    const calendar = CalendarApp.getDefaultCalendar();
    const start = new Date(startTime);
    const end = new Date(endTime);
    let event;
    
    if (isAllDay) {
      const startDate = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const endDate = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      
      if (startDate.getTime() === endDate.getTime()) {
        event = calendar.createAllDayEvent(title, startDate, { description: description || '' });
      } else {
        // 複数日終日イベント（Google Calendar API仕様に合わせて終了日の翌日0時を設定）
        const nextDayOfEnd = new Date(endDate.getTime() + (24 * 60 * 60 * 1000));
        event = calendar.createAllDayEvent(title, startDate, nextDayOfEnd, { description: description || '' });
      }
    } else {
      event = calendar.createEvent(title, start, end, {
        description: description || ''
      });
    }

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
 * Googleカレンダーのイベントを削除する
 * @param {string} eventId イベントID
 * @returns {Object} 処理結果
 */
function deleteCalendarEvent(eventId) {
  try {
    assertTermsAgreed();
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
 * Googleカレンダーのイベントを更新（編集）する
 * @param {string} eventId イベントID
 * @param {string} title タイトル
 * @param {string} startTime 開始日時/開始日(ISO文字列)
 * @param {string} endTime 終了日時/終了日(ISO文字列)
 * @param {string} description 説明
 * @param {string|number} color イベントのカラーID (1-11)
 * @param {boolean} isAllDay 終日フラグ
 * @returns {Object} 処理結果
 */
function updateCalendarEvent(eventId, title, startTime, endTime, description, color, isAllDay) {
  try {
    assertTermsAgreed();
    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId);
    if (!event) {
      return { success: false, message: 'イベントが見つかりませんでした' };
    }
    
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    event.setTitle(title);
    event.setDescription(description || '');
    
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

/**
 * Google Tasksの全タスクリストとタスクを取得する
 * @returns {Array|Object} タスクリストの配列、またはエラーオブジェクト
 */
function getTasks() {
  try {
    assertTermsAgreed();
    const taskLists = Tasks.Tasklists.list();
    if (!taskLists.items) {
      return [];
    }
    
    const result = [];
    
    for (const list of taskLists.items) {
      const listId = list.id;
      const tasksResult = Tasks.Tasks.list(listId, {
        showCompleted: true,
        showHidden: true
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
 * Google Tasksにタスクを追加する
 * @param {string} taskListId タスクリストID
 * @param {string} title タイトル
 * @param {string} notes メモ
 * @param {string} dueDate 期限(ISO文字列)
 * @returns {Object} 処理結果
 */
function addTask(taskListId, title, notes, dueDate) {
  try {
    assertTermsAgreed();
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
 * @param {string} taskListId タスクリストID
 * @param {string} taskId タスクID
 * @returns {Object} 処理結果
 */
function deleteTask(taskListId, taskId) {
  try {
    assertTermsAgreed();
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
 * Google Tasksのタスクを完了済みにする
 * @param {string} taskListId タスクリストID
 * @param {string} taskId タスクID
 * @returns {Object} 処理結果
 */
function completeTask(taskListId, taskId) {
  try {
    assertTermsAgreed();
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
 * Google Tasksのタスクを更新する
 * @param {string} taskListId タスクリストID
 * @param {string} taskId タスクID
 * @param {string} title タイトル
 * @param {string} notes メモ
 * @param {string} dueDate 期限(ISO文字列またはnull)
 * @param {string} [newTaskListId] 移動先タスクリストID (省略時は同一リスト)
 * @returns {Object} 処理結果
 */
function updateTask(taskListId, taskId, title, notes, dueDate, newTaskListId) {
  try {
    assertTermsAgreed();
    const targetListId = newTaskListId || taskListId;

    if (targetListId !== taskListId) {
      // 異なるタスクリストへの移動: 元のタスクを取得し新リストに作成後、旧タスクを削除
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
      // 同一リスト内での更新
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
