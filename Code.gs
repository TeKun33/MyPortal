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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('__LOG__');
  if (!sheet) {
    sheet = ss.insertSheet('__LOG__');
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(email);
    
    if (!sheet) {
      sheet = ss.insertSheet(email);
      // Row 1: Settings
      sheet.getRange(1, 1, 1, 5).setValues([['__SETTINGS__', 'theme', 'ocean', 'darkMode', 'false']]);
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
 * @returns {Object} {theme, darkMode} 
 */
function getUserSettings(email) {
  try {
    const sheet = getOrCreateUserSheet(email);
    const settingsRange = sheet.getRange(1, 1, 1, 5).getValues()[0];
    
    return {
      theme: settingsRange[2],
      darkMode: settingsRange[4] === 'true' || settingsRange[4] === true
    };
  } catch (error) {
    console.error('getUserSettings Error:', error);
    return { theme: 'ocean', darkMode: false };
  }
}

/**
 * ユーザーの設定を保存する
 * @param {string} email ユーザーのメールアドレス
 * @param {Object} settings {theme, darkMode}
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
    
    writeLog(email, '設定変更', 'テーマ: ' + (settings.theme || '') + ', ダークモード: ' + (settings.darkMode || ''));
    return { success: true };
  } catch (error) {
    console.error('saveUserSettings Error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * ブックマーク一覧を取得する
 * @param {string} email ユーザーのメールアドレス
 * @returns {Array} ブックマークオブジェクトの配列
 */
function getBookmarks(email) {
  try {
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
function addBookmark(email, title, url, category) {
  try {
    const sheet = getOrCreateUserSheet(email);
    
    let domain = '';
    try {
      const urlObj = url.match(/^(?:https?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n?]+)/im);
      if (urlObj && urlObj[1]) {
        domain = urlObj[1];
      }
    } catch (e) {
      // 無視
    }
    
    const icon = `https://www.google.com/s2/favicons?domain=${domain}&sz=32`;
    const addedAt = new Date().toISOString();
    
    sheet.appendRow(['', title, url, category || '', icon, addedAt]);
    
    writeLog(email, 'ブックマーク追加', title + ' (' + url + ')');
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
function deleteBookmark(email, index) {
  try {
    const sheet = getOrCreateUserSheet(email);
    const rowToDelete = index + 3;
    
    if (rowToDelete > 2 && rowToDelete <= sheet.getLastRow()) {
      sheet.deleteRow(rowToDelete);
    }
    
    writeLog(email, 'ブックマーク削除', 'インデックス: ' + index);
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
      let domain = '';
      try {
        const urlObj = (bm.url || '').match(/^(?:https?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n?]+)/im);
        if (urlObj && urlObj[1]) {
          domain = urlObj[1];
        }
      } catch (e) {}
      const icon = bm.icon || `https://www.google.com/s2/favicons?domain=${domain}&sz=32`;
      const addedAt = bm.addedAt || new Date().toISOString();
      return ['', bm.title || '', bm.url || '', bm.category || '', icon, addedAt];
    });
    
    sheet.getRange(3, 1, rows.length, 6).setValues(rows);
    writeLog(email, 'ブックマーク並び替え', rows.length + ' 件の並び順を更新');
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
    writeLog(email, 'カレンダー予定追加', title + ' (' + startTime + ' 〜 ' + endTime + ')');
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
    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId);
    
    if (event) {
      const title = event.getTitle();
      event.deleteEvent();
      const email = Session.getActiveUser().getEmail();
      writeLog(email, 'カレンダー予定削除', title + ' (ID: ' + eventId + ')');
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
    writeLog(email, 'カレンダー予定更新', title + ' (ID: ' + eventId + ')');
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
    const newTask = {
      title: title,
      notes: notes || ''
    };
    
    if (dueDate) {
      newTask.due = new Date(dueDate).toISOString();
    }
    
    const task = Tasks.Tasks.insert(newTask, taskListId);
    const email = Session.getActiveUser().getEmail();
    writeLog(email, 'タスク追加', title + ' (リストID: ' + taskListId + ')');
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
    Tasks.Tasks.remove(taskListId, taskId);
    const email = Session.getActiveUser().getEmail();
    writeLog(email, 'タスク削除', 'タスクID: ' + taskId);
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
    const task = Tasks.Tasks.get(taskListId, taskId);
    task.status = 'completed';
    Tasks.Tasks.patch(task, taskListId, taskId);
    const email = Session.getActiveUser().getEmail();
    writeLog(email, 'タスク完了', 'タスク: ' + (task.title || taskId));
    return { success: true };
  } catch (error) {
    console.error('completeTask Error:', error);
    return { success: false, error: error.message };
  }
}
