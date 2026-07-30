/**
 * SheetRepository.gs
 * スプレッドシートへの読み書き処理を集約
 * 行ごとにgetRangeを呼ばず、一括読み書きを優先する
 */

// スプレッドシートID（空ならアクティブなスプレッドシートを使用）
const SPREADSHEET_ID = ''; // 例: 'YOUR_SPREADSHEET_ID'

/**
 * スプレッドシートを取得
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * シートをname指定で取得
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSheet(name) {
  const ss = getSpreadsheet();
  return ss.getSheetByName(name) || null;
}

/**
 * シートを取得または作成
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

/**
 * 設定値を上書き保存
 * @param {string} key
 * @param {*} value
 */
function setSetting(key, value) {
  const sheet = getSheet(SHEET.SETTINGS);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      clearSettingsCache();
      return;
    }
  }
  // 存在しない場合は末尾に追加
  sheet.appendRow([key, value]);
  clearSettingsCache();
}

// ---------------------------------------------------------------------------
// earthquake_events
// ---------------------------------------------------------------------------

/** earthquake_events ヘッダー */
const EQ_HEADERS = [
  'event_id', 'occurred_at', 'announced_at', 'hypocenter', 'magnitude',
  'max_intensity', 'source_url', 'detected_at', 'status', 'target_count',
  'sent_count', 'failed_count', 'completed_at', 'error',
];

/**
 * 地震イベントを末尾へ追加
 * @param {Object} event
 */
function appendEvent(event) {
  const sheet = getSheet(SHEET.EARTHQUAKE_EVENTS);
  if (!sheet) throw new Error('earthquake_eventsシートが見つかりません');
  sheet.appendRow([
    event.eventId,
    event.occurredAt ? event.occurredAt.toISOString() : '',
    event.announcedAt ? event.announcedAt.toISOString() : '',
    event.hypocenter || '',
    event.magnitude !== null && event.magnitude !== undefined ? event.magnitude : '',
    event.maxIntensity || '',
    event.sourceUrl || '',
    event.detectedAt || nowIso(),
    event.status || EVENT_STATUS.DETECTED,
    event.targetCount || 0,
    event.sentCount || 0,
    event.failedCount || 0,
    event.completedAt || '',
    event.error || '',
  ]);
}

/**
 * event_idでイベント行を検索して更新
 * @param {string} eventId
 * @param {Object} values - 更新するキーと値のペア
 */
function updateEvent(eventId, values) {
  const sheet = getSheet(SHEET.EARTHQUAKE_EVENTS);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(eventId).trim()) {
      for (const [key, val] of Object.entries(values)) {
        const col = headers.indexOf(key);
        if (col >= 0) {
          sheet.getRange(i + 1, col + 1).setValue(val);
        }
      }
      return;
    }
  }
}

/**
 * event_idでイベント行を取得
 * @param {string} eventId
 * @returns {Object|null}
 */
function findEvent(eventId) {
  const sheet = getSheet(SHEET.EARTHQUAKE_EVENTS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(eventId).trim()) {
      const obj = {};
      headers.forEach((h, idx) => { obj[String(h).trim()] = data[i][idx]; });
      return obj;
    }
  }
  return null;
}

/**
 * completed / partial_failed / test_completed なeventが存在するか確認
 * @param {string} eventId
 * @returns {boolean}
 */
function isEventAlreadyCompleted(eventId) {
  const ev = findEvent(eventId);
  if (!ev) return false;
  const completedStatuses = [
    EVENT_STATUS.COMPLETED,
    EVENT_STATUS.PARTIAL_FAILED,
    EVENT_STATUS.TEST_COMPLETED,
  ];
  return completedStatuses.includes(String(ev.status).trim());
}

// ---------------------------------------------------------------------------
// notification_status
// ---------------------------------------------------------------------------

/** notification_status ヘッダー */
const NS_HEADERS = [
  'notification_key', 'event_id', 'employee_id', 'email', 'channel',
  'status', 'attempts', 'created_at', 'sending_at', 'sent_at',
  'responded_at', 'error',
];

/**
 * notification_key で通知レコードを検索
 * @param {string} notificationKey
 * @returns {Object|null}
 */
function findNotification(notificationKey) {
  const sheet = getSheet(SHEET.NOTIFICATION_STATUS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === notificationKey) {
      const obj = {};
      headers.forEach((h, idx) => { obj[String(h).trim()] = data[i][idx]; });
      obj._rowIndex = i + 1;
      return obj;
    }
  }
  return null;
}

/**
 * notification_key で通知レコードを更新
 * @param {string} notificationKey
 * @param {Object} values
 */
function updateNotification(notificationKey, values) {
  const sheet = getSheet(SHEET.NOTIFICATION_STATUS);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === notificationKey) {
      for (const [key, val] of Object.entries(values)) {
        const col = headers.indexOf(key);
        if (col >= 0) {
          sheet.getRange(i + 1, col + 1).setValue(val);
        }
      }
      return;
    }
  }
}

/**
 * 通知レコードを末尾へ追加
 * @param {Object} record
 */
function appendNotification(record) {
  const sheet = getSheet(SHEET.NOTIFICATION_STATUS);
  if (!sheet) throw new Error('notification_statusシートが見つかりません');
  sheet.appendRow([
    record.notificationKey,
    record.eventId,
    record.employeeId,
    record.email,
    record.channel,
    record.status || NOTIFICATION_STATUS.PENDING,
    record.attempts || 0,
    record.createdAt || nowIso(),
    '',  // sending_at
    '',  // sent_at
    '',  // responded_at
    '',  // error
  ]);
}

/**
 * 特定イベントの全通知レコードを取得
 * @param {string} eventId
 * @returns {Array<Object>}
 */
function getNotificationsByEventId(eventId) {
  const sheet = getSheet(SHEET.NOTIFICATION_STATUS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(eventId).trim()) {
      const obj = {};
      headers.forEach((h, idx) => { obj[String(h).trim()] = data[i][idx]; });
      obj._rowIndex = i + 1;
      results.push(obj);
    }
  }
  return results;
}

/**
 * 特定イベントの pending 通知レコードを取得
 * @param {string} eventId
 * @returns {Array<Object>}
 */
function getPendingNotifications(eventId) {
  return getNotificationsByEventId(eventId).filter(
    n => String(n.status).trim() === NOTIFICATION_STATUS.PENDING
  );
}

// ---------------------------------------------------------------------------
// send_errors
// ---------------------------------------------------------------------------

/**
 * 送信エラーを記録
 * @param {Object} error
 */
function appendSendError(error) {
  const sheet = getSheet(SHEET.SEND_ERRORS);
  if (!sheet) return;
  sheet.appendRow([
    error.timestamp || nowIso(),
    error.eventId || '',
    error.employeeId || '',
    error.email || '',
    error.channel || '',
    error.attempt || 0,
    error.responseCode || '',
    error.error || '',
  ]);
}

// ---------------------------------------------------------------------------
// form_responses（読み取り専用）
// ---------------------------------------------------------------------------

/**
 * form_responsesシートの全回答を取得
 * @returns {Array<Object>}
 */
function getAllFormResponses() {
  const sheet = getSheet(SHEET.FORM_RESPONSES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, idx) => { obj[String(h).trim()] = data[i][idx]; });
    results.push(obj);
  }
  return results;
}

/**
 * 特定eventIdの回答を全件取得
 * @param {string} eventId
 * @returns {Array<Object>}
 */
function getFormResponsesByEventId(eventId) {
  return getAllFormResponses().filter(r =>
    String(r['event_id'] || r['eventId'] || '').trim() === String(eventId).trim()
  );
}
