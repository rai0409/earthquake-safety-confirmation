/**
 * SheetRepository.gs
 * スプレッドシートへの読み書き処理を集約
 * 行ごとにgetRangeを呼ばず、一括読み書きを優先する
 */

/**
 * スプレッドシートを取得
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSpreadsheet() {
  const spreadsheetId = PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      'Script PropertiesのSPREADSHEET_IDが未設定です'
    );
  }

  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      `スプレッドシートを開けません: ${safeErrorMessage(error)}`
    );
  }
}

/**
 * シートをname指定で取得
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getSheet(name) {
  const spreadsheet = getSpreadsheet();
  return spreadsheet.getSheetByName(name) || null;
}

/**
 * シートを取得または作成
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(name) {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }

  return sheet;
}

/**
 * Googleフォームの回答先シートを取得する
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getFormResponseSheet() {
  const sheetName = String(
    getSetting(
      'form_response_sheet_name',
      'フォームの回答 1'
    )
  ).trim();

  if (!sheetName) {
    return null;
  }

  return getSpreadsheet().getSheetByName(sheetName) || null;
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

  if (!sheet) {
    return;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === String(key).trim()) {
      sheet.getRange(i + 1, 2).setValue(value);
      clearSettingsCache();
      return;
    }
  }

  sheet.appendRow([key, value]);
  clearSettingsCache();
}

// ---------------------------------------------------------------------------
// earthquake_events
// ---------------------------------------------------------------------------

const EQ_HEADERS = [
  'event_id',
  'occurred_at',
  'announced_at',
  'hypocenter',
  'magnitude',
  'max_intensity',
  'source_url',
  'detected_at',
  'status',
  'target_count',
  'sent_count',
  'failed_count',
  'completed_at',
  'error',
];

/**
 * 地震イベントを末尾へ追加
 * @param {Object} event
 */
function appendEvent(event) {
  const sheet = getSheet(SHEET.EARTHQUAKE_EVENTS);

  if (!sheet) {
    throw new Error(
      'earthquake_eventsシートが見つかりません'
    );
  }

  sheet.appendRow([
    event.eventId,
    event.occurredAt
      ? event.occurredAt.toISOString()
      : '',
    event.announcedAt
      ? event.announcedAt.toISOString()
      : '',
    event.hypocenter || '',
    event.magnitude !== null &&
    event.magnitude !== undefined
      ? event.magnitude
      : '',
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
 * @param {Object} values
 */
function updateEvent(eventId, values) {
  const sheet = getSheet(SHEET.EARTHQUAKE_EVENTS);

  if (!sheet) {
    return;
  }

  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    return;
  }

  const headers = data[0].map(
    header => String(header || '').trim()
  );

  const eventIdColumnIndex = headers.indexOf('event_id');

  if (eventIdColumnIndex < 0) {
    throw new Error(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      'earthquake_eventsシートにevent_id列がありません'
    );
  }

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][eventIdColumnIndex] || '').trim() ===
      String(eventId || '').trim()
    ) {
      for (const [key, value] of Object.entries(values)) {
        const columnIndex = headers.indexOf(key);

        if (columnIndex >= 0) {
          sheet
            .getRange(i + 1, columnIndex + 1)
            .setValue(value);
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

  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return null;
  }

  const headers = data[0].map(
    header => String(header || '').trim()
  );
  const eventIdColumnIndex = headers.indexOf('event_id');

  if (eventIdColumnIndex < 0) {
    return null;
  }

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][eventIdColumnIndex] || '').trim() ===
      String(eventId || '').trim()
    ) {
      const event = {};

      headers.forEach((header, index) => {
        event[header] = data[i][index];
      });

      return event;
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
  const event = findEvent(eventId);

  if (!event) {
    return false;
  }

  const completedStatuses = [
    EVENT_STATUS.COMPLETED,
    EVENT_STATUS.PARTIAL_FAILED,
    EVENT_STATUS.TEST_COMPLETED,
  ];

  return completedStatuses.includes(
    String(event.status || '').trim()
  );
}

// ---------------------------------------------------------------------------
// notification_status
// ---------------------------------------------------------------------------

const NS_HEADERS = [
  'notification_key',
  'event_id',
  'employee_id',
  'name',
  'email',
  'channel',
  'status',
  'attempts',
  'created_at',
  'sending_at',
  'sent_at',
  'responded_at',
  'error',
];

/**
 * notification_keyで通知レコードを検索
 * @param {string} notificationKey
 * @returns {Object|null}
 */
function findNotification(notificationKey) {
  const sheet = getSheet(SHEET.NOTIFICATION_STATUS);

  if (!sheet) {
    return null;
  }

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return null;
  }

  const headers = data[0].map(
    header => String(header || '').trim()
  );
  const keyColumnIndex = headers.indexOf('notification_key');

  if (keyColumnIndex < 0) {
    return null;
  }

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][keyColumnIndex] || '').trim() ===
      String(notificationKey || '').trim()
    ) {
      const notification = {};

      headers.forEach((header, index) => {
        notification[header] = data[i][index];
      });

      notification._rowIndex = i + 1;
      return notification;
    }
  }

  return null;
}

/**
 * notification_keyで通知レコードを更新
 * @param {string} notificationKey
 * @param {Object} values
 */
function updateNotification(notificationKey, values) {
  const sheet = getSheet(SHEET.NOTIFICATION_STATUS);

  if (!sheet) {
    return;
  }

  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    return;
  }

  const headers = data[0].map(
    header => String(header || '').trim()
  );
  const keyColumnIndex = headers.indexOf('notification_key');

  if (keyColumnIndex < 0) {
    throw new Error(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      'notification_statusシートにnotification_key列がありません'
    );
  }

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][keyColumnIndex] || '').trim() ===
      String(notificationKey || '').trim()
    ) {
      for (const [key, value] of Object.entries(values)) {
        const columnIndex = headers.indexOf(key);

        if (columnIndex >= 0) {
          sheet
            .getRange(i + 1, columnIndex + 1)
            .setValue(value);
        }
      }

      return;
    }
  }
}

/**
 * 通知レコードを末尾へ追加
 * シートの列順ではなくヘッダー名に基づいて値を配置する
 *
 * @param {Object} record
 */
function appendNotification(record) {
  const sheet = getSheet(SHEET.NOTIFICATION_STATUS);

  if (!sheet) {
    throw new Error(
      'notification_statusシートが見つかりません'
    );
  }

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    throw new Error(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      'notification_statusシートにヘッダーがありません'
    );
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(header => String(header || '').trim());

  const missingHeaders = NS_HEADERS.filter(
    header => !headers.includes(header)
  );

  if (missingHeaders.length > 0) {
    throw new Error(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      'notification_statusシートの必須列が不足しています: ' +
      missingHeaders.join(', ')
    );
  }

  const valuesByHeader = {
    notification_key: record.notificationKey,
    event_id: record.eventId,
    employee_id: record.employeeId,
    name: record.name || '',
    email: record.email,
    channel: record.channel,
    status:
      record.status || NOTIFICATION_STATUS.PENDING,
    attempts: record.attempts || 0,
    created_at: record.createdAt || nowIso(),
    sending_at: '',
    sent_at: '',
    responded_at: '',
    error: '',
  };

  const row = headers.map(header => {
    if (
      Object.prototype.hasOwnProperty.call(
        valuesByHeader,
        header
      )
    ) {
      return valuesByHeader[header];
    }

    return '';
  });

  sheet.appendRow(row);
}

/**
 * 特定イベントの全通知レコードを取得
 * @param {string} eventId
 * @returns {Array<Object>}
 */
function getNotificationsByEventId(eventId) {
  const sheet = getSheet(SHEET.NOTIFICATION_STATUS);

  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const headers = data[0].map(
    header => String(header || '').trim()
  );
  const eventIdColumnIndex = headers.indexOf('event_id');

  if (eventIdColumnIndex < 0) {
    throw new Error(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      'notification_statusシートにevent_id列がありません'
    );
  }

  const results = [];

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][eventIdColumnIndex] || '').trim() ===
      String(eventId || '').trim()
    ) {
      const notification = {};

      headers.forEach((header, index) => {
        notification[header] = data[i][index];
      });

      notification._rowIndex = i + 1;
      results.push(notification);
    }
  }

  return results;
}

/**
 * 特定イベントのpending通知レコードを取得
 * @param {string} eventId
 * @returns {Array<Object>}
 */
function getPendingNotifications(eventId) {
  return getNotificationsByEventId(eventId).filter(
    notification =>
      String(notification.status || '').trim() ===
      NOTIFICATION_STATUS.PENDING
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

  if (!sheet) {
    return;
  }

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
// Googleフォーム回答（読み取り専用）
// ---------------------------------------------------------------------------

/**
 * 設定されたGoogleフォーム回答シートの全回答を取得
 * @returns {Array<Object>}
 */
function getAllFormResponses() {
  const sheet = getFormResponseSheet();

  if (!sheet) {
    return [];
  }

  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return [];
  }

  const headers = data[0].map(
    header => String(header || '').trim()
  );
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const response = {};

    headers.forEach((header, index) => {
      response[header] = data[i][index];
    });

    results.push(response);
  }

  return results;
}

/**
 * 特定eventIdの回答を全件取得
 * @param {string} eventId
 * @returns {Array<Object>}
 */
function getFormResponsesByEventId(eventId) {
  const normalizedEventId = String(
    eventId || ''
  ).trim();

  if (!normalizedEventId) {
    return [];
  }

  return getAllFormResponses().filter(response =>
    String(
      response['event_id'] ||
      response['eventId'] ||
      ''
    ).trim() === normalizedEventId
  );
}
