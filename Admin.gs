/**
 * Admin.gs
 * 管理者用操作関数
 */

// settingsシートの初期データ
const DEFAULT_SETTINGS = [
  ['key', 'value', 'description'],
  ['enabled', 'FALSE', 'システムを有効にするか (TRUE/FALSE)'],
  ['test_mode', 'TRUE', 'テストモード (TRUE=テスト受信者のみ送信)'],
  ['send_mode', 'gmail', '送信方式: gmail または outlook'],
  ['threshold', '5-', '通知する最低震度 (1/2/3/4/5-/5+/6-/6+/7)'],
  ['max_event_age_minutes', '30', '地震発生からの最大経過時間（分）'],
  ['check_interval_minutes', '10', '地震確認トリガーの間隔（分）'],
  ['stale_sending_minutes', '30', 'sending状態をstaleと判断する時間（分）'],
  ['form_response_sheet_name', 'フォームの回答 1', 'Googleフォームの回答先シート名'],
  ['admin_email', '', '管理者メールアドレス'],
  ['test_recipient_email', '', 'テスト送信先メールアドレス'],
  ['form_base_url', '', 'GoogleフォームURL (viewform形式)'],
  ['form_event_entry_id', '', 'フォームevent_idフィールドのentry ID (例: entry.123456789)'],
  ['form_employee_id_entry_id', '', 'フォームemployee_idフィールドのentry ID'],
  ['form_name_entry_id', '', 'フォームnameフィールドのentry ID'],
  ['form_id', '', 'GoogleフォームID (onFormSubmitトリガー用)'],
  ['sender_display_name', '災害時安否確認', 'メール送信者表示名'],
  ['outlook_sender_email', '', 'Outlook送信者メールアドレス (Script Propertiesが優先)'],
  ['retry_max_attempts', '3', 'Outlook送信最大再試行回数'],
  ['retry_wait_seconds', '3', 'Outlook再試行待機秒数'],
  ['immediate_admin_alert', 'TRUE', '緊急回答時の即時管理者通知 (TRUE/FALSE)'],
  ['summary_enabled', 'TRUE', 'summary自動更新を有効にするか'],
  ['form_col_event_id', 'event_id', 'form_responsesシートのevent_id列名'],
  ['form_col_employee_id', 'employee_id', 'form_responsesシートのemployee_id列名'],
  ['form_col_name', 'name', 'form_responsesシートのname列名'],
  ['form_col_safety_status', 'safety_status', 'form_responsesシートのsafety_status列名'],
  ['form_col_attendance_status', 'attendance_status', 'form_responsesシートのattendance_status列名'],
  ['form_col_comment', 'comment', 'form_responsesシートのcomment列名'],
  ['earthquake_info_url', '', '気象庁地震情報取得URL (アダプタ実装時に設定)'],
];

/**
 * 必要シートとヘッダーを作成する
 * 既存データを破壊しない。既存シートは不足列のみ確認する
 */
function setupSpreadsheet() {
  Logger.log('setupSpreadsheet: 開始');

  _setupSettingsSheet();
  _setupSheet(SHEET.EMPLOYEES, ['employee_id', 'name', 'email', 'active', 'group']);
  _setupSheet(SHEET.EARTHQUAKE_EVENTS, EQ_HEADERS);
  _setupSheet(SHEET.NOTIFICATION_STATUS, NS_HEADERS);
  _setupSheet(SHEET.SEND_ERRORS, ['timestamp', 'event_id', 'employee_id', 'email', 'channel', 'attempt', 'response_code', 'error']);
  _setupSheet(SHEET.SUMMARY, ['event_id', 'occurred_at', 'hypocenter', 'max_intensity', 'target_employee_count', 'sent_count', 'failed_count', 'response_count', 'unanswered_count', 'response_rate', 'no_damage_count', 'personal_injury_count', 'family_damage_count', 'property_damage_count', 'multiple_damage_count', 'difficult_to_respond_count', 'available_to_work_count', 'unavailable_disaster_count', 'unavailable_other_count', 'last_updated_at']);
  _setupSheet(
    SHEET.CURRENT_EVENT_DETAILS,
    ['section', 'employee_id', 'email', 'status', 'comment']
  );
  Logger.log('setupSpreadsheet: 完了');
  Logger.log('次のステップ: settingsシートへ必要な値を入力し、installTriggers()を実行してください');
}

/**
 * 管理対象スプレッドシートIDをScript Propertiesへ保存する
 *
 * @param {string} spreadsheetId
 */
function setSpreadsheetId(spreadsheetId) {
  const normalizedId = String(spreadsheetId || '').trim();

  if (!normalizedId) {
    throw new Error('spreadsheetIdを指定してください');
  }

  SpreadsheetApp.openById(normalizedId);

  PropertiesService
    .getScriptProperties()
    .setProperty('SPREADSHEET_ID', normalizedId);

  Logger.log('SPREADSHEET_IDを保存しました');
}

/**
 * 現在の管理対象スプレッドシートIDを確認する
 *
 * @returns {string}
 */
function getConfiguredSpreadsheetId() {
  const spreadsheetId = PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID') || '';

  Logger.log(
    spreadsheetId
      ? `SPREADSHEET_ID=${spreadsheetId}`
      : 'SPREADSHEET_IDは未設定です'
  );

  return spreadsheetId;
}

/**
 * 固定訓練イベントを使用して送信処理を実行する
 */
function runDrillEarthquake() {
  clearSettingsCache();

  if (getBoolSetting('test_mode', true)) {
    throw new Error(
      'runDrillEarthquakeを実行する前に、' +
      'settings.test_modeをFALSEへ変更してください。' +
      '1通だけの確認にはsendTestNotification()を使用してください。'
    );
  }

  const eventId = `DRILL-${
    Utilities.formatDate(
      new Date(),
      'Asia/Tokyo',
      'yyyyMMddHHmmss'
    )
  }`;

  processEarthquakeEvent({
    eventId,
    occurredAt: new Date(),
    announcedAt: new Date(),
    hypocenter: '安否確認訓練',
    magnitude: null,
    maxIntensity: '5-',
    sourceUrl: 'drill://manual',
  });

  Logger.log(`訓練イベントを実行しました: ${eventId}`);
}
/**
 * settingsシートをセットアップ（デフォルト値付き）
 */
function _setupSettingsSheet() {
  let sheet = getSheet(SHEET.SETTINGS);
  if (!sheet) {
    sheet = getOrCreateSheet(SHEET.SETTINGS);
    sheet.getRange(1, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    Logger.log(`シート作成: ${SHEET.SETTINGS}`);
    return;
  }

  // 既存シートの場合：欠けているキーのみ追加
  const existingData = sheet.getDataRange().getValues();
  const existingKeys = existingData.slice(1).map(row => String(row[0]).trim());

  DEFAULT_SETTINGS.slice(1).forEach(([key, value, description]) => {
    if (!existingKeys.includes(key)) {
      sheet.appendRow([key, value, description]);
      Logger.log(`settings追加: ${key}`);
    }
  });
}

/**
 * シートをセットアップ（ヘッダー確認）
 * @param {string} sheetName
 * @param {string[]} headers
 */
function _setupSheet(sheetName, headers) {
  let sheet = getSheet(sheetName);
  if (!sheet) {
    sheet = getOrCreateSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log(`シート作成: ${sheetName}`);
    return;
  }

  // 既存シートの列ヘッダー確認（不足列を末尾に追加）
  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach((h, i) => {
    if (!existingHeaders.includes(h)) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(h);
      Logger.log(`${sheetName}: 列追加 "${h}"`);
    }
  });
}

/**
 * トリガーをインストール
 */
function installTriggersAdmin() {
  clearSettingsCache();
  installTriggers();
}

/**
 * 本システムのトリガーを削除
 */
function removeManagedTriggersAdmin() {
  removeManagedTriggers();
}

/**
 * 手動で地震確認を実行（テスト・デバッグ用）
 */
function runManualCheck() {
  clearSettingsCache();
  checkEarthquakeAndNotify();
}

/**
 * test_recipient_emailへ固定テストデータで送信
 */
function sendTestNotification() {
  clearSettingsCache();

  const testRecipient = getSetting('test_recipient_email', '');
  if (!testRecipient || !isValidEmail(String(testRecipient))) {
    Logger.log('test_recipient_emailが未設定または無効です');
    return;
  }

  const testEvent = {
    eventId: `TEST-${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmm')}`,
    occurredAt: new Date(),
    announcedAt: null,
    hypocenter: 'テスト震源地（テストデータ）',
    magnitude: 5.0,
    maxIntensity: '5-',
    sourceUrl: '',
  };

  const testEmployee = {
    employeeId: 'TEST001',
    name: 'テスト太郎',
    email: testRecipient,
    active: true,
    group: 'テスト',
  };

  let formUrl;
  try {
    formUrl = buildPrefilledFormUrl(testEvent, testEmployee);
  } catch (err) {
    formUrl = '(フォームURLの設定が必要です)';
    Logger.log(`フォームURL生成スキップ: ${safeErrorMessage(err)}`);
  }

  const subject = buildEmailSubject(testEvent, true);
  const htmlBody = buildEmailHtmlBody(testEvent, testEmployee, formUrl);
  const textBody = buildEmailTextBody(testEvent, testEmployee, formUrl);

  const sendMode = String(getSetting('send_mode', SEND_MODE.GMAIL)).trim().toLowerCase();
  const result = sendNotification(
    { email: testRecipient, name: 'テスト受信者' },
    subject, htmlBody, textBody, sendMode
  );

  if (result.success) {
    Logger.log(`テスト送信成功: ${testRecipient}`);
  } else {
    Logger.log(`テスト送信失敗: ${result.error}`);
  }
}

/**
 * 最新イベントのsummaryを再生成
 */
function rebuildCurrentSummary() {
  const sheet = getSheet(SHEET.EARTHQUAKE_EVENTS);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('地震イベントが見つかりません');
    return;
  }

  // 最終行のevent_idを取得
  const lastRow = sheet.getLastRow();
  const eventId = sheet.getRange(lastRow, 1).getValue();
  if (!eventId) {
    Logger.log('最新イベントのevent_idが空です');
    return;
  }

  Logger.log(`rebuildCurrentSummary: event_id=${eventId}`);
  rebuildSummary(String(eventId));
  Logger.log('rebuildCurrentSummary: 完了');
}

/**
 * 対象イベントのfailedをpendingに戻して再送できるようにする
 * sentまたはrespondedは変更しない
 * @param {string} eventId
 */
function resendFailedNotifications(eventId) {
  if (!eventId) {
    Logger.log('event_idを指定してください');
    return;
  }

  const notifications = getNotificationsByEventId(eventId);
  let resetCount = 0;

  for (const n of notifications) {
    if (String(n.status).trim() === NOTIFICATION_STATUS.FAILED) {
      updateNotification(String(n.notification_key), {
        status: NOTIFICATION_STATUS.PENDING,
        error: '',
      });
      resetCount++;
    }
  }

  Logger.log(`resendFailedNotifications: event_id=${eventId}, ${resetCount}件をpendingへリセット`);

  if (resetCount > 0) {
    // イベントステータスも処理中に戻す
    updateEvent(eventId, { status: EVENT_STATUS.PROCESSING });
    // 再処理を実行
    processEarthquakeEvent(findEvent(eventId) ? _rowToEarthquakeEvent(findEvent(eventId)) : null);
  }
}

/**
 * earthquake_eventsの行データをEarthquakeEvent形式へ変換
 * @param {Object} row
 * @returns {EarthquakeEvent|null}
 */
function _rowToEarthquakeEvent(row) {
  if (!row) return null;
  return {
    eventId: String(row.event_id || ''),
    occurredAt: row.occurred_at ? new Date(row.occurred_at) : new Date(),
    announcedAt: row.announced_at ? new Date(row.announced_at) : null,
    hypocenter: String(row.hypocenter || ''),
    magnitude: row.magnitude !== '' ? Number(row.magnitude) : null,
    maxIntensity: String(row.max_intensity || ''),
    sourceUrl: String(row.source_url || ''),
  };
}

/**
 * 一定時間以上sendingのままの行をfailedへ変更（stale対処）
 * @param {string} eventId
 * @param {number} thresholdMinutes - デフォルト30分
 */
function resetStaleSendingNotifications(eventId, thresholdMinutes) {
  if (!eventId) {
    Logger.log('event_idを指定してください');
    return;
  }
  const threshold = thresholdMinutes || 30;
  const thresholdMs = threshold * 60 * 1000;
  const now = Date.now();

  const notifications = getNotificationsByEventId(eventId);
  let resetCount = 0;

  for (const n of notifications) {
    if (String(n.status).trim() !== NOTIFICATION_STATUS.SENDING) continue;
    const sendingAt = n.sending_at ? new Date(n.sending_at).getTime() : 0;
    if (sendingAt && (now - sendingAt) > thresholdMs) {
      updateNotification(String(n.notification_key), {
        status: NOTIFICATION_STATUS.FAILED,
        error: `STALE: sending状態が${threshold}分以上継続`,
      });
      resetCount++;
    }
  }

  Logger.log(`resetStaleSendingNotifications: event_id=${eventId}, ${resetCount}件をfailedへ更新`);
}

/**
 * システムを無効化（settings.enabled = FALSE）
 */
function disableSystem() {
  setSetting('enabled', 'FALSE');
  clearSettingsCache();
  Logger.log('システムを無効化しました (enabled=FALSE)');
}
