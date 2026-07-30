/**
 * Config.gs
 * 設定値の取得・キャッシュ
 */

// シート名定数
const SHEET = {
  SETTINGS: 'settings',
  EMPLOYEES: 'employees',
  EARTHQUAKE_EVENTS: 'earthquake_events',
  NOTIFICATION_STATUS: 'notification_status',
  SEND_ERRORS: 'send_errors',
  FORM_RESPONSES: 'form_responses',
  SUMMARY: 'summary',
};

// earthquake_events status 定数
const EVENT_STATUS = {
  DETECTED: 'detected',
  PROCESSING: 'processing',
  COMPLETED: 'completed',
  PARTIAL_FAILED: 'partial_failed',
  FAILED: 'failed',
  SKIPPED: 'skipped',
  TEST_COMPLETED: 'test_completed',
};

// notification_status status 定数
const NOTIFICATION_STATUS = {
  PENDING: 'pending',
  SENDING: 'sending',
  SENT: 'sent',
  FAILED: 'failed',
  RESPONDED: 'responded',
  SKIPPED: 'skipped',
};

// 送信モード定数
const SEND_MODE = {
  GMAIL: 'gmail',
  OUTLOOK: 'outlook',
};

// エラーカテゴリ定数
const ERROR_CATEGORY = {
  CONFIG_MISSING: 'CONFIG_MISSING',
  EARTHQUAKE_FETCH_FAILED: 'EARTHQUAKE_FETCH_FAILED',
  EARTHQUAKE_PARSE_FAILED: 'EARTHQUAKE_PARSE_FAILED',
  BELOW_THRESHOLD: 'BELOW_THRESHOLD',
  EVENT_TOO_OLD: 'EVENT_TOO_OLD',
  DUPLICATE_EVENT: 'DUPLICATE_EVENT',
  GMAIL_QUOTA_INSUFFICIENT: 'GMAIL_QUOTA_INSUFFICIENT',
  OUTLOOK_TOKEN_FAILED: 'OUTLOOK_TOKEN_FAILED',
  OUTLOOK_PERMISSION_DENIED: 'OUTLOOK_PERMISSION_DENIED',
  OUTLOOK_RATE_LIMITED: 'OUTLOOK_RATE_LIMITED',
  OUTLOOK_SEND_FAILED: 'OUTLOOK_SEND_FAILED',
  FORM_RESPONSE_INVALID: 'FORM_RESPONSE_INVALID',
  EMPLOYEE_NOT_FOUND: 'EMPLOYEE_NOT_FOUND',
};

// Apps Script実行時間余裕（秒）
const EXECUTION_TIME_BUFFER_SECONDS = 30;
const MAX_EXECUTION_SECONDS = 360; // 6分（上限6分）

// 設定キャッシュ（スクリプト実行中の再取得を避ける）
let _settingsCache = null;

/**
 * settingsシートから全設定を取得してキャッシュ
 * @returns {Object} key→value のマップ
 */
function getSettings() {
  if (_settingsCache !== null) return _settingsCache;
  const sheet = getSheet(SHEET.SETTINGS);
  if (!sheet) throw new Error('settingsシートが見つかりません');
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const value = data[i][1];
    if (key) map[key] = value;
  }
  _settingsCache = map;
  return map;
}

/**
 * 設定値を1件取得
 * @param {string} key
 * @param {*} defaultValue
 * @returns {*}
 */
function getSetting(key, defaultValue) {
  const settings = getSettings();
  if (key in settings) return settings[key];
  return defaultValue !== undefined ? defaultValue : null;
}

/**
 * boolean設定値を取得
 * @param {string} key
 * @param {boolean} defaultValue
 * @returns {boolean}
 */
function getBoolSetting(key, defaultValue) {
  const val = getSetting(key, null);
  if (val === null) return defaultValue;
  return parseBoolean(val);
}

/**
 * 数値設定値を取得
 * @param {string} key
 * @param {number} defaultValue
 * @returns {number}
 */
function getNumSetting(key, defaultValue) {
  const val = getSetting(key, null);
  if (val === null) return defaultValue;
  const n = Number(val);
  return isNaN(n) ? defaultValue : n;
}

/**
 * キャッシュをクリア（テスト等で使用）
 */
function clearSettingsCache() {
  _settingsCache = null;
}
