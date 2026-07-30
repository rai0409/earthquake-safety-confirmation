/**
 * Utils.gs
 * 汎用ユーティリティ関数
 */

/**
 * 現在時刻のISO文字列（Asia/Tokyo）
 * @returns {string}
 */
function nowIso() {
  return new Date().toISOString();
}

/**
 * 文字列をbooleanへ変換
 * @param {*} value
 * @returns {boolean}
 */
function parseBoolean(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    return value.trim().toUpperCase() === 'TRUE';
  }
  return false;
}

/**
 * メールアドレスを正規化（trim + lowercase）
 * @param {string} email
 * @returns {string}
 */
function normalizeEmail(email) {
  if (!email || typeof email !== 'string') return '';
  return email.trim().toLowerCase();
}

/**
 * HTMLエスケープ
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * 指定秒数スリープ（Apps Script実行時間に注意）
 * @param {number} seconds
 */
function sleepSeconds(seconds) {
  Utilities.sleep(seconds * 1000);
}

/**
 * notification_key を生成
 * @param {string} eventId
 * @param {string} employeeId
 * @returns {string}
 */
function generateNotificationKey(eventId, employeeId) {
  return `${eventId}:${employeeId}:initial`;
}

/**
 * メールアドレス形式を簡易検証
 * @param {string} email
 * @returns {boolean}
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const normalized = normalizeEmail(email);
  // RFC5322の簡易版
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized);
}

/**
 * 値が空（null / undefined / 空文字 / 空白のみ）か判定
 * @param {*} value
 * @returns {boolean}
 */
function isEmpty(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

/**
 * スタックトレースを含まない安全なエラーメッセージを返す
 * @param {Error} err
 * @returns {string}
 */
function safeErrorMessage(err) {
  if (!err) return 'Unknown error';
  return err.message || String(err);
}
