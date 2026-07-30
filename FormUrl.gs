/**
 * FormUrl.gs
 * Googleフォーム事前入力URLの生成
 */

/**
 * 社員・イベント情報をもとにGoogleフォームの事前入力URLを生成
 *
 * @param {EarthquakeEvent} event
 * @param {Employee} employee
 * @returns {string} 事前入力URL
 */
function buildPrefilledFormUrl(event, employee) {
  const baseUrl = getSetting('form_base_url', '');
  if (!baseUrl || String(baseUrl).trim() === '') {
    throw new Error(`${ERROR_CATEGORY.CONFIG_MISSING}: form_base_urlが設定されていません`);
  }

  const eventEntryId = getSetting('form_event_entry_id', '');
  const employeeIdEntryId = getSetting('form_employee_id_entry_id', '');
  const nameEntryId = getSetting('form_name_entry_id', '');

  if (!eventEntryId || !employeeIdEntryId || !nameEntryId) {
    throw new Error(`${ERROR_CATEGORY.CONFIG_MISSING}: フォームentry IDが不足しています`);
  }

  // URLクエリパラメータを構築
  const params = [
    `${encodeURIComponent(String(eventEntryId).trim())}=${encodeURIComponent(event.eventId)}`,
    `${encodeURIComponent(String(employeeIdEntryId).trim())}=${encodeURIComponent(employee.employeeId)}`,
    `${encodeURIComponent(String(nameEntryId).trim())}=${encodeURIComponent(employee.name)}`,
  ];

  const base = String(baseUrl).trim().replace(/[?&]+$/, '');
  return `${base}?${params.join('&')}`;
}
