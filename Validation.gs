/**
 * Validation.gs
 * 設定値・社員データの検証
 */

/**
 * システム全体の設定を検証して問題一覧を返す
 * @returns {{ issues: string[], ok: boolean }}
 */
function validateConfiguration() {
  const issues = [];

  const settingsIssues = validateSettings();
  issues.push(...settingsIssues);

  const employeeIssues = validateEmployees();
  issues.push(...employeeIssues);

  const formIssues = validateFormConfiguration();
  issues.push(...formIssues);

  const triggerIssues = validateTriggers();
  issues.push(...triggerIssues);

  const responseSheetIssues = validateFormResponseSheet();
  issues.push(...responseSheetIssues);

  if (issues.length > 0) {
    Logger.log('設定検証結果: 問題あり');
    issues.forEach(issue => Logger.log(`  - ${issue}`));
  } else {
    Logger.log('設定検証結果: 問題なし');
  }

  return { issues, ok: issues.length === 0 };
}

/**
 * settingsシートの必須設定を検証
 * @returns {string[]}
 */
function validateSettings() {
  const issues = [];

  const adminEmail = getSetting('admin_email', '');
  if (!adminEmail || !isValidEmail(String(adminEmail))) {
    issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: admin_emailが未設定または無効`);
  }

  const testRecipient = getSetting('test_recipient_email', '');
  if (!testRecipient || !isValidEmail(String(testRecipient))) {
    issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: test_recipient_emailが未設定または無効`);
  }

  const sendMode = String(getSetting('send_mode', '')).trim().toLowerCase();
  if (sendMode !== SEND_MODE.GMAIL && sendMode !== SEND_MODE.OUTLOOK) {
    issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: send_modeが無効 ("gmail"または"outlook"を設定)`);
  }

  const threshold = getSetting('threshold', '');
  if (!threshold || intensityToRank(String(threshold).trim()) === null) {
    issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: thresholdが無効 (例: "5-")`);
  }

  // Outfookモード時のScript Properties確認
  if (sendMode === SEND_MODE.OUTLOOK) {
    const msValidation = validateMicrosoftConfiguration();
    if (!msValidation.ok) {
      issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: Outlook用Script Propertiesが未設定: ${msValidation.missing.join(', ')}`);
    }
  }

  // Gmailモード時のquota確認（0より大きければOK）
  if (sendMode === SEND_MODE.GMAIL) {
    try {
      const quota = MailApp.getRemainingDailyQuota();
      if (quota <= 0) {
        issues.push(`${ERROR_CATEGORY.GMAIL_QUOTA_INSUFFICIENT}: Gmail送信残数が0です`);
      }
    } catch (err) {
      issues.push(`Gmail quota取得失敗: ${safeErrorMessage(err)}`);
    }
  }

  return issues;
}

/**
 * 社員シートを検証
 * @returns {string[]}
 */
function validateEmployees() {
  const issues = [];

  const sheet = getSheet(SHEET.EMPLOYEES);
  if (!sheet) {
    issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: employeesシートが見つかりません`);
    return issues;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    issues.push('employeesシートにデータがありません');
    return issues;
  }

  // 重複チェック
  const { duplicates } = checkDuplicateEmployeeIds();
  if (duplicates.length > 0) {
    issues.push(`employee_idが重複しています: ${duplicates.join(', ')}`);
  }

  // 全行の検証
  let invalidCount = 0;
  let activeCount = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowStr = row.map(c => String(c).trim()).join('');
    if (!rowStr) continue;

    const validation = validateEmployee(row);
    if (!validation.valid) {
      invalidCount++;
    } else if (parseBoolean(row[EMP_COL.ACTIVE])) {
      activeCount++;
    }
  }

  if (invalidCount > 0) {
    issues.push(`社員データに${invalidCount}件の不正行があります(スクリプトログを確認)`);
  }
  if (activeCount === 0) {
    issues.push('アクティブな社員が0人です');
  }

  return issues;
}

/**
 * フォーム設定を検証
 * @returns {string[]}
 */
function validateFormConfiguration() {
  const issues = [];

  const formBaseUrl = getSetting('form_base_url', '');
  if (!formBaseUrl || String(formBaseUrl).trim() === '') {
    issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: form_base_urlが未設定`);
  }

  const eventEntryId = getSetting('form_event_entry_id', '');
  if (!eventEntryId) {
    issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: form_event_entry_idが未設定`);
  }

  const employeeIdEntryId = getSetting('form_employee_id_entry_id', '');
  if (!employeeIdEntryId) {
    issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: form_employee_id_entry_idが未設定`);
  }

  const nameEntryId = getSetting('form_name_entry_id', '');
  if (!nameEntryId) {
    issues.push(`${ERROR_CATEGORY.CONFIG_MISSING}: form_name_entry_idが未設定`);
  }

  return issues;
}

/**
 * トリガーの存在を検証
 * @returns {string[]}
 */
function validateTriggers() {
  const issues = [];

  if (!hasManagedTrigger(TRIGGER_HANDLER_EARTHQUAKE)) {
    issues.push(`トリガー未設定: ${TRIGGER_HANDLER_EARTHQUAKE} (installTriggers()を実行してください)`);
  }

  if (!hasManagedTrigger(TRIGGER_HANDLER_FORM)) {
    issues.push(`トリガー未設定: ${TRIGGER_HANDLER_FORM} (installTriggers()を実行するかGUIで設定してください)`);
  }

  return issues;
}

/**
 * Googleフォーム回答先シートの存在を確認
 * @returns {string[]}
 */
function validateFormResponseSheet() {
  const issues = [];
  const sheetName = String(
    getSetting(
      'form_response_sheet_name',
      'フォームの回答 1'
    )
  ).trim();
  const sheet = getFormResponseSheet();
  if (!sheet) {
    issues.push(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      `Googleフォーム回答先シート "${sheetName}" が見つかりません`
    );
  }
  return issues;
}
