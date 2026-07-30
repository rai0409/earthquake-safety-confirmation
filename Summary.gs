/**
 * Summary.gs
 * 安否確認の集計処理・summaryシートへの書き込み
 */

// summaryシートのレイアウト（行番号）
const SUMMARY_HEADER_ROW = 1;
const SUMMARY_DATA_START_ROW = 2;

/**
 * 指定イベントのsummaryを再集計してシートへ書き込む
 * @param {string} eventId
 */
function rebuildSummary(eventId) {
  if (!eventId) return;

  const event = findEvent(eventId);
  const notifications = getNotificationsByEventId(eventId);
  const latestResponses = getLatestResponsesByEmployee(eventId);

  // 集計計算
  const targetCount = notifications.filter(n =>
    String(n.status).trim() !== NOTIFICATION_STATUS.SKIPPED
  ).length;

  const sentCount = notifications.filter(n =>
    [NOTIFICATION_STATUS.SENT, NOTIFICATION_STATUS.RESPONDED].includes(String(n.status).trim())
  ).length;

  const failedCount = notifications.filter(n =>
    String(n.status).trim() === NOTIFICATION_STATUS.FAILED
  ).length;

  const responseCount = Object.keys(latestResponses).length;
  const unansweredCount = Math.max(0, sentCount - responseCount);
  const responseRate = sentCount > 0
    ? `${Math.round((responseCount / sentCount) * 100)}%`
    : '0%';

  // 安否状況別集計
  const responses = Object.values(latestResponses);
  const safetyCount = {
    no_damage: 0,
    personal_injury: 0,
    family_damage: 0,
    property_damage: 0,
    multiple_damage: 0,
    difficult_to_respond: 0,
  };
  const attendanceCount = {
    available: 0,
    unavailable_disaster: 0,
    unavailable_other: 0,
    unknown: 0,
  };

  for (const r of responses) {
    const safety = String(r['safety_status'] || '').trim();
    const attendance = String(r['attendance_status'] || '').trim();

    // 安否状況
    if (safety === '被害なし') safetyCount.no_damage++;
    else if (safety === '本人にけがあり') safetyCount.personal_injury++;
    else if (safety === '家族に被害あり') safetyCount.family_damage++;
    else if (safety === '住居・物品に被害あり') safetyCount.property_damage++;
    else if (safety === '複数の被害あり') safetyCount.multiple_damage++;
    else if (safety === '現時点で回答困難') safetyCount.difficult_to_respond++;

    // 出社可否
    if (attendance === '出社可能') attendanceCount.available++;
    else if (attendance === '出社不可（災害）') attendanceCount.unavailable_disaster++;
    else if (attendance === '出社不可（その他）') attendanceCount.unavailable_other++;
    else attendanceCount.unknown++;
  }

  const summaryRow = [
    eventId,
    event ? (event.occurred_at || '') : '',
    event ? (event.hypocenter || '') : '',
    event ? (event.max_intensity || '') : '',
    targetCount,
    sentCount,
    failedCount,
    responseCount,
    unansweredCount,
    responseRate,
    safetyCount.no_damage,
    safetyCount.personal_injury,
    safetyCount.family_damage,
    safetyCount.property_damage,
    safetyCount.multiple_damage,
    safetyCount.difficult_to_respond,
    attendanceCount.available,
    attendanceCount.unavailable_disaster,
    attendanceCount.unavailable_other,
    nowIso(),
  ];

  writeSummaryRow(eventId, summaryRow);

  // 詳細一覧の書き込み
  writeSummaryDetails(eventId, notifications, latestResponses);
}

/**
 * summaryシートへ集計行を書き込む
 * @param {string} eventId
 * @param {Array} rowData
 */
function writeSummaryRow(eventId, rowData) {
  const sheet = getSheet(SHEET.SUMMARY);
  if (!sheet) return;

  const headers = [
    'event_id', 'occurred_at', 'hypocenter', 'max_intensity',
    'target_employee_count', 'sent_count', 'failed_count',
    'response_count', 'unanswered_count', 'response_rate',
    'no_damage_count', 'personal_injury_count', 'family_damage_count',
    'property_damage_count', 'multiple_damage_count', 'difficult_to_respond_count',
    'available_to_work_count', 'unavailable_disaster_count', 'unavailable_other_count',
    'last_updated_at',
  ];

  // ヘッダー行の確認・設定
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (!firstRow[0] || String(firstRow[0]).trim() !== 'event_id') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // 既存行を検索
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < Math.min(data.length, 200); i++) {
    if (String(data[i][0]).trim() === String(eventId).trim()) {
      sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
      return;
    }
  }

  // 新規追加
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, rowData.length).setValues([rowData]);
}

/**
 * summaryシートの下部へ詳細一覧を書き込む
 * @param {string} eventId
 * @param {Array} notifications
 * @param {Object} latestResponses
 */
function writeSummaryDetails(eventId, notifications, latestResponses) {
  const sheet = getSheet(SHEET.SUMMARY);
  if (!sheet) return;

  // 詳細セクションの開始行（集計表の下に余白を空けて）
  const detailStartRow = 25;

  // 未回答者一覧
  const unanswered = buildUnansweredList(eventId, notifications, latestResponses);
  // 送信失敗者一覧
  const failed = buildFailedList(eventId, notifications);
  // 緊急確認対象一覧
  const emergency = buildEmergencyList(eventId, latestResponses);

  let currentRow = detailStartRow;

  // セクションヘッダー付きで書き込む
  const sections = [
    { title: `■ 未回答者一覧 (${unanswered.length}名)`, data: unanswered, headers: ['employee_id', 'email', 'sent_at'] },
    { title: `■ 送信失敗者一覧 (${failed.length}名)`, data: failed, headers: ['employee_id', 'email', 'error'] },
    { title: `■ 緊急確認対象一覧 (${emergency.length}名)`, data: emergency, headers: ['employee_id', 'safety_status', 'comment'] },
  ];

  for (const section of sections) {
    sheet.getRange(currentRow, 1).setValue(section.title);
    currentRow++;
    if (section.data.length > 0) {
      sheet.getRange(currentRow, 1, 1, section.headers.length).setValues([section.headers]);
      currentRow++;
      const rows = section.data.map(item => section.headers.map(h => item[h] || ''));
      sheet.getRange(currentRow, 1, rows.length, section.headers.length).setValues(rows);
      currentRow += rows.length;
    } else {
      sheet.getRange(currentRow, 1).setValue('(なし)');
      currentRow++;
    }
    currentRow++; // 空行
  }
}

/**
 * 未回答者リストを生成
 * @param {string} eventId
 * @param {Array} notifications
 * @param {Object} latestResponses
 * @returns {Array}
 */
function buildUnansweredList(eventId, notifications, latestResponses) {
  if (!notifications) notifications = getNotificationsByEventId(eventId);
  if (!latestResponses) latestResponses = getLatestResponsesByEmployee(eventId);

  return notifications
    .filter(n => String(n.status).trim() === NOTIFICATION_STATUS.SENT)
    .filter(n => !latestResponses[String(n.employee_id)])
    .map(n => ({
      employee_id: String(n.employee_id),
      email: String(n.email),
      sent_at: String(n.sent_at || ''),
    }));
}

/**
 * 送信失敗者リストを生成
 * @param {string} eventId
 * @param {Array} notifications
 * @returns {Array}
 */
function buildFailedList(eventId, notifications) {
  if (!notifications) notifications = getNotificationsByEventId(eventId);

  return notifications
    .filter(n => String(n.status).trim() === NOTIFICATION_STATUS.FAILED)
    .map(n => ({
      employee_id: String(n.employee_id),
      email: String(n.email),
      error: String(n.error || ''),
    }));
}

/**
 * 緊急確認対象リストを生成
 * @param {string} eventId
 * @param {Object} latestResponses
 * @returns {Array}
 */
function buildEmergencyList(eventId, latestResponses) {
  if (!latestResponses) latestResponses = getLatestResponsesByEmployee(eventId);

  return Object.values(latestResponses)
    .filter(r => isEmergencyResponse({
      safetyStatus: String(r['safety_status'] || ''),
      comment: String(r['comment'] || ''),
    }))
    .map(r => ({
      employee_id: String(r['employee_id'] || ''),
      safety_status: String(r['safety_status'] || ''),
      comment: String(r['comment'] || ''),
    }));
}
