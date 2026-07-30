/**
 * Summary.gs
 * 安否確認の集計処理・summaryシートへの書き込み
 */

// summaryシートのレイアウト
const SUMMARY_HEADER_ROW = 1;
const SUMMARY_DATA_START_ROW = 2;

/**
 * 指定イベントのsummaryを再集計してシートへ書き込む。
 *
 * @param {string} eventId
 */
function rebuildSummary(eventId) {
  const normalizedEventId = String(eventId || '').trim();

  if (!normalizedEventId) {
    return;
  }

  const event = findEvent(normalizedEventId);
  const notifications = getNotificationsByEventId(
    normalizedEventId
  );
  const latestResponses = getLatestResponsesByEmployee(
    normalizedEventId
  );

  const targetCount = notifications.filter(notification =>
    String(notification.status || '').trim() !==
    NOTIFICATION_STATUS.SKIPPED
  ).length;

  const sentCount = notifications.filter(notification =>
    [
      NOTIFICATION_STATUS.SENT,
      NOTIFICATION_STATUS.RESPONDED,
    ].includes(
      String(notification.status || '').trim()
    )
  ).length;

  const failedCount = notifications.filter(notification =>
    String(notification.status || '').trim() ===
    NOTIFICATION_STATUS.FAILED
  ).length;

  const responses = Object.values(latestResponses);
  const responseCount = responses.length;
  const unansweredCount = Math.max(
    0,
    sentCount - responseCount
  );
  const responseRate = sentCount > 0
    ? `${Math.round((responseCount / sentCount) * 100)}%`
    : '0%';

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

  for (const response of responses) {
    const safetyStatus = String(
      response['safety_status'] || ''
    ).trim();
    const attendanceStatus = String(
      response['attendance_status'] || ''
    ).trim();

    if (safetyStatus === '被害なし') {
      safetyCount.no_damage++;
    } else if (safetyStatus === '本人にけがあり') {
      safetyCount.personal_injury++;
    } else if (safetyStatus === '家族に被害あり') {
      safetyCount.family_damage++;
    } else if (
      safetyStatus === '住居・物品に被害あり'
    ) {
      safetyCount.property_damage++;
    } else if (safetyStatus === '複数の被害あり') {
      safetyCount.multiple_damage++;
    } else if (
      safetyStatus === '現時点で回答困難'
    ) {
      safetyCount.difficult_to_respond++;
    }

    if (attendanceStatus === '出社可能') {
      attendanceCount.available++;
    } else if (
      attendanceStatus === '出社不可（災害）'
    ) {
      attendanceCount.unavailable_disaster++;
    } else if (
      attendanceStatus === '出社不可（その他）'
    ) {
      attendanceCount.unavailable_other++;
    } else {
      attendanceCount.unknown++;
    }
  }

  const summaryRow = [
    normalizedEventId,
    event ? event.occurred_at || '' : '',
    event ? event.hypocenter || '' : '',
    event ? event.max_intensity || '' : '',
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

  writeSummaryRow(normalizedEventId, summaryRow);
  writeSummaryDetails(
    normalizedEventId,
    notifications,
    latestResponses
  );
}

/**
 * summaryシートへイベント単位の集計行を書き込む。
 * 同じevent_idが存在する場合は更新し、存在しない場合は追加する。
 *
 * @param {string} eventId
 * @param {Array<*>} rowData
 */
function writeSummaryRow(eventId, rowData) {
  const sheet = getSheet(SHEET.SUMMARY);

  if (!sheet) {
    Logger.log(
      'writeSummaryRow: summaryシートが見つかりません'
    );
    return;
  }

  const headers = [
    'event_id',
    'occurred_at',
    'hypocenter',
    'max_intensity',
    'target_employee_count',
    'sent_count',
    'failed_count',
    'response_count',
    'unanswered_count',
    'response_rate',
    'no_damage_count',
    'personal_injury_count',
    'family_damage_count',
    'property_damage_count',
    'multiple_damage_count',
    'difficult_to_respond_count',
    'available_to_work_count',
    'unavailable_disaster_count',
    'unavailable_other_count',
    'last_updated_at',
  ];

  ensureSummaryHeaders(sheet, headers);

  const lastRow = sheet.getLastRow();

  if (lastRow >= SUMMARY_DATA_START_ROW) {
    const eventIds = sheet
      .getRange(
        SUMMARY_DATA_START_ROW,
        1,
        lastRow - SUMMARY_DATA_START_ROW + 1,
        1
      )
      .getValues();

    for (
      let index = 0;
      index < eventIds.length;
      index++
    ) {
      const existingEventId = String(
        eventIds[index][0] || ''
      ).trim();

      if (existingEventId === String(eventId).trim()) {
        sheet
          .getRange(
            SUMMARY_DATA_START_ROW + index,
            1,
            1,
            rowData.length
          )
          .setValues([rowData]);

        return;
      }
    }
  }

  const appendRow = Math.max(
    SUMMARY_DATA_START_ROW,
    lastRow + 1
  );

  sheet
    .getRange(appendRow, 1, 1, rowData.length)
    .setValues([rowData]);
}

/**
 * summaryシートのヘッダーを保証する。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} headers
 */
function ensureSummaryHeaders(sheet, headers) {
  const currentHeaders = sheet
    .getRange(
      SUMMARY_HEADER_ROW,
      1,
      1,
      headers.length
    )
    .getValues()[0]
    .map(header => String(header || '').trim());

  const isValid = headers.every(
    (header, index) =>
      currentHeaders[index] === header
  );

  if (!isValid) {
    sheet
      .getRange(
        SUMMARY_HEADER_ROW,
        1,
        1,
        headers.length
      )
      .setValues([headers]);
  }
}

/**
 * 最新イベントの詳細一覧を専用シートへ書き込む。
 *
 * summaryシートへ固定行で詳細を書き込むと、
 * イベント履歴が増えた際に集計表と衝突するため、
 * current_event_detailsシートへ分離する。
 *
 * @param {string} eventId
 * @param {Array<Object>} notifications
 * @param {Object<string, Object>} latestResponses
 */
function writeSummaryDetails(
  eventId,
  notifications,
  latestResponses
) {
  const sheet = getSheet(
    SHEET.CURRENT_EVENT_DETAILS
  );

  if (!sheet) {
    Logger.log(
      'writeSummaryDetails: ' +
      'current_event_detailsシートが見つかりません'
    );
    return;
  }

  const unanswered = buildUnansweredList(
    eventId,
    notifications,
    latestResponses
  );
  const failed = buildFailedList(
    eventId,
    notifications
  );
  const emergency = buildEmergencyList(
    eventId,
    latestResponses
  );

  const rows = [
    [
      'section',
      'employee_id',
      'email',
      'status',
      'comment',
    ],
    [
      'metadata',
      '',
      '',
      `event_id=${eventId}`,
      `updated_at=${nowIso()}`,
    ],
  ];

  if (unanswered.length === 0) {
    rows.push([
      'unanswered',
      '',
      '',
      'none',
      '未回答者なし',
    ]);
  } else {
    for (const item of unanswered) {
      rows.push([
        'unanswered',
        item.employee_id,
        item.email,
        'sent',
        item.sent_at
          ? `sent_at=${item.sent_at}`
          : '',
      ]);
    }
  }

  if (failed.length === 0) {
    rows.push([
      'failed',
      '',
      '',
      'none',
      '送信失敗なし',
    ]);
  } else {
    for (const item of failed) {
      rows.push([
        'failed',
        item.employee_id,
        item.email,
        'failed',
        item.error,
      ]);
    }
  }

  if (emergency.length === 0) {
    rows.push([
      'emergency',
      '',
      '',
      'none',
      '緊急確認対象なし',
    ]);
  } else {
    for (const item of emergency) {
      rows.push([
        'emergency',
        item.employee_id,
        '',
        item.safety_status,
        item.comment,
      ]);
    }
  }

  sheet.clearContents();

  sheet
    .getRange(
      1,
      1,
      rows.length,
      rows[0].length
    )
    .setValues(rows);
}

/**
 * 未回答者リストを生成する。
 *
 * sent状態かつ最新回答が存在しない通知を対象にする。
 *
 * @param {string} eventId
 * @param {Array<Object>=} notifications
 * @param {Object<string, Object>=} latestResponses
 * @returns {Array<{
 *   employee_id: string,
 *   email: string,
 *   sent_at: string
 * }>}
 */
function buildUnansweredList(
  eventId,
  notifications,
  latestResponses
) {
  const resolvedNotifications =
    notifications ||
    getNotificationsByEventId(eventId);

  const resolvedLatestResponses =
    latestResponses ||
    getLatestResponsesByEmployee(eventId);

  return resolvedNotifications
    .filter(notification =>
      String(notification.status || '').trim() ===
      NOTIFICATION_STATUS.SENT
    )
    .filter(notification => {
      const employeeId = String(
        notification.employee_id || ''
      ).trim();

      return (
        employeeId &&
        !resolvedLatestResponses[employeeId]
      );
    })
    .map(notification => ({
      employee_id: String(
        notification.employee_id || ''
      ),
      email: String(notification.email || ''),
      sent_at: formatSheetValue(
        notification.sent_at
      ),
    }));
}

/**
 * 送信失敗者リストを生成する。
 *
 * @param {string} eventId
 * @param {Array<Object>=} notifications
 * @returns {Array<{
 *   employee_id: string,
 *   email: string,
 *   error: string
 * }>}
 */
function buildFailedList(
  eventId,
  notifications
) {
  const resolvedNotifications =
    notifications ||
    getNotificationsByEventId(eventId);

  return resolvedNotifications
    .filter(notification =>
      String(notification.status || '').trim() ===
      NOTIFICATION_STATUS.FAILED
    )
    .map(notification => ({
      employee_id: String(
        notification.employee_id || ''
      ),
      email: String(notification.email || ''),
      error: String(notification.error || ''),
    }));
}

/**
 * 緊急確認対象リストを生成する。
 *
 * @param {string} eventId
 * @param {Object<string, Object>=} latestResponses
 * @returns {Array<{
 *   employee_id: string,
 *   safety_status: string,
 *   comment: string
 * }>}
 */
function buildEmergencyList(
  eventId,
  latestResponses
) {
  const resolvedLatestResponses =
    latestResponses ||
    getLatestResponsesByEmployee(eventId);

  return Object.values(resolvedLatestResponses)
    .filter(response =>
      isEmergencyResponse({
        safetyStatus: String(
          response['safety_status'] || ''
        ),
        attendanceStatus: String(
          response['attendance_status'] || ''
        ),
        comment: String(
          response['comment'] || ''
        ),
      })
    )
    .map(response => ({
      employee_id: String(
        response['employee_id'] ||
        response['employeeId'] ||
        ''
      ),
      safety_status: String(
        response['safety_status'] || ''
      ),
      comment: String(
        response['comment'] || ''
      ),
    }));
}

/**
 * スプレッドシートから取得した値を安全に文字列化する。
 *
 * @param {*} value
 * @returns {string}
 */
function formatSheetValue(value) {
  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return '';
  }

  if (
    Object.prototype.toString.call(value) ===
    '[object Date]'
  ) {
    const date = new Date(value);

    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(
        date,
        'Asia/Tokyo',
        'yyyy-MM-dd HH:mm:ss'
      );
    }
  }

  return String(value);
}