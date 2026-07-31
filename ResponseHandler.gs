/**
 * ResponseHandler.gs
 * Googleフォーム回答の受信・処理
 */

// 緊急キーワード（自由記述の参考判定のみ）
const EMERGENCY_KEYWORDS = ['救助', '救急', '動けない', '閉じ込め', '重傷'];

/**
 * onFormSubmitトリガーのハンドラ
 * フォーム回答を処理してnotification_statusを更新し、summaryを再集計
 * @param {Object} e - フォーム送信イベント
 */
function handleFormSubmit(e) {
  try {
    if (!e || !e.namedValues) {
      Logger.log(`${ERROR_CATEGORY.FORM_RESPONSE_INVALID}: イベントオブジェクトが無効`);
      return;
    }

    // フォームの列名はシステムの設定に依存するため、設定ファイルから取得できるようにする
    // デフォルト列名はGoogleフォームに準拠
    const colEventId = getSetting('form_col_event_id', 'event_id');
    const colEmployeeId = getSetting('form_col_employee_id', 'employee_id');
    const colName = getSetting('form_col_name', 'name');
    const colSafetyStatus = getSetting('form_col_safety_status', 'safety_status');
    const colAttendanceStatus = getSetting('form_col_attendance_status', 'attendance_status');
    const colComment = getSetting('form_col_comment', 'comment');

    const getValue = (colName) => {
      const arr = e.namedValues[colName];
      return arr && arr.length > 0 ? String(arr[0]).trim() : '';
    };

    const eventId = getValue(colEventId);
    const employeeId = getValue(colEmployeeId);
    const name = getValue(colName);
    const safetyStatus = getValue(colSafetyStatus);
    const attendanceStatus = getValue(colAttendanceStatus);
    const comment = getValue(colComment);

    if (!eventId || !employeeId) {
      Logger.log(
        `${ERROR_CATEGORY.FORM_RESPONSE_INVALID}: ` +
        'event_idまたはemployee_idが空'
      );
      return;
    }
    const eventRecord = findEvent(eventId);

    if (!eventRecord) {
      Logger.log(
        `${ERROR_CATEGORY.FORM_RESPONSE_INVALID}: ` +
        `未登録event_id "${eventId}"`
      );
      return;
    }

    const notificationKey = generateNotificationKey(eventId, employeeId);
    const existing = findNotification(notificationKey);

    if (existing) {
      updateNotification(notificationKey, {
        status: NOTIFICATION_STATUS.RESPONDED,
        responded_at: nowIso(),
      });
    } else {
      Logger.log(
        `${ERROR_CATEGORY.EMPLOYEE_NOT_FOUND}: ` +
        `notification_key "${notificationKey}" が見つかりません`
      );
      return;
    }

    // 緊急判定
    const response = { safetyStatus, attendanceStatus, comment, name, employeeId, eventId };
    if (isEmergencyResponse(response)) {
      const immediateAlert = getBoolSetting('immediate_admin_alert', true);
      if (immediateAlert) {
        const alertMsg = `[緊急通知] employee_id: ${employeeId} / 名前: ${name}\n安否状況: ${safetyStatus}\nコメント: ${comment || '(なし)'}`;
        notifyAdmin(alertMsg);
      }
    }

    // summaryを再集計
    try {
      rebuildSummary(eventId);
    } catch (summaryErr) {
      Logger.log(`summary再集計エラー: ${safeErrorMessage(summaryErr)}`);
    }

  } catch (err) {
    Logger.log(`handleFormSubmit エラー: ${safeErrorMessage(err)}`);
  }
}

/**
 * 緊急回答かどうかを判定
 * @param {Object} response
 * @returns {boolean}
 */
function isEmergencyResponse(response) {
  // 本人けがあり または 複数の被害あり
  const emergencySafetyStatuses = ['本人にけがあり', '複数の被害あり'];
  if (emergencySafetyStatuses.includes(response.safetyStatus)) return true;

  // コメントのキーワード判定（参考情報として扱う）
  if (response.comment) {
    const comment = String(response.comment);
    for (const keyword of EMERGENCY_KEYWORDS) {
      if (comment.includes(keyword)) return true;
    }
  }

  return false;
}

/**
 * 特定イベントの最新回答を社員ごとに1件だけ取得
 * 同じemployeeIdから複数回答がある場合は最新を有効とする
 * @param {string} eventId
 * @returns {Object} employeeId -> response オブジェクト
 */
function getLatestResponsesByEmployee(eventId) {
  const responses = getFormResponsesByEventId(eventId);
  const latestByEmployee = {};

  // タイムスタンプ順にソート（昇順）
  responses.sort((a, b) => {
    const ta = new Date(a['timestamp'] || 0).getTime();
    const tb = new Date(b['timestamp'] || 0).getTime();
    return ta - tb;
  });

  for (const r of responses) {
    const empId = String(r['employee_id'] || r['employeeId'] || '').trim();
    if (empId) {
      latestByEmployee[empId] = r; // 後のもので上書き → 最新が残る
    }
  }

  return latestByEmployee;
}
