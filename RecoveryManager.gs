/**
 * RecoveryManager.gs
 * 未完了通知の再開、stale sendingの回復、イベント集計値の再計算
 */

/**
 * 一定時間以上sending状態の通知をfailedへ変更する。
 *
 * sending直後にApps Scriptが強制終了した場合、
 * 実際には送信済みの可能性があるため、自動でpendingには戻さない。
 * 管理者が送信状況を確認したうえで再送する。
 *
 * @returns {number} failedへ変更した件数
 */
function resetAllStaleSendingNotifications() {
  const sheet = getSheet(SHEET.NOTIFICATION_STATUS);

  if (!sheet || sheet.getLastRow() < 2) {
    return 0;
  }

  const thresholdMinutes = Math.max(
    1,
    getNumSetting('stale_sending_minutes', 30)
  );
  const thresholdMs = thresholdMinutes * 60 * 1000;
  const now = Date.now();

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(header => String(header).trim());

  const notificationKeyIndex = headers.indexOf('notification_key');
  const statusIndex = headers.indexOf('status');
  const sendingAtIndex = headers.indexOf('sending_at');

  if (
    notificationKeyIndex < 0 ||
    statusIndex < 0 ||
    sendingAtIndex < 0
  ) {
    throw new Error(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      'notification_statusシートの必須列が不足しています'
    );
  }

  let resetCount = 0;

  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const status = String(data[rowIndex][statusIndex] || '').trim();

    if (status !== NOTIFICATION_STATUS.SENDING) {
      continue;
    }

    const sendingAtValue = data[rowIndex][sendingAtIndex];
    const sendingAtMs = new Date(sendingAtValue).getTime();

    if (!Number.isFinite(sendingAtMs)) {
      continue;
    }

    if (now - sendingAtMs <= thresholdMs) {
      continue;
    }

    const notificationKey = String(
      data[rowIndex][notificationKeyIndex] || ''
    ).trim();

    if (!notificationKey) {
      continue;
    }

    updateNotification(notificationKey, {
      status: NOTIFICATION_STATUS.FAILED,
      error:
        `${ERROR_CATEGORY.STALE_SENDING}: ` +
        `${thresholdMinutes}分以上sending状態`,
    });

    resetCount++;
  }

  if (resetCount > 0) {
    Logger.log(
      `resetAllStaleSendingNotifications: ` +
      `${resetCount}件をfailedへ変更`
    );
  }

  return resetCount;
}

/**
 * processing状態のイベントに残っているpending通知を再開する。
 *
 * @returns {{
 *   processedEvents: number,
 *   sent: number,
 *   failed: number,
 *   incompleteEvents: number
 * }}
 */
function resumePendingNotifications() {
  const eventSheet = getSheet(SHEET.EARTHQUAKE_EVENTS);

  const resultSummary = {
    processedEvents: 0,
    sent: 0,
    failed: 0,
    incompleteEvents: 0,
  };

  if (!eventSheet || eventSheet.getLastRow() < 2) {
    return resultSummary;
  }

  const data = eventSheet.getDataRange().getValues();
  const headers = data[0].map(header => String(header).trim());

  const eventIdIndex = headers.indexOf('event_id');
  const statusIndex = headers.indexOf('status');

  if (eventIdIndex < 0 || statusIndex < 0) {
    throw new Error(
      `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
      'earthquake_eventsシートの必須列が不足しています'
    );
  }

  const isTestMode = getBoolSetting('test_mode', true);

  if (isTestMode) {
    Logger.log(
      'resumePendingNotifications: ' +
      'test_mode=TRUEのため実通知の再開をスキップ'
    );
    return resultSummary;
  }

  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const eventId = String(data[rowIndex][eventIdIndex] || '').trim();
    const status = String(data[rowIndex][statusIndex] || '').trim();

    if (!eventId || status !== EVENT_STATUS.PROCESSING) {
      continue;
    }

    const eventRecord = findEvent(eventId);
    const event = convertStoredEventToEarthquakeEvent(eventRecord);

    if (!event) {
      appendSendError({
        eventId,
        error:
          `${ERROR_CATEGORY.EARTHQUAKE_PARSE_FAILED}: ` +
          '保存済みイベントを復元できません',
      });
      continue;
    }

    const pendingBefore = getPendingNotifications(eventId);

    if (pendingBefore.length === 0) {
      updateEventAggregateStatus(eventId, false);
      continue;
    }

    const sendResult = processPendingNotifications(
      eventId,
      event,
      false
    );

    resultSummary.processedEvents++;
    resultSummary.sent += sendResult.sent;
    resultSummary.failed += sendResult.failed;

    if (sendResult.incomplete) {
      resultSummary.incompleteEvents++;
    }

    updateEventAggregateStatus(eventId, sendResult.incomplete);

    if (getBoolSetting('summary_enabled', true)) {
      try {
        rebuildSummary(eventId);
      } catch (error) {
        Logger.log(
          `resumePendingNotifications summary生成失敗: ` +
          `${safeErrorMessage(error)}`
        );
      }
    }

    if (sendResult.incomplete) {
      break;
    }
  }

  return resultSummary;
}

/**
 * earthquake_eventsの保存行をEarthquakeEventへ変換する。
 *
 * @param {Object|null} eventRecord
 * @returns {EarthquakeEvent|null}
 */
function convertStoredEventToEarthquakeEvent(eventRecord) {
  if (!eventRecord) {
    return null;
  }

  const eventId = String(eventRecord.event_id || '').trim();

  if (!eventId) {
    return null;
  }

  const occurredAt = new Date(eventRecord.occurred_at);

  if (isNaN(occurredAt.getTime())) {
    return null;
  }

  let announcedAt = null;

  if (eventRecord.announced_at) {
    const parsedAnnouncedAt = new Date(eventRecord.announced_at);

    if (!isNaN(parsedAnnouncedAt.getTime())) {
      announcedAt = parsedAnnouncedAt;
    }
  }

  const magnitudeValue =
    eventRecord.magnitude === '' ||
    eventRecord.magnitude === null ||
    eventRecord.magnitude === undefined
      ? null
      : Number(eventRecord.magnitude);

  return {
    eventId,
    occurredAt,
    announcedAt,
    hypocenter: String(eventRecord.hypocenter || ''),
    magnitude:
      magnitudeValue !== null && Number.isFinite(magnitudeValue)
        ? magnitudeValue
        : null,
    maxIntensity: String(eventRecord.max_intensity || ''),
    sourceUrl: String(eventRecord.source_url || ''),
  };
}

/**
 * notification_statusを基準にイベントの合計値と状態を更新する。
 *
 * @param {string} eventId
 * @param {boolean} incomplete
 * @returns {{
 *   status: string,
 *   sentCount: number,
 *   failedCount: number,
 *   pendingCount: number,
 *   sendingCount: number
 * }}
 */
function updateEventAggregateStatus(eventId, incomplete) {
  const notifications = getNotificationsByEventId(eventId);

  const sentCount = notifications.filter(notification => {
    const status = String(notification.status || '').trim();

    return [
      NOTIFICATION_STATUS.SENT,
      NOTIFICATION_STATUS.RESPONDED,
    ].includes(status);
  }).length;

  const failedCount = notifications.filter(notification =>
    String(notification.status || '').trim() ===
    NOTIFICATION_STATUS.FAILED
  ).length;

  const pendingCount = notifications.filter(notification =>
    String(notification.status || '').trim() ===
    NOTIFICATION_STATUS.PENDING
  ).length;

  const sendingCount = notifications.filter(notification =>
    String(notification.status || '').trim() ===
    NOTIFICATION_STATUS.SENDING
  ).length;

  let finalStatus;

  if (incomplete || pendingCount > 0 || sendingCount > 0) {
    finalStatus = EVENT_STATUS.PROCESSING;
  } else if (failedCount === 0) {
    finalStatus = EVENT_STATUS.COMPLETED;
  } else if (sentCount > 0) {
    finalStatus = EVENT_STATUS.PARTIAL_FAILED;
  } else {
    finalStatus = EVENT_STATUS.FAILED;
  }

  updateEvent(eventId, {
    status: finalStatus,
    sent_count: sentCount,
    failed_count: failedCount,
    completed_at:
      finalStatus === EVENT_STATUS.PROCESSING
        ? ''
        : nowIso(),
  });

  return {
    status: finalStatus,
    sentCount,
    failedCount,
    pendingCount,
    sendingCount,
  };
}