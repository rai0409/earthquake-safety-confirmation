/**
 * Main.gs
 * エントリポイント・メイン処理
 */

/**
 * 地震確認と安否確認送信のメイン関数
 * time-drivenトリガーから呼び出される
 */
function checkEarthquakeAndNotify() {
  // LockServiceによる重複実行防止
  const lock = LockService.getScriptLock();
  const lockAcquired = lock.tryLock(5000);
  if (!lockAcquired) {
    Logger.log('checkEarthquakeAndNotify: 別の処理が実行中のためスキップ');
    return;
  }

  try {
    clearSettingsCache();

    const enabled = getBoolSetting('enabled', false);
    if (!enabled) {
      Logger.log('システムが無効です (enabled=FALSE)');
      return;
    }

    // 地震情報の取得
    let earthquake;
    try {
      earthquake = fetchLatestEarthquake();
    } catch (err) {
      Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_FETCH_FAILED}: ${safeErrorMessage(err)}`);
      notifyAdmin(`[${ERROR_CATEGORY.EARTHQUAKE_FETCH_FAILED}] 地震情報の取得に失敗しました: ${safeErrorMessage(err)}`);
      appendSendError({
        eventId: '',
        channel: '',
        error: `${ERROR_CATEGORY.EARTHQUAKE_FETCH_FAILED}: ${safeErrorMessage(err)}`,
      });
      return;
    }

    if (!earthquake) {
      Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_FETCH_FAILED}: 地震情報を取得できませんでした`);
      return;
    }

    // 送信判定
    const { ok, reason } = shouldNotifyEarthquake(earthquake);
    if (!ok) {
      Logger.log(`送信対象外: ${reason}`);
      return;
    }

    // イベント処理
    processEarthquakeEvent(earthquake);

  } finally {
    lock.releaseLock();
  }
}

/**
 * 地震イベントを処理して安否確認メールを送信する
 * @param {EarthquakeEvent} event
 */
function processEarthquakeEvent(event) {
  if (!event || !event.eventId) {
    Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_PARSE_FAILED}: eventが無効`);
    return;
  }

  const isTestMode = getBoolSetting('test_mode', true);
  const sendMode = String(getSetting('send_mode', SEND_MODE.GMAIL)).trim().toLowerCase();

  Logger.log(`地震処理開始: eventId=${event.eventId}, 震度=${event.maxIntensity}, testMode=${isTestMode}, sendMode=${sendMode}`);

  // 社員一覧を取得
  let employees;
  let invalidCount;
  try {
    const result = getActiveEmployees();
    employees = result.employees;
    invalidCount = result.invalidCount;
  } catch (err) {
    const errMsg = safeErrorMessage(err);
    Logger.log(`社員一覧取得失敗: ${errMsg}`);
    appendSendError({ eventId: event.eventId, channel: sendMode, error: errMsg });
    notifyAdmin(`[社員一覧取得失敗] ${errMsg}`);
    return;
  }

  if (invalidCount > 0) {
    Logger.log(`社員データに${invalidCount}件の不正行があります`);
    notifyAdmin(`[社員データ警告] 不正な社員行が${invalidCount}件あります。スクリプトログを確認してください。`);
  }

  if (employees.length === 0) {
    Logger.log('アクティブな社員が0人のため送信をスキップ');
    appendSendError({ eventId: event.eventId, channel: sendMode, error: 'アクティブな社員が0人' });
    return;
  }

  Logger.log(`送信対象社員: ${employees.length}名 (不正行: ${invalidCount}件)`);

  // イベントをearthquake_eventsへ記録
  const existingEvent = findEvent(event.eventId);
  if (!existingEvent) {
    appendEvent({
      ...event,
      detectedAt: nowIso(),
      status: EVENT_STATUS.DETECTED,
      targetCount: employees.length,
      sentCount: 0,
      failedCount: 0,
    });
  } else {
    updateEvent(event.eventId, {
      status: EVENT_STATUS.PROCESSING,
      target_count: employees.length,
    });
  }

  // 通知レコードを作成
  const { created, skipped } = createNotificationRecords(event, employees, sendMode);
  Logger.log(`通知レコード: 新規作成=${created}, スキップ=${skipped}`);

  // pending通知を送信
  let sentCount = 0;
  let failedCount = 0;
  let incomplete = false;

  try {
    const result = processPendingNotifications(event.eventId, event, isTestMode);
    sentCount = result.sent;
    failedCount = result.failed;
    incomplete = result.incomplete;
  } catch (err) {
    const errMsg = safeErrorMessage(err);
    Logger.log(`通知送信中にエラー: ${errMsg}`);
    appendSendError({ eventId: event.eventId, channel: sendMode, error: errMsg });
    notifyAdmin(`[送信エラー] ${errMsg}`);
  }

  Logger.log(`送信結果: 送信済=${sentCount}, 失敗=${failedCount}, 未完了=${incomplete}`);

  // イベントステータスを更新
  let finalStatus;
  if (isTestMode) {
    finalStatus = EVENT_STATUS.TEST_COMPLETED;
  } else if (incomplete) {
    finalStatus = EVENT_STATUS.PROCESSING; // 次回トリガーで続きを処理
  } else if (failedCount === 0) {
    finalStatus = EVENT_STATUS.COMPLETED;
  } else if (sentCount > 0) {
    finalStatus = EVENT_STATUS.PARTIAL_FAILED;
  } else {
    finalStatus = EVENT_STATUS.FAILED;
  }

  updateEvent(event.eventId, {
    status: finalStatus,
    sent_count: sentCount,
    failed_count: failedCount,
    completed_at: incomplete ? '' : nowIso(),
  });

  // 失敗がある場合は管理者へ通知
  if (failedCount > 0) {
    notifyAdmin(`[送信失敗] event_id=${event.eventId}: ${failedCount}件の送信が失敗しました。resendFailedNotifications()で再送できます。`);
  }

  // summaryを生成
  if (getBoolSetting('summary_enabled', true) && !isTestMode) {
    try {
      rebuildSummary(event.eventId);
    } catch (err) {
      Logger.log(`summary生成エラー: ${safeErrorMessage(err)}`);
    }
  }

  Logger.log(`地震処理完了: eventId=${event.eventId}, status=${finalStatus}`);
}
