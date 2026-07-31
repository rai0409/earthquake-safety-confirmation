/**
 * Notification.gs
 * 通知レコードの作成・送信処理・メール本文生成
 */

/**
 * 1通分のメール件名を生成
 * @param {EarthquakeEvent} event
 * @param {boolean} isTest
 * @returns {string}
 */
function buildEmailSubject(event, isTest) {
  const prefix = isTest ? '【テスト】' : '';
  return `${prefix}【安否確認 ${event.eventId}】災害時安否確認のお願い`;
}

/**
 * メールHTML本文を生成
 * @param {EarthquakeEvent} event
 * @param {Employee} employee
 * @param {string} formUrl
 * @returns {string}
 */
function buildEmailHtmlBody(event, employee, formUrl) {
  const occurredAt = event.occurredAt
    ? Utilities.formatDate(event.occurredAt, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm')
    : '不明';
  const magnitude = event.magnitude !== null && event.magnitude !== undefined
    ? `M${event.magnitude}`
    : '不明';

  return `<!DOCTYPE html>
<html lang="ja">
<head><meta charset="UTF-8"></head>
<body style="font-family:sans-serif;line-height:1.7;color:#333;">
<p>${escapeHtml(employee.name)} さん</p>
<p>お疲れ様です。安否確認担当者です。</p>
<p>先ほど地震が発生しました。皆さんの安全を確認するため、以下の安否確認フォームへご回答をお願いします。</p>

<table style="border-collapse:collapse;margin:16px 0;">
  <tr>
    <td style="padding:4px 12px;font-weight:bold;white-space:nowrap;">発生日時</td>
    <td style="padding:4px 12px;">${escapeHtml(occurredAt)}</td>
  </tr>
  <tr>
    <td style="padding:4px 12px;font-weight:bold;white-space:nowrap;">震源地</td>
    <td style="padding:4px 12px;">${escapeHtml(event.hypocenter || '不明')}</td>
  </tr>
  <tr>
    <td style="padding:4px 12px;font-weight:bold;white-space:nowrap;">最大震度</td>
    <td style="padding:4px 12px;">震度 ${escapeHtml(event.maxIntensity || '不明')}</td>
  </tr>
  <tr>
    <td style="padding:4px 12px;font-weight:bold;white-space:nowrap;">マグニチュード</td>
    <td style="padding:4px 12px;">${escapeHtml(magnitude)}</td>
  </tr>
</table>

<p style="margin:24px 0;">
  <a href="${formUrl}" style="background:#d32f2f;color:#fff;padding:12px 24px;text-decoration:none;border-radius:4px;font-weight:bold;">
    安否確認フォームへ回答する
  </a>
</p>
<p style="font-size:0.9em;color:#666;">
  ボタンが開かない場合は以下のURLをブラウザに貼り付けてください：<br>
  <span style="word-break:break-all;">${escapeHtml(formUrl)}</span>
</p>

<hr style="border:none;border-top:1px solid #ddd;margin:24px 0;">
<p style="font-size:0.85em;color:#888;">
  ※ このメールへの返信では回答が記録されません。必ず上のフォームからご回答ください。<br>
  ※ 可能な範囲で速やかにご回答いただけますようお願いします。
</p>
</body>
</html>`;
}

/**
 * メールプレーンテキスト本文を生成
 * @param {EarthquakeEvent} event
 * @param {Employee} employee
 * @param {string} formUrl
 * @returns {string}
 */
function buildEmailTextBody(event, employee, formUrl) {
  const occurredAt = event.occurredAt
    ? Utilities.formatDate(event.occurredAt, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm')
    : '不明';
  const magnitude = event.magnitude !== null ? `M${event.magnitude}` : '不明';

  return [
    `${employee.name} さん`,
    '',
    'お疲れ様です。安否確認担当者です。',
    '',
    '地震が発生しました。以下の安否確認フォームへご回答をお願いします。',
    '',
    `発生日時: ${occurredAt}`,
    `震源地: ${event.hypocenter || '不明'}`,
    `最大震度: 震度 ${event.maxIntensity || '不明'}`,
    `マグニチュード: ${magnitude}`,
    '',
    '■ 安否確認フォームURL',
    formUrl,
    '',
    '※ このメールへの返信では回答が記録されません。必ず上のフォームからご回答ください。',
    '※ 可能な範囲で速やかにご回答いただけますようお願いします。',
  ].join('\n');
}

/**
 * 地震イベントに対して社員全員の通知レコードをpendingで作成
 * 既存のnotification_keyがある場合はスキップ
 * @param {EarthquakeEvent} event
 * @param {Employee[]} employees
 * @param {string} channel - 'gmail' | 'outlook'
 * @returns {{ created: number, skipped: number }}
 */
function createNotificationRecords(event, employees, channel) {
  let created = 0;
  let skipped = 0;

  for (const emp of employees) {
    const notificationKey = generateNotificationKey(event.eventId, emp.employeeId);
    const existing = findNotification(notificationKey);

    if (existing) {
      const st = String(existing.status).trim();
      if ([NOTIFICATION_STATUS.SENT, NOTIFICATION_STATUS.RESPONDED].includes(st)) {
        skipped++;
        continue;
      }
    }

    if (!existing) {
      appendNotification({
        notificationKey,
        eventId: event.eventId,
        employeeId: emp.employeeId,
        name: emp.name,
        email: emp.email,
        channel,
        status: NOTIFICATION_STATUS.PENDING,
        attempts: 0,
        createdAt: nowIso(),
      });
      created++;
    }
  }

  return { created, skipped };
}

/**
 * pending通知を処理
 * Apps Script実行時間を考慮して途中で打ち切る
 * @param {string} eventId
 * @param {EarthquakeEvent} event
 * @param {boolean} isTestMode
 * @returns {{ sent: number, failed: number, incomplete: boolean }}
 */
function processPendingNotifications(eventId, event, isTestMode) {
  const sendMode = String(getSetting('send_mode', SEND_MODE.GMAIL)).trim().toLowerCase();
  if (sendMode !== SEND_MODE.GMAIL && sendMode !== SEND_MODE.OUTLOOK) {
    throw new Error(`不正なsend_mode: "${sendMode}". "gmail" または "outlook" を設定してください`);
  }

  const allPendingList = getPendingNotifications(eventId);
  const pendingList = isTestMode
    ? allPendingList.slice(0, 1)
    : allPendingList;
  const testRecipient = getSetting('test_recipient_email', '');

  let sentCount = 0;
  let failedCount = 0;
  const startTime = Date.now();

  // Gmail quota事前確認
  if (!isTestMode && sendMode === SEND_MODE.GMAIL) {
    const quotaCheck = validateGmailQuota(pendingList.length);
    if (!quotaCheck.ok) {
      const errMsg = `${ERROR_CATEGORY.GMAIL_QUOTA_INSUFFICIENT}: 残り${quotaCheck.remaining}通、必要${pendingList.length}通`;
      Logger.log(errMsg);
      appendSendError({
        eventId,
        channel: SEND_MODE.GMAIL,
        error: errMsg,
      });
      notifyAdmin(`[${ERROR_CATEGORY.GMAIL_QUOTA_INSUFFICIENT}] Gmailの本日の送信残数が不足しています。残${quotaCheck.remaining}通、必要${pendingList.length}通`);
      // 全通知をfailedへ
      for (const n of pendingList) {
        updateNotification(n.notification_key, {
          status: NOTIFICATION_STATUS.FAILED,
          error: errMsg,
        });
      }
      return { sent: 0, failed: pendingList.length, incomplete: false };
    }
  }

  // Outlook設定確認
  if (!isTestMode && sendMode === SEND_MODE.OUTLOOK) {
    const msValidation = validateMicrosoftConfiguration();
    if (!msValidation.ok) {
      const errMsg = `${ERROR_CATEGORY.CONFIG_MISSING}: ${msValidation.missing.join(', ')}`;
      Logger.log(errMsg);
      notifyAdmin(`[CONFIG_MISSING] Outlook送信に必要なScript Propertiesが未設定です: ${msValidation.missing.join(', ')}`);
      for (const n of pendingList) {
        updateNotification(n.notification_key, {
          status: NOTIFICATION_STATUS.FAILED,
          error: errMsg,
        });
      }
      return { sent: 0, failed: pendingList.length, incomplete: false };
    }
  }

  for (const n of pendingList) {
    // Apps Script実行時間チェック（残り30秒を切ったら中断）
    const elapsedSeconds = (Date.now() - startTime) / 1000;
    if (elapsedSeconds > MAX_EXECUTION_SECONDS - EXECUTION_TIME_BUFFER_SECONDS) {
      Logger.log(`実行時間制限のため、pending通知${pendingList.length - sentCount - failedCount}件を残して中断`);
      return { sent: sentCount, failed: failedCount, incomplete: true };
    }

    const notificationKey = String(n.notification_key).trim();
    const employeeEmail = String(n.email).trim();

    // 送信中マーク
    updateNotification(notificationKey, {
      status: NOTIFICATION_STATUS.SENDING,
      sending_at: nowIso(),
    });

    // 送信先の決定（テストモードはtest_recipient_emailのみ）
    let actualRecipient;
    if (isTestMode) {
      if (!testRecipient || !isValidEmail(testRecipient)) {
        failedCount++;
        Logger.log(`テストモード: test_recipient_emailが未設定のためスキップ`);
        updateNotification(notificationKey, {
          status: NOTIFICATION_STATUS.SKIPPED,
          error: 'test_recipient_emailが未設定',
        });
        continue;
      }
      actualRecipient = { email: testRecipient, name: 'テスト受信者' };
    } else {
      actualRecipient = {
        email: employeeEmail,
        name: String(n.name || n.employee_id),
      };
    }

    // フォームURLの生成
    let formUrl;
    const tempEmployee = {
      employeeId: String(n.employee_id),
      name: String(n.name || n.employee_id),
      email: employeeEmail,
    };
    try {
      formUrl = buildPrefilledFormUrl(event, tempEmployee);
    } catch (err) {
      const errMsg = safeErrorMessage(err);
      Logger.log(`フォームURL生成失敗 employee_id=${n.employee_id}: ${errMsg}`);
      updateNotification(notificationKey, {
        status: NOTIFICATION_STATUS.FAILED,
        error: errMsg,
      });
      failedCount++;
      continue;
    }

    const subject = buildEmailSubject(event, isTestMode);
    const htmlBody = buildEmailHtmlBody(event, tempEmployee, formUrl);
    const textBody = buildEmailTextBody(event, tempEmployee, formUrl);

    // 送信
    const result = sendNotification(actualRecipient, subject, htmlBody, textBody, sendMode);

    const currentAttempts = (Number(n.attempts) || 0) + 1;

    if (result.success) {
      // テストモードの場合は実社員のstatusをsentにしない
      if (isTestMode) {
        updateNotification(notificationKey, {
          status: NOTIFICATION_STATUS.PENDING, // testではpendingのまま
          attempts: currentAttempts,
        });
      } else {
        updateNotification(notificationKey, {
          status: NOTIFICATION_STATUS.SENT,
          sent_at: nowIso(),
          attempts: currentAttempts,
          error: '',
        });
      }
      sentCount++;
    } else {
      updateNotification(notificationKey, {
        status: NOTIFICATION_STATUS.FAILED,
        attempts: currentAttempts,
        error: result.error || '送信失敗',
      });
      appendSendError({
        timestamp: nowIso(),
        eventId,
        employeeId: n.employee_id,
        email: employeeEmail,
        channel: sendMode,
        attempt: currentAttempts,
        responseCode: result.responseCode || '',
        error: result.error || '送信失敗',
      });
      failedCount++;
    }
  }

  return { sent: sentCount, failed: failedCount, incomplete: false };
}

/**
 * 送信モードに応じた送信処理
 * GmailとOutlookを同時に使用しない
 * @param {Object} recipient
 * @param {string} subject
 * @param {string} htmlBody
 * @param {string} textBody
 * @param {string} sendMode
 * @returns {{ success: boolean, responseCode: number|null, error: string|null }}
 */
function sendNotification(recipient, subject, htmlBody, textBody, sendMode) {
  if (sendMode === SEND_MODE.GMAIL) {
    return sendMailGmail(recipient, subject, htmlBody, textBody);
  } else if (sendMode === SEND_MODE.OUTLOOK) {
    const result = sendMailOutlook(recipient, subject, htmlBody);
    return result;
  } else {
    return { success: false, responseCode: null, error: `不正なsend_mode: ${sendMode}` };
  }
}

/**
 * 管理者へ通知メールを送信（MailApp使用）
 * @param {string} body
 */
function notifyAdmin(body) {
  try {
    const adminEmail = getSetting('admin_email', '');
    if (!adminEmail || !isValidEmail(adminEmail)) {
      Logger.log('notifyAdmin: admin_emailが未設定のためスキップ');
      return;
    }
    const senderName = getSetting('sender_display_name', '災害時安否確認');
    MailApp.sendEmail({
      to: adminEmail,
      subject: '[安否確認システム] 管理者通知',
      body: body,
      name: senderName,
    });
  } catch (err) {
    Logger.log(`管理者通知失敗: ${safeErrorMessage(err)}`);
  }
}
