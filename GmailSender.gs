/**
 * GmailSender.gs
 * Gmail（MailApp）を使ったメール送信
 */

/**
 * Gmail quota が十分か確認
 * @param {number} requiredCount
 * @returns {{ ok: boolean, remaining: number }}
 */
function validateGmailQuota(requiredCount) {
  const remaining = MailApp.getRemainingDailyQuota();
  return {
    ok: remaining >= requiredCount,
    remaining,
  };
}

/**
 * Gmailでメールを1通送信
 * @param {Object} recipient - { email, name }
 * @param {string} subject
 * @param {string} htmlBody
 * @param {string} textBody
 * @returns {{ success: boolean, error: string|null }}
 */
function sendMailGmail(recipient, subject, htmlBody, textBody) {
  try {
    const senderName = getSetting('sender_display_name', '災害時安否確認');
    MailApp.sendEmail({
      to: recipient.email,
      subject: subject,
      htmlBody: htmlBody,
      body: textBody,
      name: senderName,
    });
    return { success: true, error: null };
  } catch (err) {
    return { success: false, error: safeErrorMessage(err) };
  }
}
