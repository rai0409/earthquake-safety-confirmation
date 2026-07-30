/**
 * OutlookSender.gs
 * Microsoft Graph API を使ったOutlookメール送信
 * Power Automate は使用しない
 */

const GRAPH_TOKEN_CACHE_KEY = 'ms_graph_access_token';
const GRAPH_TOKEN_EXPIRY_KEY = 'ms_graph_token_expires_at';
// トークンはキャッシュに保存するが有効期限より短くする（余裕を1分持たせる）
const TOKEN_EXPIRY_BUFFER_SECONDS = 60;

// Graph API エンドポイント
const GRAPH_SEND_MAIL_PATH = 'https://graph.microsoft.com/v1.0/users/{senderEmail}/sendMail';
const GRAPH_TOKEN_URL = 'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token';

// 再試行対象ステータスコード
const RETRY_STATUS_CODES = [429, 500, 502, 503, 504];
// 再試行しないステータスコード
const NO_RETRY_STATUS_CODES = [400, 401, 403, 404];

/**
 * Outlook送信に必要なScript Propertiesを検証
 * @returns {{ ok: boolean, missing: string[] }}
 */
function validateMicrosoftConfiguration() {
  const props = PropertiesService.getScriptProperties();
  const required = ['MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET', 'MS_SENDER_EMAIL'];
  const missing = required.filter(key => !props.getProperty(key));
  return { ok: missing.length === 0, missing };
}

/**
 * Microsoft Graph API アクセストークンを取得
 * CacheServiceへの保存はアクセストークンのみ（client secretは保存しない）
 * @returns {string|null} アクセストークン or null
 */
function getMicrosoftAccessToken() {
  const cache = CacheService.getScriptCache();

  // キャッシュから取得を試みる
  const cachedToken = cache.get(GRAPH_TOKEN_CACHE_KEY);
  if (cachedToken) return cachedToken;

  const props = PropertiesService.getScriptProperties();
  const tenantId = props.getProperty('MS_TENANT_ID');
  const clientId = props.getProperty('MS_CLIENT_ID');
  const clientSecret = props.getProperty('MS_CLIENT_SECRET');

  if (!tenantId || !clientId || !clientSecret) {
    Logger.log(`${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: ${ERROR_CATEGORY.CONFIG_MISSING} - Script Propertiesが不足しています`);
    return null;
  }

  const tokenUrl = GRAPH_TOKEN_URL.replace('{tenantId}', encodeURIComponent(tenantId));

  const payload = {
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  };

  const options = {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: Object.entries(payload)
      .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
      .join('&'),
    muteHttpExceptions: true,
  };

  let response;
  try {
    response = UrlFetchApp.fetch(tokenUrl, options);
  } catch (err) {
    Logger.log(`${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: ネットワークエラー - ${safeErrorMessage(err)}`);
    return null;
  }

  const code = response.getResponseCode();
  if (code !== 200) {
    // secretやtokenをログへ出さない
    Logger.log(`${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: HTTP ${code}`);
    return null;
  }

  let tokenData;
  try {
    tokenData = JSON.parse(response.getContentText());
  } catch (err) {
    Logger.log(`${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: レスポンス解析失敗`);
    return null;
  }

  const accessToken = tokenData['access_token'];
  if (!accessToken) {
    Logger.log(`${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: access_tokenが取得できませんでした`);
    return null;
  }

  // 有効期限（expiresIn）よりバッファを引いた時間だけキャッシュ
  const expiresIn = tokenData['expires_in'] || 3600;
  const cacheDuration = Math.max(60, expiresIn - TOKEN_EXPIRY_BUFFER_SECONDS);
  // CacheServiceの上限は21600秒
  const cacheSeconds = Math.min(cacheDuration, 21600);
  cache.put(GRAPH_TOKEN_CACHE_KEY, accessToken, cacheSeconds);

  return accessToken;
}

/**
 * 再試行付きHTTPリクエスト
 * @param {string} url
 * @param {Object} options - UrlFetchApp.fetch オプション
 * @param {Object} retryConfig - { maxAttempts, waitSeconds }
 * @returns {{ response: HTTPResponse|null, attempts: number, error: string|null }}
 */
function fetchWithRetry(url, options, retryConfig) {
  const maxAttempts = retryConfig.maxAttempts || 3;
  const waitSeconds = retryConfig.waitSeconds || 3;
  let lastError = null;
  let response = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      response = UrlFetchApp.fetch(url, options);
    } catch (err) {
      lastError = safeErrorMessage(err);
      Logger.log(`fetchWithRetry 試行${attempt}/${maxAttempts}: ネットワークエラー - ${lastError}`);
      if (attempt < maxAttempts) sleepSeconds(waitSeconds);
      continue;
    }

    const code = response.getResponseCode();

    if (code === 202 || (code >= 200 && code < 300)) {
      return { response, attempts: attempt, error: null };
    }

    // 再試行しないコード
    if (NO_RETRY_STATUS_CODES.includes(code)) {
      Logger.log(`fetchWithRetry: HTTP ${code} は再試行対象外`);
      return { response, attempts: attempt, error: `HTTP ${code}` };
    }

    // 再試行対象コード
    if (RETRY_STATUS_CODES.includes(code)) {
      lastError = `HTTP ${code}`;
      Logger.log(`fetchWithRetry 試行${attempt}/${maxAttempts}: HTTP ${code}`);

      // 429 Retry-After ヘッダーを可能な範囲で尊重
      if (code === 429 && attempt < maxAttempts) {
        const headers = response.getAllHeaders();
        const retryAfterStr = headers['Retry-After'] || headers['retry-after'];
        const retryAfterSeconds = retryAfterStr ? parseInt(retryAfterStr, 10) : waitSeconds;
        // Apps Scriptの実行時間を超えないよう上限を設ける
        const actualWait = Math.min(retryAfterSeconds, 30);
        Logger.log(`429 Retry-After: ${actualWait}秒待機`);
        sleepSeconds(actualWait);
        continue;
      }

      if (attempt < maxAttempts) sleepSeconds(waitSeconds);
      continue;
    }

    // その他のステータスコード
    lastError = `HTTP ${code}`;
    Logger.log(`fetchWithRetry: 予期しないステータスコード ${code}`);
    return { response, attempts: attempt, error: lastError };
  }

  return { response, attempts: maxAttempts, error: lastError };
}

/**
 * Microsoft Graph API でメールを1通送信
 * @param {Object} recipient - { email, name }
 * @param {string} subject
 * @param {string} htmlBody
 * @returns {{ success: boolean, responseCode: number|null, error: string|null }}
 */
function sendMailOutlook(recipient, subject, htmlBody) {
  const validation = validateMicrosoftConfiguration();
  if (!validation.ok) {
    return {
      success: false,
      responseCode: null,
      error: `${ERROR_CATEGORY.CONFIG_MISSING}: ${validation.missing.join(', ')}`,
    };
  }

  const accessToken = getMicrosoftAccessToken();
  if (!accessToken) {
    return {
      success: false,
      responseCode: null,
      error: ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED,
    };
  }

  const props = PropertiesService.getScriptProperties();
  const senderEmail = props.getProperty('MS_SENDER_EMAIL');
  // settingsのoutlook_sender_emailもフォールバックとして使用可能
  const effectiveSenderEmail = senderEmail || getSetting('outlook_sender_email', '');
  if (!effectiveSenderEmail) {
    return {
      success: false,
      responseCode: null,
      error: `${ERROR_CATEGORY.CONFIG_MISSING}: MS_SENDER_EMAILが設定されていません`,
    };
  }

  const endpoint = GRAPH_SEND_MAIL_PATH.replace('{senderEmail}', encodeURIComponent(effectiveSenderEmail));

  const mailPayload = {
    message: {
      subject: subject,
      body: {
        contentType: 'HTML',
        content: htmlBody,
      },
      toRecipients: [
        {
          emailAddress: {
            address: recipient.email,
            name: recipient.name || '',
          },
        },
      ],
    },
    saveToSentItems: true,
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
    payload: JSON.stringify(mailPayload),
    muteHttpExceptions: true,
  };

  const retryConfig = {
    maxAttempts: getNumSetting('retry_max_attempts', 3),
    waitSeconds: getNumSetting('retry_wait_seconds', 3),
  };

  const { response, attempts, error } = fetchWithRetry(endpoint, options, retryConfig);

  if (!response) {
    return {
      success: false,
      responseCode: null,
      error: error || ERROR_CATEGORY.OUTLOOK_SEND_FAILED,
    };
  }

  const code = response.getResponseCode();
  Logger.log(`Outlook送信: ${recipient.email} HTTP ${code} 試行${attempts}回`);

  if (code === 202) {
    return { success: true, responseCode: code, error: null };
  }

  // エラーカテゴリの分類
  let errorCategory = ERROR_CATEGORY.OUTLOOK_SEND_FAILED;
  if (code === 401 || code === 403) {
    errorCategory = ERROR_CATEGORY.OUTLOOK_PERMISSION_DENIED;
    // 401の場合はトークンキャッシュをクリア
    if (code === 401) {
      CacheService.getScriptCache().remove(GRAPH_TOKEN_CACHE_KEY);
    }
  } else if (code === 429) {
    errorCategory = ERROR_CATEGORY.OUTLOOK_RATE_LIMITED;
  }

  return {
    success: false,
    responseCode: code,
    error: `${errorCategory}: HTTP ${code} (試行${attempts}回)`,
  };
}
