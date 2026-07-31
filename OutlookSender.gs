/**
 * OutlookSender.gs
 * Microsoft Graph APIを使ったOutlookメール送信
 * Power Automateは使用しない
 */

const GRAPH_TOKEN_CACHE_KEY =
  'ms_graph_access_token';

// トークン有効期限より短くキャッシュするための余裕
const TOKEN_EXPIRY_BUFFER_SECONDS = 60;

// Microsoft Graph APIエンドポイント
const GRAPH_SEND_MAIL_PATH =
  'https://graph.microsoft.com/v1.0/users/{senderEmail}/sendMail';

const GRAPH_TOKEN_URL =
  'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token';

// 再試行対象ステータスコード
const RETRY_STATUS_CODES = [
  429,
  500,
  502,
  503,
  504,
];

// 再試行しないステータスコード
const NO_RETRY_STATUS_CODES = [
  400,
  401,
  403,
  404,
];

/**
 * Outlook送信に必要なScript Propertiesを検証する。
 *
 * @returns {{
 *   ok: boolean,
 *   missing: string[]
 * }}
 */
function validateMicrosoftConfiguration() {
  const properties =
    PropertiesService.getScriptProperties();

  const requiredKeys = [
    'MS_TENANT_ID',
    'MS_CLIENT_ID',
    'MS_CLIENT_SECRET',
    'MS_SENDER_EMAIL',
  ];

  const missing = requiredKeys.filter(key => {
    const value = properties.getProperty(key);
    return !value || !String(value).trim();
  });

  return {
    ok: missing.length === 0,
    missing,
  };
}

/**
 * Microsoft Graph APIのアクセストークンを取得する。
 *
 * client secretはScript Propertiesから読み取り、
 * CacheServiceにはアクセストークンだけを保存する。
 *
 * @returns {string|null}
 */
function getMicrosoftAccessToken() {
  const cache = CacheService.getScriptCache();

  const cachedToken = cache.get(
    GRAPH_TOKEN_CACHE_KEY
  );

  if (cachedToken) {
    return cachedToken;
  }

  const properties =
    PropertiesService.getScriptProperties();

  const tenantId = String(
    properties.getProperty('MS_TENANT_ID') || ''
  ).trim();
  const clientId = String(
    properties.getProperty('MS_CLIENT_ID') || ''
  ).trim();
  const clientSecret = String(
    properties.getProperty('MS_CLIENT_SECRET') || ''
  ).trim();

  if (
    !tenantId ||
    !clientId ||
    !clientSecret
  ) {
    Logger.log(
      `${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: ` +
      `${ERROR_CATEGORY.CONFIG_MISSING} - ` +
      'Script Propertiesが不足しています'
    );

    return null;
  }

  const tokenUrl = GRAPH_TOKEN_URL.replace(
    '{tenantId}',
    encodeURIComponent(tenantId)
  );

  const payload = {
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  };

  const options = {
    method: 'post',
    contentType:
      'application/x-www-form-urlencoded',
    payload: Object.entries(payload)
      .map(([key, value]) =>
        `${encodeURIComponent(key)}=` +
        `${encodeURIComponent(value)}`
      )
      .join('&'),
    muteHttpExceptions: true,
  };

  let response;

  try {
    response = UrlFetchApp.fetch(
      tokenUrl,
      options
    );
  } catch (error) {
    Logger.log(
      `${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: ` +
      `ネットワークエラー - ` +
      `${safeErrorMessage(error)}`
    );

    return null;
  }

  const responseCode =
    response.getResponseCode();

  if (responseCode !== 200) {
    Logger.log(
      `${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: ` +
      `HTTP ${responseCode}`
    );

    return null;
  }

  let tokenData;

  try {
    tokenData = JSON.parse(
      response.getContentText()
    );
  } catch (error) {
    Logger.log(
      `${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: ` +
      'レスポンス解析失敗'
    );

    return null;
  }

  const accessToken = String(
    tokenData['access_token'] || ''
  ).trim();

  if (!accessToken) {
    Logger.log(
      `${ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED}: ` +
      'access_tokenが取得できませんでした'
    );

    return null;
  }

  const parsedExpiresIn = Number(
    tokenData['expires_in']
  );
  const expiresIn = Number.isFinite(
    parsedExpiresIn
  )
    ? parsedExpiresIn
    : 3600;

  const cacheDuration = Math.max(
    60,
    expiresIn -
      TOKEN_EXPIRY_BUFFER_SECONDS
  );

  // CacheServiceの保存上限は21600秒
  const cacheSeconds = Math.min(
    cacheDuration,
    21600
  );

  cache.put(
    GRAPH_TOKEN_CACHE_KEY,
    accessToken,
    cacheSeconds
  );

  return accessToken;
}

/**
 * Microsoft Graphのアクセストークンキャッシュを削除する。
 */
function clearMicrosoftAccessTokenCache() {
  CacheService
    .getScriptCache()
    .remove(GRAPH_TOKEN_CACHE_KEY);
}

/**
 * 再試行付きでHTTPリクエストを実行する。
 *
 * @param {string} url
 * @param {Object} options
 * @param {{
 *   maxAttempts: number,
 *   waitSeconds: number
 * }} retryConfig
 * @returns {{
 *   response: GoogleAppsScript.URL_Fetch.HTTPResponse|null,
 *   attempts: number,
 *   error: string|null
 * }}
 */
function fetchWithRetry(
  url,
  options,
  retryConfig
) {
  const configuredMaxAttempts = Number(
    retryConfig &&
    retryConfig.maxAttempts
  );
  const configuredWaitSeconds = Number(
    retryConfig &&
    retryConfig.waitSeconds
  );

  const maxAttempts = Math.max(
    1,
    Number.isFinite(configuredMaxAttempts)
      ? Math.floor(configuredMaxAttempts)
      : 3
  );

  const waitSeconds = Math.max(
    0,
    Number.isFinite(configuredWaitSeconds)
      ? configuredWaitSeconds
      : 3
  );

  let lastError = null;
  let lastResponse = null;

  for (
    let attempt = 1;
    attempt <= maxAttempts;
    attempt++
  ) {
    try {
      lastResponse = UrlFetchApp.fetch(
        url,
        options
      );
    } catch (error) {
      lastError = safeErrorMessage(error);

      Logger.log(
        `fetchWithRetry 試行` +
        `${attempt}/${maxAttempts}: ` +
        `ネットワークエラー - ${lastError}`
      );

      if (attempt < maxAttempts) {
        sleepSeconds(waitSeconds);
      }

      continue;
    }

    const responseCode =
      lastResponse.getResponseCode();

    if (
      responseCode >= 200 &&
      responseCode < 300
    ) {
      return {
        response: lastResponse,
        attempts: attempt,
        error: null,
      };
    }

    if (
      NO_RETRY_STATUS_CODES.includes(
        responseCode
      )
    ) {
      Logger.log(
        `fetchWithRetry: HTTP ${responseCode} ` +
        'は再試行対象外'
      );

      return {
        response: lastResponse,
        attempts: attempt,
        error: `HTTP ${responseCode}`,
      };
    }

    if (
      RETRY_STATUS_CODES.includes(
        responseCode
      )
    ) {
      lastError = `HTTP ${responseCode}`;

      Logger.log(
        `fetchWithRetry 試行` +
        `${attempt}/${maxAttempts}: ` +
        `HTTP ${responseCode}`
      );

      if (attempt >= maxAttempts) {
        break;
      }

      if (responseCode === 429) {
        const retryAfterSeconds =
          getRetryAfterSeconds(
            lastResponse,
            waitSeconds
          );

        const actualWaitSeconds = Math.min(
          retryAfterSeconds,
          30
        );

        Logger.log(
          `429 Retry-After: ` +
          `${actualWaitSeconds}秒待機`
        );

        sleepSeconds(actualWaitSeconds);
      } else {
        sleepSeconds(waitSeconds);
      }

      continue;
    }

    lastError = `HTTP ${responseCode}`;

    Logger.log(
      'fetchWithRetry: ' +
      `予期しないステータスコード ` +
      `${responseCode}`
    );

    return {
      response: lastResponse,
      attempts: attempt,
      error: lastError,
    };
  }

  return {
    response: lastResponse,
    attempts: maxAttempts,
    error: lastError,
  };
}

/**
 * Retry-Afterヘッダーから待機秒数を取得する。
 *
 * 秒数形式とHTTP-date形式の両方に対応する。
 *
 * @param {GoogleAppsScript.URL_Fetch.HTTPResponse} response
 * @param {number} fallbackSeconds
 * @returns {number}
 */
function getRetryAfterSeconds(
  response,
  fallbackSeconds
) {
  const headers = response.getAllHeaders();

  const rawValue =
    headers['Retry-After'] ||
    headers['retry-after'];

  if (
    rawValue === null ||
    rawValue === undefined ||
    rawValue === ''
  ) {
    return Math.max(0, fallbackSeconds);
  }

  const numericSeconds = Number(rawValue);

  if (
    Number.isFinite(numericSeconds) &&
    numericSeconds >= 0
  ) {
    return numericSeconds;
  }

  const retryDateMs = new Date(
    String(rawValue)
  ).getTime();

  if (Number.isFinite(retryDateMs)) {
    return Math.max(
      0,
      Math.ceil(
        (retryDateMs - Date.now()) / 1000
      )
    );
  }

  return Math.max(0, fallbackSeconds);
}

/**
 * Microsoft Graph APIでメールを1通送信する。
 *
 * @param {{
 *   email: string,
 *   name?: string
 * }} recipient
 * @param {string} subject
 * @param {string} htmlBody
 * @returns {{
 *   success: boolean,
 *   responseCode: number|null,
 *   error: string|null
 * }}
 */
function sendMailOutlook(
  recipient,
  subject,
  htmlBody
) {
  const recipientEmail = String(
    recipient && recipient.email
      ? recipient.email
      : ''
  ).trim();

  if (
    !recipientEmail ||
    !isValidEmail(recipientEmail)
  ) {
    return {
      success: false,
      responseCode: null,
      error: '送信先メールアドレスが無効です',
    };
  }

  const validation =
    validateMicrosoftConfiguration();

  if (!validation.ok) {
    return {
      success: false,
      responseCode: null,
      error:
        `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
        `${validation.missing.join(', ')}`,
    };
  }

  const properties =
    PropertiesService.getScriptProperties();

  const senderEmail = String(
    properties.getProperty(
      'MS_SENDER_EMAIL'
    ) || ''
  ).trim();

  if (
    !senderEmail ||
    !isValidEmail(senderEmail)
  ) {
    return {
      success: false,
      responseCode: null,
      error:
        `${ERROR_CATEGORY.CONFIG_MISSING}: ` +
        'MS_SENDER_EMAILが未設定または無効です',
    };
  }

  let accessToken =
    getMicrosoftAccessToken();

  if (!accessToken) {
    return {
      success: false,
      responseCode: null,
      error:
        ERROR_CATEGORY.OUTLOOK_TOKEN_FAILED,
    };
  }

  const endpoint =
    GRAPH_SEND_MAIL_PATH.replace(
      '{senderEmail}',
      encodeURIComponent(senderEmail)
    );

  const mailPayload = {
    message: {
      subject: String(subject || ''),
      body: {
        contentType: 'HTML',
        content: String(htmlBody || ''),
      },
      toRecipients: [
        {
          emailAddress: {
            address: recipientEmail,
            name: String(
              recipient.name || ''
            ),
          },
        },
      ],
    },
    saveToSentItems: true,
  };

  const retryConfig = {
    maxAttempts: getNumSetting(
      'retry_max_attempts',
      3
    ),
    waitSeconds: getNumSetting(
      'retry_wait_seconds',
      3
    ),
  };

  let sendResult = executeGraphSendMail(
    endpoint,
    accessToken,
    mailPayload,
    retryConfig
  );

  /*
   * キャッシュ済みtokenが失効している場合だけ、
   * キャッシュを削除して新しいtokenで1回再実行する。
   */
  if (
    sendResult.response &&
    sendResult.response.getResponseCode() === 401
  ) {
    clearMicrosoftAccessTokenCache();

    accessToken =
      getMicrosoftAccessToken();

    if (accessToken) {
      sendResult = executeGraphSendMail(
        endpoint,
        accessToken,
        mailPayload,
        retryConfig
      );
    }
  }

  const response = sendResult.response;
  const attempts = sendResult.attempts;
  const error = sendResult.error;

  if (!response) {
    return {
      success: false,
      responseCode: null,
      error:
        error ||
        ERROR_CATEGORY.OUTLOOK_SEND_FAILED,
    };
  }

  const responseCode =
    response.getResponseCode();

  Logger.log(
    `Outlook送信: ${recipientEmail} ` +
    `HTTP ${responseCode} ` +
    `試行${attempts}回`
  );

  if (
    responseCode >= 200 &&
    responseCode < 300
  ) {
    return {
      success: true,
      responseCode,
      error: null,
    };
  }

  const errorCategory =
    classifyOutlookSendError(
      responseCode
    );

  if (responseCode === 401) {
    clearMicrosoftAccessTokenCache();
  }

  return {
    success: false,
    responseCode,
    error:
      `${errorCategory}: ` +
      `HTTP ${responseCode} ` +
      `(試行${attempts}回)`,
  };
}

/**
 * Graph sendMailリクエストを実行する。
 *
 * @param {string} endpoint
 * @param {string} accessToken
 * @param {Object} mailPayload
 * @param {{
 *   maxAttempts: number,
 *   waitSeconds: number
 * }} retryConfig
 * @returns {{
 *   response: GoogleAppsScript.URL_Fetch.HTTPResponse|null,
 *   attempts: number,
 *   error: string|null
 * }}
 */
function executeGraphSendMail(
  endpoint,
  accessToken,
  mailPayload,
  retryConfig
) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization:
        `Bearer ${accessToken}`,
    },
    payload: JSON.stringify(mailPayload),
    muteHttpExceptions: true,
  };

  return fetchWithRetry(
    endpoint,
    options,
    retryConfig
  );
}

/**
 * Graph送信失敗のHTTPコードを分類する。
 *
 * @param {number} responseCode
 * @returns {string}
 */
function classifyOutlookSendError(
  responseCode
) {
  if (
    responseCode === 401 ||
    responseCode === 403
  ) {
    return (
      ERROR_CATEGORY
        .OUTLOOK_PERMISSION_DENIED
    );
  }

  if (responseCode === 429) {
    return (
      ERROR_CATEGORY
        .OUTLOOK_RATE_LIMITED
    );
  }

  return ERROR_CATEGORY.OUTLOOK_SEND_FAILED;
}