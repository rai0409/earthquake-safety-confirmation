/**
 * Earthquake.gs
 * 気象庁地震情報の取得・正規化・送信判定
 *
 * 【重要】fetchLatestEarthquake() の取得先URLは確定していないため、
 * アダプタ関数として分離している。
 * README の「地震情報アダプタの設定方法」を参照して実装を置き換えること。
 */

/**
 * 震度文字列を比較用の数値ランクへ変換
 * @param {string} intensity - "1","2","3","4","5-","5+","6-","6+","7"
 * @returns {number|null}
 */
function intensityToRank(intensity) {
  const map = {
    '1': 10,
    '2': 20,
    '3': 30,
    '4': 40,
    '5-': 45,
    '5+': 50,
    '6-': 55,
    '6+': 60,
    '7': 70,
  };
  const key = String(intensity).trim();
  if (key in map) return map[key];
  return null;
}

/**
 * 2つの震度を比較する
 * @param {string} a
 * @param {string} b
 * @returns {number} 正: a>b, 0: a=b, 負: a<b
 */
function compareIntensity(a, b) {
  const ra = intensityToRank(a);
  const rb = intensityToRank(b);
  if (ra === null || rb === null) return null;
  return ra - rb;
}

/**
 * 内部共通形式
 * @typedef {Object} EarthquakeEvent
 * @property {string} eventId
 * @property {Date} occurredAt
 * @property {Date|null} announcedAt
 * @property {string} hypocenter
 * @property {number|null} magnitude
 * @property {string} maxIntensity
 * @property {string} sourceUrl
 */

/**
 * 気象庁地震情報を取得して内部共通形式へ変換するアダプタ
 *
 * !! アダプタ実装ノート !!
 * 気象庁が公開しているデータ形式は変更される可能性があるため、
 * このファイル内でのみ外部データ形式に依存する処理を記述する。
 * 実際のURLはsettingsシートの earthquake_info_url キーへ設定するか、
 * 下記の PLACEHOLDER を置き換えること。
 * 詳細は README の「地震情報アダプタの設定方法」を参照。
 *
 * @returns {EarthquakeEvent|null} 取得・解析に失敗した場合はnull
 */
function fetchLatestEarthquake() {
  // --- アダプタ実装開始 ---
  // settingsからURLを取得する例:
  // const url = getSetting('earthquake_info_url', '');
  //
  // ここに実際の取得・解析処理を記述する。
  // 必ず内部共通形式（EarthquakeEvent）へ変換して返すこと。
  // 取得・解析失敗時は null を返し、例外をスローしないこと。
  //
  // 実装例（気象庁JSONゆれんち形式を参考にした仮実装）:
  // try {
  //   const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  //   if (response.getResponseCode() !== 200) {
  //     Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_FETCH_FAILED}: HTTP ${response.getResponseCode()}`);
  //     return null;
  //   }
  //   const json = JSON.parse(response.getContentText());
  //   return normalizeEarthquakeData(json);
  // } catch (err) {
  //   Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_FETCH_FAILED}: ${safeErrorMessage(err)}`);
  //   return null;
  // }

  Logger.log('fetchLatestEarthquake: アダプタ未実装。README を参照して実装してください。');
  return null;
  // --- アダプタ実装終了 ---
}

/**
 * 外部データを内部共通形式へ変換
 * 実際のデータ形式に合わせてこの関数を変更する
 *
 * @param {Object} raw - 外部APIのレスポンスJSON
 * @returns {EarthquakeEvent|null}
 */
function normalizeEarthquakeData(raw) {
  try {
    if (!raw) return null;

    // !! 以下はアダプタ実装者が実際のデータ形式に合わせて書き換える !!
    // 下記はフィールド名のプレースホルダー。実際のキー名に置換すること。
    const eventId = String(raw['YOUR_EVENT_ID_FIELD'] || '').trim();
    const occurredAtStr = raw['YOUR_OCCURRED_AT_FIELD'];
    const announcedAtStr = raw['YOUR_ANNOUNCED_AT_FIELD'] || null;
    const hypocenter = String(raw['YOUR_HYPOCENTER_FIELD'] || '').trim();
    const magnitude = raw['YOUR_MAGNITUDE_FIELD'] != null
      ? Number(raw['YOUR_MAGNITUDE_FIELD'])
      : null;
    const maxIntensity = String(raw['YOUR_MAX_INTENSITY_FIELD'] || '').trim();
    const sourceUrl = String(raw['YOUR_SOURCE_URL_FIELD'] || '').trim();

    if (!eventId) {
      Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_PARSE_FAILED}: event_idが空`);
      return null;
    }

    const occurredAt = occurredAtStr ? new Date(occurredAtStr) : null;
    if (!occurredAt || isNaN(occurredAt.getTime())) {
      Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_PARSE_FAILED}: occurredAtが無効`);
      return null;
    }

    if (intensityToRank(maxIntensity) === null) {
      Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_PARSE_FAILED}: 未知の震度値 "${maxIntensity}"`);
      return null;
    }

    return {
      eventId,
      occurredAt,
      announcedAt: announcedAtStr ? new Date(announcedAtStr) : null,
      hypocenter,
      magnitude: magnitude !== null && !isNaN(magnitude) ? magnitude : null,
      maxIntensity,
      sourceUrl,
    };
  } catch (err) {
    Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_PARSE_FAILED}: ${safeErrorMessage(err)}`);
    return null;
  }
}

/**
 * 地震イベントが通知対象か判定する
 * 送信判定の全条件（仕様書 §8.4）を確認する
 *
 * @param {EarthquakeEvent} event
 * @returns {{ ok: boolean, reason: string }}
 */
function shouldNotifyEarthquake(event) {
  // 条件1: enabled が TRUE
  const enabled = getBoolSetting('enabled', false);
  if (!enabled) {
    return { ok: false, reason: 'システムが無効です (enabled=FALSE)' };
  }

  // 条件2: 地震情報が正常に取得できた（nullでない）
  if (!event) {
    return { ok: false, reason: `${ERROR_CATEGORY.EARTHQUAKE_FETCH_FAILED}` };
  }

  // 条件3: event_idが空でない
  if (!event.eventId) {
    return { ok: false, reason: `${ERROR_CATEGORY.EARTHQUAKE_PARSE_FAILED}: event_idが空` };
  }

  // 条件4: 最大震度がthreshold以上
  const threshold = getSetting('threshold', '5-');
  const eventRank = intensityToRank(event.maxIntensity);
  const thresholdRank = intensityToRank(String(threshold).trim());
  if (eventRank === null) {
    return { ok: false, reason: `${ERROR_CATEGORY.EARTHQUAKE_PARSE_FAILED}: 未知の震度値 "${event.maxIntensity}"` };
  }
  if (thresholdRank === null) {
    return { ok: false, reason: `${ERROR_CATEGORY.CONFIG_MISSING}: 不正なthreshold値 "${threshold}"` };
  }
  if (eventRank < thresholdRank) {
    return { ok: false, reason: `${ERROR_CATEGORY.BELOW_THRESHOLD}: 震度${event.maxIntensity}はthreshold(${threshold})未満` };
  }

  // 条件5: occurred_atが現在からmax_event_age_minutes以内
  const maxAgeMinutes = getNumSetting('max_event_age_minutes', 30);
  const nowMs = Date.now();
  const diffMs = nowMs - event.occurredAt.getTime();
  if (diffMs > maxAgeMinutes * 60 * 1000) {
    return { ok: false, reason: `${ERROR_CATEGORY.EVENT_TOO_OLD}: 発生から${Math.floor(diffMs / 60000)}分経過` };
  }
  if (diffMs < 0) {
    return { ok: false, reason: `${ERROR_CATEGORY.EVENT_TOO_OLD}: 発生時刻が未来` };
  }

  // 条件6: earthquake_eventsに同じevent_idの完了済み記録がない
  if (isEventAlreadyCompleted(event.eventId)) {
    return { ok: false, reason: `${ERROR_CATEGORY.DUPLICATE_EVENT}: event_id "${event.eventId}" は完了済み` };
  }

  return { ok: true, reason: '送信対象' };
}
