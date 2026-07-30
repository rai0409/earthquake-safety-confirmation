/**
 * TriggerManager.gs
 * Apps Script トリガーの管理
 */

// トリガーハンドラ関数名（ScriptAppと紐付けるための文字列）
const TRIGGER_HANDLER_EARTHQUAKE = 'checkEarthquakeAndNotify';
const TRIGGER_HANDLER_FORM = 'handleFormSubmit';
// トリガーを識別するためのPropertiesキー
const MANAGED_TRIGGER_IDS_KEY = 'managed_trigger_ids';

/**
 * 本システムが管理するトリガーIDをScriptPropertiesへ保存
 * @param {string} triggerId
 */
function saveManagedTriggerId(triggerId) {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(MANAGED_TRIGGER_IDS_KEY);
  const ids = existing ? JSON.parse(existing) : [];
  if (!ids.includes(triggerId)) {
    ids.push(triggerId);
    props.setProperty(MANAGED_TRIGGER_IDS_KEY, JSON.stringify(ids));
  }
}

/**
 * 管理対象トリガーIDのリストを取得
 * @returns {string[]}
 */
function getManagedTriggerIds() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(MANAGED_TRIGGER_IDS_KEY);
  return existing ? JSON.parse(existing) : [];
}

/**
 * 特定ハンドラ関数のトリガーが既に存在するか確認
 * @param {string} handlerFunction
 * @returns {boolean}
 */
function hasManagedTrigger(handlerFunction) {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.some(t => t.getHandlerFunction() === handlerFunction);
}

/**
 * トリガーのインストール
 * 重複トリガーを作らない
 * フォーム回答は回答先スプレッドシートのonFormSubmitで受信する
 */
function installTriggers() {
  const checkIntervalMinutes = getNumSetting('check_interval_minutes', 10);

  // 地震確認トリガー
  if (!hasManagedTrigger(TRIGGER_HANDLER_EARTHQUAKE)) {
    const trigger = ScriptApp.newTrigger(TRIGGER_HANDLER_EARTHQUAKE)
      .timeBased()
      .everyMinutes(checkIntervalMinutes)
      .create();
    saveManagedTriggerId(trigger.getUniqueId());
    Logger.log(`トリガー作成: ${TRIGGER_HANDLER_EARTHQUAKE} (${checkIntervalMinutes}分ごと)`);
  } else {
    Logger.log(`トリガー既存: ${TRIGGER_HANDLER_EARTHQUAKE}`);
  }

  // e.namedValuesを使用するため、
  // Googleフォーム本体ではなく回答先スプレッドシートへ設定する
  if (!hasManagedTrigger(TRIGGER_HANDLER_FORM)) {
    try {
      const spreadsheet = getSpreadsheet();
      const trigger = ScriptApp.newTrigger(TRIGGER_HANDLER_FORM)
        .forSpreadsheet(spreadsheet)
        .onFormSubmit()
        .create();
      saveManagedTriggerId(trigger.getUniqueId());
      Logger.log(`トリガー作成: ${TRIGGER_HANDLER_FORM}`);
    } catch (err) {
      Logger.log(`フォームトリガー作成失敗: ${safeErrorMessage(err)}`);
      Logger.log(
        'SPREADSHEET_IDとGoogleフォームの回答先設定を確認してください。'
      );
    }
  } else {
    Logger.log(`トリガー既存: ${TRIGGER_HANDLER_FORM}`);
  }
}

/**
 * 本システムが作成したトリガーだけを削除
 * 他スクリプトのトリガーは削除しない
 */
function removeManagedTriggers() {
  const managedIds = getManagedTriggerIds();
  const allTriggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;

  for (const trigger of allTriggers) {
    const id = trigger.getUniqueId();
    const handlerFn = trigger.getHandlerFunction();
    // 管理対象IDに含まれるか、本システムのハンドラ関数のトリガーを削除
    if (managedIds.includes(id) ||
        handlerFn === TRIGGER_HANDLER_EARTHQUAKE ||
        handlerFn === TRIGGER_HANDLER_FORM) {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
      Logger.log(`トリガー削除: ${handlerFn} (${id})`);
    }
  }

  // 管理リストをクリア
  PropertiesService.getScriptProperties().deleteProperty(MANAGED_TRIGGER_IDS_KEY);
  Logger.log(`${removedCount}件のトリガーを削除しました`);
}
