# セットアップ・利用ガイド

このドキュメントでは、**災害時安否確認システム** をゼロから動作させるまでの手順を説明します。

---

## 目次

1. [前提条件](#前提条件)
2. [ステップ1：Google スプレッドシートと GAS プロジェクトの作成](#ステップ1)
3. [ステップ2：GAS ファイルのコピー](#ステップ2)
4. [ステップ3：スプレッドシートの初期化](#ステップ3)
5. [ステップ4：Google フォームの作成と連携](#ステップ4)
6. [ステップ5：settings シートの設定](#ステップ5)
7. [ステップ6：地震情報アダプタの実装](#ステップ6)
8. [ステップ7：メール送信方式の設定](#ステップ7)
9. [ステップ8：トリガーのインストール](#ステップ8)
10. [ステップ9：テスト送信の実行](#ステップ9)
11. [ステップ10：本番稼働](#ステップ10)
12. [日常運用・管理者操作](#日常運用管理者操作)
13. [トラブルシューティング](#トラブルシューティング)

---

## 前提条件

- Google アカウントがあること
- (Outlook モード使用時) Microsoft Azure AD テナント・アプリ登録が済んでいること
- 地震情報を提供する API または RSS のエンドポイントが利用可能であること

---

## ステップ1：スプレッドシートと GAS プロジェクトの作成 {#ステップ1}

1. Google ドライブで新しいスプレッドシートを作成する
2. スプレッドシートのメニューから **[拡張機能] → [Apps Script]** を開く
3. プロジェクト名を `anpi-system` などわかりやすい名前に変更する

---

## ステップ2：GAS ファイルのコピー {#ステップ2}

GAS エディタで、既存の `コード.gs` を削除してから以下の順に新規ファイルを作成し、各ファイルのコードを貼り付けます。

| ファイル名 | 内容 |
|-----------|------|
| `Config.gs` | 定数・設定値キャッシュ |
| `Utils.gs` | 汎用ユーティリティ |
| `SheetRepository.gs` | シート操作 |
| `Earthquake.gs` | 地震情報アダプタ（後で実装） |
| `Employee.gs` | 社員情報取得 |
| `FormUrl.gs` | フォーム URL 生成 |
| `GmailSender.gs` | Gmail 送信 |
| `OutlookSender.gs` | Outlook 送信 |
| `Notification.gs` | 通知処理 |
| `ResponseHandler.gs` | フォーム回答処理 |
| `Summary.gs` | 集計処理 |
| `TriggerManager.gs` | トリガー管理 |
| `Validation.gs` | バリデーション |
| `Admin.gs` | 管理者関数 |
| `Main.gs` | エントリポイント |

`appsscript.json` は GAS エディタの歯車アイコンから「appsscript.json マニフェストファイルをエディタで表示する」を有効にして貼り付けます。

---

## ステップ3：スプレッドシートの初期化 {#ステップ3}

GAS エディタで `setupSpreadsheet` を選択して実行します。

```
Admin.gs → 関数: setupSpreadsheet
```

初回のみ「権限を確認」ダイアログが表示されます。「詳細」→「プロジェクト名へ移動（安全でない）」をクリックして承認してください。

実行後、スプレッドシートに以下のシートが自動作成されます：

- `settings` / `employees` / `earthquake_events`
- `notification_status` / `send_errors` / `form_responses` / `summary`

---

## ステップ4：Google フォームの作成と連携 {#ステップ4}

### 4-1. フォームを作成する

Google フォームで以下の質問を追加します（フィールド名は任意だが、後の設定で列名を合わせること）。

| 質問タイプ | 質問文（例） | 備考 |
|-----------|-------------|------|
| 短答式（非表示推奨） | event_id | 事前入力で自動設定 |
| 短答式（非表示推奨） | employee_id | 事前入力で自動設定 |
| 短答式（非表示推奨） | name | 事前入力で自動設定 |
| プルダウン | 安否状況 | 「被害なし」「本人にけがあり」など |
| プルダウン | 出社可否 | 「出社可能」「出社不可（災害）」など |
| 段落 | コメント (任意) | 自由記述 |

### 4-2. entry ID を取得する

フォームのプレビュー画面を開き、ブラウザの開発者ツールで各フィールドの `entry.XXXXXXXXX` 形式の ID を確認します。または URL に手動で質問を入力した状態の `?entry.XXX=test` を確認します。

### 4-3. フォームの回答先スプレッドシートを連携する

フォームの「回答」タブ → スプレッドシートアイコン → **既存のスプレッドシートを選択** → 先ほど作成したスプレッドシートの `form_responses` シートへ連携します。

> [!IMPORTANT]
> フォーム回答のシート名が `form_responses` になっていることを確認してください。自動生成される場合は名前を変更してください。

---

## ステップ5：settings シートの設定 {#ステップ5}

スプレッドシートの `settings` シートを開き、`value` 列を埋めます。

```
enabled                 → FALSE（最初はFALSEのまま）
test_mode               → TRUE
send_mode               → gmail（または outlook）
threshold               → 5-
admin_email             → your-admin@example.com
test_recipient_email    → your-test@example.com
form_base_url           → https://docs.google.com/forms/d/FORM_ID/viewform
form_event_entry_id     → entry.123456789（フォームの entry ID）
form_employee_id_entry_id → entry.987654321
form_name_entry_id      → entry.111222333
form_id                 → FORM_ID（URLから取得）
```

> [!NOTE]
> `enabled` は本番稼働を開始するまで `FALSE` にしておいてください。

---

## ステップ6：地震情報アダプタの実装 {#ステップ6}

`Earthquake.gs` の `fetchLatestEarthquake()` 関数はスタブになっています。
利用する API に合わせて実装してください。

### 実装例（気象庁 JSON API を使用する場合）

```javascript
function fetchLatestEarthquake() {
  const url = getSetting('earthquake_info_url', '');
  if (!url) {
    Logger.log('earthquake_info_url が未設定です');
    return null;
  }

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_FETCH_FAILED}: HTTP ${response.getResponseCode()}`);
      return null;
    }
    const json = JSON.parse(response.getContentText());
    return normalizeEarthquakeData(json);
  } catch (err) {
    Logger.log(`${ERROR_CATEGORY.EARTHQUAKE_FETCH_FAILED}: ${safeErrorMessage(err)}`);
    return null;
  }
}
```

次に `normalizeEarthquakeData(raw)` 関数内のプレースホルダー (`YOUR_EVENT_ID_FIELD` など) を実際の API フィールド名に書き換えてください。

### 返却すべき EarthquakeEvent 形式

```javascript
{
  eventId: "一意のイベントID（文字列）",
  occurredAt: new Date("発生日時"),
  announcedAt: new Date("発表日時") | null,
  hypocenter: "震源地名（文字列）",
  magnitude: 5.0 | null,
  maxIntensity: "5-",  // "1","2","3","4","5-","5+","6-","6+","7" のいずれか
  sourceUrl: "情報源URL"
}
```

---

## ステップ7：メール送信方式の設定 {#ステップ7}

### Gmail モード（推奨・シンプル）

`settings` シートで `send_mode = gmail` に設定するだけです。追加設定は不要です。
Gmail の1日の送信上限（通常 100〜500 通）に注意してください。

### Outlook モード（Microsoft Graph API）

1. **Azure AD でアプリ登録** を行い、`Mail.Send` のアプリケーション権限を付与します
2. GAS エディタの **[プロジェクト設定] → [スクリプト プロパティ]** に以下を追加します：

   | プロパティ名 | 値 |
   |-------------|-----|
   | `MS_TENANT_ID` | Azure AD のテナント ID |
   | `MS_CLIENT_ID` | アプリ登録のクライアント ID |
   | `MS_CLIENT_SECRET` | アプリ登録のクライアントシークレット |
   | `MS_SENDER_EMAIL` | 送信元メールアドレス |

3. `settings` シートで `send_mode = outlook` に設定します

> [!CAUTION]
> クライアントシークレットはスプレッドシートの `settings` シートやスクリプトコードに直書きしないでください。必ず Script Properties に保存してください。

---

## ステップ8：トリガーのインストール {#ステップ8}

GAS エディタで `installTriggersAdmin` を実行します。

```
Admin.gs → 関数: installTriggersAdmin
```

実行後、以下のトリガーが自動登録されます：

| トリガー | 関数 | タイミング |
|---------|------|-----------|
| 地震確認 | `checkEarthquakeAndNotify` | `check_interval_minutes` 分ごと（デフォルト10分） |
| フォーム回答 | `handleFormSubmit` | `form_id` 設定時のみ自動登録 |

> [!NOTE]
> `form_id` が未設定の場合は、フォーム回答トリガーを GAS エディタの「トリガー」メニューから手動で設定してください。

---

## ステップ9：テスト送信の実行 {#ステップ9}

1. `settings` で `test_mode = TRUE`, `test_recipient_email = your@email.com` を確認
2. GAS エディタで `sendTestNotification` を実行

```
Admin.gs → 関数: sendTestNotification
```

指定した `test_recipient_email` に安否確認メールが届けば成功です。

### テスト時のメール例

```
件名: 【テスト】【安否確認 TEST-202406011030】災害時安否確認のお願い

テスト太郎 さん

お疲れ様です。安否確認担当者です。
先ほど地震が発生しました。...

[安否確認フォームへ回答する]  ← ボタンリンク
```

---

## ステップ10：本番稼働 {#ステップ10}

1. `employees` シートに社員情報を入力する

   | employee_id | name | email | active | group |
   |------------|------|-------|--------|-------|
   | EMP001 | 山田 太郎 | taro@example.com | TRUE | 本社 |
   | EMP002 | 鈴木 花子 | hanako@example.com | TRUE | 大阪 |

2. `settings` シートで `test_mode = FALSE` に変更する
3. `settings` シートで `enabled = TRUE` に変更する

以上で本番稼働開始です。次のトリガー実行タイミングから地震監視が開始されます。

---

## 日常運用・管理者操作

### 主な管理者関数

| 関数名 | 説明 | 場所 |
|--------|------|------|
| `runManualCheck()` | 手動で地震確認を実行 | Admin.gs |
| `sendTestNotification()` | テストメールを1通送信 | Admin.gs |
| `resendFailedNotifications(eventId)` | 失敗した通知を再送 | Admin.gs |
| `rebuildCurrentSummary()` | 最新イベントのサマリーを再集計 | Admin.gs |
| `resetStaleSendingNotifications(eventId, minutes)` | 送信中のまま止まった通知をリセット | Admin.gs |
| `disableSystem()` | システムを即時無効化 | Admin.gs |
| `setupSpreadsheet()` | シートの再初期化（データは消えない） | Admin.gs |
| `installTriggersAdmin()` | トリガーを再インストール | Admin.gs |
| `removeManagedTriggersAdmin()` | システムのトリガーを全削除 | Admin.gs |

### 失敗通知の再送手順

1. `earthquake_events` シートで対象の `event_id` を確認する
2. GAS エディタで `resendFailedNotifications` を実行する前に、関数内の引数にイベントIDを設定

```javascript
// Admin.gs の resendFailedNotifications の呼び出し例
// スクリプトエディタで直接実行する場合は以下のラッパーを追加
function resendFailed_EQ20240601001() {
  resendFailedNotifications('EQ20240601001');
}
```

### summary シートで進捗を確認する

`summary` シートで以下の情報を確認できます：

| 項目 | 説明 |
|------|------|
| `target_employee_count` | 送信対象者数 |
| `sent_count` | 送信成功数 |
| `response_count` | 回答数 |
| `response_rate` | 回答率（%） |
| `no_damage_count` | 「被害なし」の回答数 |
| `personal_injury_count` | 「本人にけがあり」の回答数 |
| 下部の詳細一覧 | 未回答者・送信失敗者・緊急確認対象の一覧 |

---

## トラブルシューティング

### Q. メールが届かない

1. `notification_status` シートで該当社員のステータスを確認する
2. `send_errors` シートでエラー内容を確認する
3. スクリプトログ（GAS エディタ → [実行数]）を確認する

### Q. 地震が検知されない

- `settings.enabled` が `TRUE` になっているか確認
- `fetchLatestEarthquake()` アダプタが正しく実装されているか確認（`runManualCheck()` で手動テスト）
- トリガーが登録されているか GAS エディタの「トリガー」メニューで確認

### Q. Outlook で "CONFIG_MISSING" エラーが出る

- Script Properties に `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_SENDER_EMAIL` が設定されているか確認
- Azure AD でアプリに `Mail.Send` の**アプリケーション**権限（委任ではなく）が付与されているか確認
- 管理者同意が完了しているか確認

### Q. 同じ地震で2回メールが送信された

- `earthquake_events` シートの `status` が `completed` になっているイベントは再送されません
- `LockService` により同時実行も防いでいますが、異なる `event_id` が返ってくる場合は重複します（アダプタ側で一意な ID を返すよう実装する必要があります）

### Q. GAS の実行時間が6分を超える

送信対象者が多い場合、GAS は6分以内に終了できず `incomplete=true` として中断します。次回の定期トリガー（デフォルト10分後）に残りの送信が継続されます。状況は `notification_status` シートで確認できます。

---

## テスト関数の実行方法

`Tests.gs` に16件のテスト関数が含まれています。

```
Tests.gs → 関数: runAllTests（全テスト一括実行）
または
Tests.gs → 個別のテスト関数を選択して実行
```

各テストはスクリプトログに `PASS` / `FAIL` を出力します。
