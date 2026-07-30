# 災害時安否確認フォーム自動送信・集計システム

Google Apps Script (GAS) で動作する、地震発生時の安否確認メール自動送信・回答集計システムです。
外部サーバー不要。Google スプレッドシート + Google フォーム + GAS のみで完結します。

---

## 概要

気象庁などの地震情報 API を定期監視し、設定震度以上の地震を検知すると社員全員へ安否確認フォームのリンクをメールで自動送信します。フォーム回答はリアルタイムに集計され、管理者はスプレッドシートで状況を確認できます。

```
地震情報API ──(ポーリング)──▶ GAS ──▶ 社員へメール送信 (Gmail / Outlook)
                                          │
                              Googleフォーム◀── 社員が回答
                                          │
                              スプレッドシートへ自動集計
```

---

## 主な機能

| 機能 | 説明 |
|------|------|
| 地震自動検知 | 設定間隔（デフォルト10分）ごとに地震情報を取得し、閾値（デフォルト震度5弱）以上で発火 |
| 重複送信防止 | `LockService` + `event_id` による二重実行・二重送信を完全防止 |
| Gmail / Outlook 対応 | `send_mode` の切替で Gmail (MailApp) または Microsoft Graph API (Outlook) を使い分け |
| Googleフォーム連携 | 社員ごとの事前入力URL を生成。`event_id` / `employee_id` を自動埋め込み |
| 緊急通知 | 「本人にけがあり」「複数の被害あり」などの回答を検知し、管理者へ即時アラート |
| 失敗再送 | 送信失敗レコードを `pending` へリセットして再送できる管理者関数 |
| 実行時間対応 | GAS の6分制限に対応。超過時は残りを次回トリガーへ引き継ぎ |
| テストモード | `test_mode=TRUE` 時は `test_recipient_email` の1アドレスのみへ送信 |

---

## ファイル構成

```
anpi-system/
├── appsscript.json        # GASプロジェクト設定・OAuth スコープ
├── Config.gs              # 定数・設定値キャッシュ
├── Utils.gs               # 汎用ユーティリティ（日付・メール検証など）
├── SheetRepository.gs     # スプレッドシートへのCRUD操作
├── Earthquake.gs          # 地震情報取得アダプタ・送信判定
├── Employee.gs            # 社員情報取得・検証
├── FormUrl.gs             # フォーム事前入力URL生成
├── GmailSender.gs         # Gmail (MailApp) 送信
├── OutlookSender.gs       # Microsoft Graph API 送信・トークン管理
├── Notification.gs        # 通知レコード作成・送信処理・メール本文生成
├── ResponseHandler.gs     # フォーム回答処理・緊急判定
├── Summary.gs             # 回答集計・summaryシート書き込み
├── TriggerManager.gs      # Apps Script トリガー管理
├── Validation.gs          # バリデーション関数
├── Admin.gs               # 管理者用操作関数・スプレッドシート初期化
├── Main.gs                # エントリポイント
└── Tests.gs               # テスト関数（16件）
```

---

## スプレッドシートのシート構成

| シート名 | 役割 |
|----------|------|
| `settings` | システム設定（key/value 形式） |
| `employees` | 社員マスタ（employee_id / name / email / active / group） |
| `earthquake_events` | 検知した地震イベントのログ |
| `notification_status` | 社員ごとの送信・回答ステータス |
| `send_errors` | 送信エラーの詳細ログ |
| `form_responses` | Googleフォームの回答（自動連携） |
| `summary` | イベントごとの集計サマリー |

---

## 設定キー一覧（`settings` シート）

| キー | デフォルト | 説明 |
|------|-----------|------|
| `enabled` | FALSE | システム有効化フラグ（TRUE で動作開始） |
| `test_mode` | TRUE | テストモード（TRUE = test_recipient_email のみへ送信） |
| `send_mode` | gmail | 送信方式: `gmail` または `outlook` |
| `threshold` | 5- | 通知する最低震度（例: `4`, `5-`, `5+`, `6-`, `6+`, `7`） |
| `max_event_age_minutes` | 30 | 地震発生から最大何分以内のイベントを対象にするか |
| `check_interval_minutes` | 10 | 地震確認トリガーの間隔（分） |
| `admin_email` | （空） | 管理者通知の送信先メールアドレス |
| `test_recipient_email` | （空） | テスト送信先メールアドレス |
| `form_base_url` | （空） | Google フォームの viewform URL |
| `form_event_entry_id` | （空） | フォームの event_id フィールドの entry ID |
| `form_employee_id_entry_id` | （空） | フォームの employee_id フィールドの entry ID |
| `form_name_entry_id` | （空） | フォームの name フィールドの entry ID |
| `form_id` | （空） | Google フォームの ID（onFormSubmit トリガー用） |
| `sender_display_name` | 災害時安否確認 | メール送信者の表示名 |
| `outlook_sender_email` | （空） | Outlook 送信元アドレス（Script Properties が優先） |
| `retry_max_attempts` | 3 | Outlook 送信の最大再試行回数 |
| `retry_wait_seconds` | 3 | Outlook 再試行待機（秒） |
| `immediate_admin_alert` | TRUE | 緊急回答時の管理者即時通知 |
| `summary_enabled` | TRUE | summary シートの自動更新 |
| `earthquake_info_url` | （空） | 地震情報取得 URL（アダプタ実装時に使用） |

---

## 主要な処理フロー

### 地震検知 → 送信

```
checkEarthquakeAndNotify()  [Main.gs / 定期トリガー]
  │
  ├─ fetchLatestEarthquake()     [Earthquake.gs] ← 要アダプタ実装
  ├─ shouldNotifyEarthquake()    [Earthquake.gs] 閾値・重複チェック
  │
  └─ processEarthquakeEvent()   [Main.gs]
       ├─ getActiveEmployees()        [Employee.gs]
       ├─ createNotificationRecords() [Notification.gs]
       └─ processPendingNotifications() [Notification.gs]
            ├─ buildPrefilledFormUrl()   [FormUrl.gs]
            ├─ sendMailGmail() / sendMailOutlook()
            └─ rebuildSummary()          [Summary.gs]
```

### フォーム回答受信

```
handleFormSubmit()  [ResponseHandler.gs / onFormSubmit トリガー]
  ├─ updateNotification() → status: responded
  ├─ isEmergencyResponse() → 緊急時は notifyAdmin()
  └─ rebuildSummary()
```

---

## セキュリティ設計

- **Secret 情報は Script Properties のみ** に保存（シートへは書かない）
  `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_SENDER_EMAIL`
- アクセストークンはキャッシュに保存するが、**client secret はキャッシュしない**
- ログへ secret / token を出力しない
- Outlook 送信時の 401 エラー後は**トークンキャッシュを自動クリア**

---

## 地震情報アダプタについて

`Earthquake.gs` の `fetchLatestEarthquake()` は**スタブ（未実装）** です。
利用する地震情報 API に合わせて実装してください。

→ 詳細は [SETUP_GUIDE.md](./SETUP_GUIDE.md) の「地震情報アダプタの実装」を参照

---

## ライセンス
