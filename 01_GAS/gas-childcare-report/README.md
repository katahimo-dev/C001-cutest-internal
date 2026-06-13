# gas-childcare-report

保育日報作成、事故報告、領収書登録を行うスマホ向けGAS Webアプリ。
認証・監査ログ・通知機能を含む。

## 目的

- 日報入力とレポート生成の効率化。
- 現場向けの事故報告フロー整備。
- 領収書登録と通知の一元化。

## 主な機能

- `コード.js`
  - Webアプリ本体ロジック。
  - 顧客データ取得（CacheService活用）。
  - 日報/事故報告関連の処理。
  - 認証トークンベースのセッション制御。
  - パスワード変更・リセット。
  - 監査ログのバッファリングと定期保存。
- `index.html`
  - ログイン、日報入力、各種モーダルUI。
- `LineWorks.js`
  - LineWorks Bot API通知。
  - 通知先IDの管理。
  - `TEST_MODE` による通知停止。
- `historicalNames.js`
  - PoC匿名化で利用する名前辞書。

## セキュリティ要点

- パスワードは Salt付きハッシュ運用。
- メールアドレス認証 + トークンセッション。
- セッション有効期限管理（ローリング更新）。
- 操作ログをCSVで定期保存。
- 一部データを `sessionStorage` 管理に移行。

## 必須設定

- `appsscript.json` のWebアプリ設定・スコープ確認。
- Script Properties にLineWorks認証情報、Bot情報、通知先IDを登録。
- 必要なSpreadsheet ID / Folder ID の整合確認。

## 運用手順

1. Script Properties を設定。
2. Webアプリとしてデプロイ。
3. テストユーザーでログイン・保存・通知を検証。
4. 監査ログのトリガー（例: 10分間隔）を設定。

## 関連資料

- `SECURITY_IMPLEMENTATION.md`
- `CHANGELOG.md`
