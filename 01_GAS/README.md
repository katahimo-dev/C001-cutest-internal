# 01_GAS

Google Apps Script (GAS) 群を管理するモノレポ。
保育業務向けの勤怠管理、日報作成、ルート計算、通知連携をプロジェクト単位で分離して運用する。

## プロジェクト一覧

- [gas-childcare-daily-report](gas-childcare-daily-report/README.md)
  - スタッフ出勤簿の一括作成と、ルート集計からの勤怠転記。
- [gas-childcare-report](gas-childcare-report/README.md)
  - 保育日報Webアプリ、事故報告、領収書登録、認証/監査ログ。
- [gas-integrated-system](gas-integrated-system/README.md)
  - スマホ向け勤怠修正Webアプリ。スタッフ/管理者の権限制御あり。
- [gas-project-3](gas-project-3/README.md)
  - 出勤簿テンプレート向け補助スクリプト（onEdit/onOpen）。
- [gas-root-serach](gas-root-serach/README.md)
  - カレンダー予定解析、移動ルート計算、LineWorks通知、勤怠集計出力。

## リポジトリ構成

- `CHANGELOG.md`
  - 全体の統合更新履歴（Ver. 1.0.2 以降の横断変更）。
- `pull_all_gas.ps1`
  - 複数GASプロジェクトの同期用スクリプト。
- `Kokyaku_*.csv`
  - 顧客データのサンプル/入出力用CSV。
- `rerelase/`
  - リリース関連ドキュメント。
- `分析/`
  - 分析・補助資料。

## 運用方針

- 変更履歴
  - 横断変更はルート `CHANGELOG.md` に集約。
  - サブプロジェクト側は重複記載を避ける。
- 機密情報
  - APIキー、秘密鍵、トークン、ID類は Script Properties または安全な保管先で管理。
  - ソースへ直書きしない。
- デプロイ
  - 各 `gas-*` ディレクトリ単位で管理・デプロイ。
  - `appsscript.json` のスコープ差分に注意。

## 最低セットアップ

1. Googleアカウントで各対象スプレッドシート/Driveフォルダにアクセスできる状態を準備。
2. 各プロジェクトで必要な Script Properties を設定。
3. 時間主導トリガーを用途ごとに設定（例: 日次転記、通知、ログ書き出し）。
4. テスト用データで手動実行し、対象シートへの反映を確認。

## 補足

- 個別の設定値、必須シート、実行関数は各プロジェクトREADMEを参照。
- 仕様詳細は各 `specification.md` / `SECURITY_IMPLEMENTATION.md` を参照。
