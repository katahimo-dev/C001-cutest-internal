# 更新履歴 (Release Notes)

## [Ver. 1.0.1] - 2026-04-13

### ご要望
- ルート集計ファイルの更新タイミングでは当日の勤怠反映が間に合わないため、勤怠用の独立ファイルを運用したい。
- LineWorks通知は不要で、勤怠データ保存専用の自動処理を実行したい。

### 追加機能
- 勤怠転記専用ファイル「勤怠集計」の自動作成・更新機能を追加。
- 指定日の日次データを保存する saveAttendanceData を追加。
- トリガー実行用関数 autoRunSaveAttendance を追加（当日分を更新）。
- デバッグ関数を追加。
  - debugSaveAttendanceMarch2026
  - debugSaveAttendanceApril2026

### 改善・変更
- 勤怠集計出力時、同日データを先に削除してから再書き込みする方式に変更し、重複を防止。
- 勤怠保存処理では LineWorks 通知を行わない運用に分離。
- 共有ドライブ環境を考慮し、作成ファイルの配置処理を安定化。

### 技術的な修正
- CONFIG に ATTENDANCE_FILE_NAME / ATTENDANCE_FOLDER_ID を追加。
- outputToAttendanceFile と getOrCreateSpreadsheetInFolder を追加。
- saveAttendanceDataInRange を追加し、日付レンジ処理を共通化。
- 共通関数 buildDailyRouteRows を追加し、対象日・シート名生成、予定取得、グループ化、ルート計算を一元化。
- 既存挙動（予定なしになった場合の分岐、ルート集計何もしない　勤怠保存時はその行を削除、LineWorks通知有無の分離）は維持。

### 作業時間 2.5h
