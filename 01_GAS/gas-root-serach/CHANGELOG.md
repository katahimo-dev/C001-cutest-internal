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


## [Ver. 1.0.2] - 2026-04-16

### ご要望（2026-04-16 追加）
- カレンダーから取得した訪問予定をスタッフの出勤簿ファイルへ直接反映したい。
- ルート情報（移動時間・移動距離・出勤/退勤距離）も出勤簿へ含めたい。
- 事務作業の内容説明を選択肢ではなく、カレンダーから自動取得した任意の文字列に対応したい。

### 追加機能（2026-04-16）
- 出勤簿行データ生成関数を追加。
  - `buildTimesheetRowDataFromAppointments_()`: appointments 配列から出勤簿フォーム用 rowData を生成
- 勤務時間計算ヘルパーを追加。
  - `calcDurationMinForAttendance_()`: 訪問開始〜終了時刻から勤務時間（分）を計算
- 事務作業判定ヘルパーを追加。
  - `isOfficeWorkAppointment_()`: 15分イベントを事務作業として判定

### 改善・変更（2026-04-16）
- `refreshAttendanceForStaffOnDate()` を拡張し、ライブラリの戻り値として出勤簿行データ（rowData）と appointments を返すように変更。
  - 呼び出し側（Web アプリ）でカレンダー取得と出勤簿ファイルの直接更新が容易に。

### 技術的な修正（2026-04-16）
- `refreshAttendanceForStaffOnDate()` にて、`buildTimesheetRowDataFromAppointments_()` を呼び出し、visits と officeWorks を分類・マッピング。
  - 訪問先1・2・3 の情報（顧客名、時刻、距離）を OUTPUT_COLUMNS 準拠の行データに集約。
- ルート計算結果（移動時間、移動距離）を rowData の該当列（H, Q, AG, AH, AI, AJ, etc.）に含める。
- 事務作業判定ロジック内で 15分イベント以外の短時間予定（例：時間未定）も考慮可能な拡張性を確保。

### 不具合修正

#### 概要
- イベント招待者が空の場合、該当スタッフに予定が割り当てられないという問題を修正しました。
- 「事務 → 通常訪問」のように事務の後に予定が連続するときに、事務の位置（開始位置）が空のため事務直後の訪問で移動距離が計算されない問題を修正しました。これにより、事務を経由しても正確な移動距離が反映されるようになります。

シナリオ	動作
訪問A → 事務（location有り）	A→事務の移動距離が計算される
訪問A → 事務（location無し）	事務の移動距離は空（location無いため）
訪問A → 事務（location無し） → 訪問B	事務を飛ばして、A→B の移動距離を計算する

#### 技術的な修正内容
- **[イベント] タグ招待者問題**
  - `getCalendarEvents` の EVENT オブジェクトに `ownerName` を追加し、カレンダーの所有者情報を保持するようにしました。
  - `groupEventsByStaff` にて `guestNames` が空のときは `ownerName` を用いて該当スタッフとして割当する処理を追加しました。

- **事務経由での距離計算問題**
  - `getCalendarEvents` の事務（OFFICE WORK）イベント生成時、緯度経度子空にするのではなく、location があれば緯度経度を抽出するよう変更しました。
  - `calculateDetailedRoutes` を改修し、場所なし予定（事務など）の場合、時系列で直前を遡って場所あり予定を探し、そこを移動元として距離を計算する処理を追加しました。

### 作業時間 4.0h
