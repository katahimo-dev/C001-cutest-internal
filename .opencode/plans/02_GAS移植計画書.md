---
title: "GASプロジェクト移植計画書"
subtitle: "TAYO-LINE(Lite版)への機能移行計画"
author: "かたひも 松浦 融"
date: "2026年2月7日"

**対象プロジェクト**: 保育業務GASツール群
**移行先**: TAYO-LINE (Lite版) Webシステム
**提案書参照**: 01_提案書_保育業務管理システム開発_準委任契約.md

---

# はじめに

本ドキュメントは、現行のGoogle Apps Script（GAS）で構築されている保育業務管理ツール群を、**TAYO-LINE (Lite版)** として本格的なWebシステムに移行するための詳細計画書です。

提案書（01_提案書_保育業務管理システム開発_準委任契約.md）で定義された**3階層アーキテクチャ**に沿って、各GASプロジェクトの機能をマッピングし、段階的な移行手順を定義します。

---

# 1. 現行GASプロジェクト構成

## 1-1. プロジェクト一覧

| プロジェクト名 | 主な機能 | コード量 | 移行優先度 |
|--------------|---------|---------|-----------|
| **gas-childcare-report** | AI日報生成・事故報告・認証・LineWorks連携 | ~1,300行 | **最高** |
| **gas-root-serach** | ルート検索・カレンダー連携・通知 | ~640行 | **高** |
| **gas-childcare-daily-report** | 出勤簿生成・スタッフ情報管理 | ~400行 | **高** |
| **gas-integrated-system** | スタッフ同期・統合管理 | ~150行 | **中** |
| **gas-project-3** | （要調査） | 不明 | **低** |

## 1-2. 機能別GASコード分布

```
【gas-childcare-report】★最大・最複雑
├─ Web Appフロントエンド (index.html) - 別ファイル
├─ バックエンドAPI (コード.js)
│  ├─ 認証システム (verifyLogin, checkSession)
│  ├─ AI日報生成 (generateReportWithWarnings, Gemini API)
│  ├─ 事故報告生成 (generateAccidentReport)
│  ├─ 領収書OCR (extractAmountFromImage)
│  ├─ 顧客データ管理 (getData, fetchDataFromSheet)
│  ├─ LineWorks通知 (sendToLineWorks)
│  └─ データ永続化 (saveReport, saveAccidentReport)
└─ 設定・定数 (DEFAULT_PROMPTS, SPREADSHEET_ID等)

【gas-root-serach】
├─ メイン処理 (main, runSpecifiedDate)
├─ カレンダー連携 (getCalendarEvents)
├─ ルート計算 (calculateDetailedRoutes, getRouteDetails)
├─ CSVデータ取込 (getCustomerDataFromCsv)
├─ スタッフデータ管理 (getStaffDataFromSpreadsheet)
└─ LineWorks連携 (sendDailyScheduleToLineWorks)

【gas-childcare-daily-report】
├─ 出勤簿生成 (attendance.js)
│  ├─ 一括ファイル作成 (createStaffAttendanceFiles)
│  ├─ テンプレート複製・設定 (setupAttendanceSheet)
│  └─ シート保護設定 (copyProtectionSettings)
└─ スタッフ情報同期 (staffinfomation.js, syncStaffInfo)

【gas-integrated-system】
└─ スタッフ同期 (STAFF_SYNC.js, syncStaffInfo)
```

---

# 2. TAYO-LINEアーキテクチャへのマッピング

## 2-1. 3階層構成との対応関係

提案書で定義された3階層構成に、現行GAS機能を以下のようにマッピングします：

```
【第1階層：共通基盤】（開発者権利）
├─ Auth Module ← gas-childcare-report/verifyLogin, checkSession
├─ 操作ログ・監査 ← gas-childcare-report/logToBuffer
├─ 通知基盤 ← gas-childcare-report/sendToLineWorks
├─ AI基盤 ← gas-childcare-report/callGemini
└─ 画面部品 ← gas-childcare-report/index.html (UI部品抽出)

【第2階層：業界共通基盤】（開発者権利）
├─ 顧客・児童情報管理 ← gas-childcare-report/getData
├─ スタッフ管理 ← gas-childcare-daily-report/staffinfomation.js
├─ スケジュール管理 ← gas-root-serach/getCalendarEvents
├─ 訪問ルート検索 ← gas-root-serach/calculateDetailedRoutes
├─ 日報・記録生成 ← gas-childcare-report/generateReportWithWarnings
└─ 勤怠管理 ← gas-childcare-daily-report/attendance.js

【第3階層：貴社独自機能】（キューテスト社帰属）
├─ Googleカレンダー連携 ← gas-root-serach/main.js
├─ RESERVA連携 ← gas-root-serach/getCustomerDataFromCsv
├─ LineWorks連携 ← gas-childcare-report/sendToLineWorks
├─ 出勤簿テンプレート形式 ← gas-childcare-daily-report/attendance.js
├─ 独自帳票フォーマット ← (要整理)
└─ 特殊業務ルール ← (要ヒアリング)
```

## 2-2. データベース設計との対応

| Firestoreコレクション | 対応GAS機能 | データソース |
|---------------------|-----------|------------|
| **/users** | 全GASのスタッフ管理 | STAFF_SS_IDスプレッドシート |
| **/customers** | gas-childcare-report/getData | SPREADSHEET_ID顧客DB |
| **/schedules** | gas-root-serach/getCalendarEvents | Google Calendar API |
| **/routes** | gas-root-serach/calculateDetailedRoutes | 計算結果 |
| **/attendance** | gas-childcare-daily-report/attendance.js | 生成データ |
| **/reports** | gas-childcare-report/saveReport | 日報データ |
| **/accidents** | gas-childcare-report/saveAccidentReport | 事故報告データ |
| **/receipts** | gas-childcare-report/processReceiptImages | 領収書画像 |

---

# 3. フェーズ別移行計画

## Phase 1: 認証基盤 + データ基盤（80〜120時間）

### 3-1. 移行対象機能

**【認証基盤】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-childcare-report/verifyLogin | Firebase Auth + カスタム検証 | 16h |
| gas-childcare-report/checkSession | Firebase Auth Persistence | 8h |
| gas-childcare-report/computeHash | Firebase Auth (自動) | 4h |
| gas-childcare-report/requestPasswordReset | Firebase Auth パスワードリセット | 8h |

**【データ基盤】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-childcare-report/getData | Firestore Customersコレクション | 20h |
| gas-childcare-report/fetchDataFromSheet | CSVインポート機能 | 12h |
| gas-childcare-daily-report/syncStaffInfo | Firestore Usersコレクション | 16h |
| gas-integrated-system/STAFF_SYNC.js | スタッフ同期バッチ | 8h |

### 3-2. 技術的留意点

**スプレッドシート連携の移行戦略：**
```javascript
// GAS側（現行）
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const data = sheet.getDataRange().getValues();

// TAYO-LINE側（移行後）
// 1. 初期移行時: CSVエクスポート → Firestoreインポート
// 2. 並行運用期間: GAS↔Firestore双方向同期
// 3. 完全移行後: GAS連携廃止、TAYO-LINEのみ使用
```

**認証情報の移行：**
- GAS: SHA-256 + Salt方式
- TAYO-LINE: Firebase Auth（推奨）または ハッシュ移行
- パスワードリセット必須の可能性あり

---

## Phase 2: 日報機能モジュール（60〜100時間）

### 3-3. 移行対象機能

**【AI日報生成】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-childcare-report/generateReportWithWarnings | Gemini API連携 | 20h |
| gas-childcare-report/DEFAULT_PROMPTS | プロンプト管理画面 | 12h |
| gas-childcare-report/callGemini | 共通AIサービス化 | 8h |

**【事故報告】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-childcare-report/generateAccidentReport | 事故報告生成API | 12h |
| gas-childcare-report/saveAccidentReport | Firestore保存 | 8h |

**【領収書管理】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-childcare-report/extractAmountFromImage | OCR機能 | 12h |
| gas-childcare-report/processReceiptImages | Firebase Storage連携 | 8h |

**【LineWorks通知】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-childcare-report/sendToLineWorks | 通知サービス化 | 8h |
| gas-childcare-report/saveReport内通知 | イベント駆動通知 | 4h |

### 3-4. UI/UX移行戦略

**GAS Web App → Reactコンポーネント：**
```
gas-childcare-report/index.html
├─ ログイン画面 → @/features/auth/LoginPage.tsx
├─ 日報入力画面 → @/features/report/DailyReportPage.tsx
├─ 事故報告画面 → @/features/report/AccidentReportPage.tsx
├─ 顧客選択モーダル → @/features/customer/CustomerSelector.tsx
├─ 領収書アップロード → @/features/report/ReceiptUploader.tsx
└─ 設定画面 → @/features/admin/SettingsPage.tsx
```

---

## Phase 3: 移動経路機能モジュール（40〜60時間）

### 3-5. 移行対象機能

**【カレンダー連携】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-root-serach/getCalendarEvents | Google Calendar API v3 | 12h |
| gas-root-serach/groupEventsByStaff | イベントグループ化 | 8h |
| タグ判定([予約確定]等) | 同ロジック継承 | 4h |

**【ルート計算】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-root-serach/calculateDetailedRoutes | Google Maps API連携 | 16h |
| gas-root-serach/getRouteDetails | ルート詳細計算 | 8h |
| gas-root-serach/resolveLocation | 期間限定住所対応 | 4h |

**【データ取込】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-root-serach/getCustomerDataFromCsv | CSVインポート(RESERVA対応) | 8h |
| gas-root-serach/getStaffDataFromSpreadsheet | スタッフデータ同期 | 8h |

### 3-6. スケジュール連携の改善点

**GAS版の課題：**
- タイムトリガー依存（毎日18-19時実行）
- 翌日予定のみ対応
- 計算結果のスプレッドシート保存

**TAYO-LINE版の改善：**
- Cloud Functionsスケジューラー
- リアルタイムカレンダー監視
- 複数日一括計算対応
- Firestore永続化 + 履歴管理

---

## Phase 4: 出勤簿機能モジュール（80〜100時間）

### 3-7. 移行対象機能

**【出勤簿生成】**
| GASソース | 移行先 | 工数見積 |
|----------|-------|---------|
| gas-childcare-daily-report/createStaffAttendanceFiles | 出勤簿自動生成 | 24h |
| gas-childcare-daily-report/setupAttendanceSheet | 月別シート生成 | 12h |
| gas-childcare-daily-report/copyProtectionSettings | 権限設定 | 8h |

**【編集・エクスポート】**
| 機能 | 実装内容 | 工数見積 |
|-----|---------|---------|
| Web編集機能 | 日別勤怠修正 | 16h |
| CSVエクスポート | 社労士提出用 | 8h |
| Excel出力 | 帳票印刷対応 | 12h |

### 3-8. データモデル変更

**GAS版（スプレッドシート中心）：**
```
テンプレートファイル複製 → スタッフ別ファイル生成
├─ 2025出勤簿/（フォルダ）
│  ├─ 山田太郎_出勤簿_2025年度
│  ├─ 佐藤花子_出勤簿_2025年度
│  └─ ...
```

**TAYO-LINE版（データベース中心）：**
```
Firestore: /attendance/{staffId}/{year}/{month}
├─ データ構造化
├─ リアルタイム更新
├─ 一括エクスポート
└─ 権限制御（Row Level Security）
```

---

# 4. 技術的移行課題と対策

## 4-1. 認証・認可

### 課題
- GAS: スプレッドシートベースの認証
- TAYO-LINE: Firebase Authへの移行

### 対策
```javascript
// 移行スクリプト例
// 1. GASからスタッフデータエクスポート
// 2. Firebase Authへ一括インポート
// 3. 初回ログイン時パスワードリセット促進
// 4. GAS側で退職者判定継続
```

## 4-2. データ移行

### 顧客データ
```
現行: Googleスプレッドシート（複数シート）
　　├─ 顧客基本情報
　　├─ 家族情報
　　└─ 住所・緯度経度

移行先: Firestore（正規化）
　　├─ /customers/{customerId}
　　├─ /families/{familyId}
　　└─ /addresses/{addressId}
```

### 履歴データ
- 日報履歴: GASスプレッドシート → Firestore移行
- 事故報告: 同上
- 領収書画像: Google Drive → Firebase Storage

## 4-3. 外部サービス連携

### Google Calendar API
- GAS: CalendarApp（組み込み）
- TAYO-LINE: Google Calendar API v3（OAuth2認証）

### Google Maps API
- GAS: Maps.newDirectionFinder（組み込み）
- TAYO-LINE: Directions API + Distance Matrix API

### Gemini API
- GAS: UrlFetchAppで直接呼び出し
- TAYO-LINE: Vertex AI or Gemini API（同一）

### LineWorks
- GAS: 直接HTTP呼び出し
- TAYO-LINE: 同ロジック継承（共通ライブラリ化）

## 4-4. セキュリティ考慮事項

### 匿名化モード（PoC）
```javascript
// GAS側実装を移植
gas-childcare-report/getPocConfig()
gas-childcare-report/anonymizeName()
gas-childcare-report/anonymizeAddr()

// TAYO-LINEでの実装
// 環境変数でPoCモード切替
// 画面表示時のリアルタイム匿名化
```

### 機密情報管理
- GAS: PropertiesService.getScriptProperties()
- TAYO-LINE: Secret Manager + 環境変数

---

# 5. 移行スケジュールとマイルストーン

## 5-1. フェーズ別マイルストーン

```
【Phase 1: 認証+データ基盤】
Week 1-2: 開発環境構築・Firebase設定
Week 3-4: 認証システム実装
Week 5-6: データインポート機能
Week 7-8: テスト・並行運用開始
★マイルストーン: ログイン機能リリース

【Phase 2: 日報機能】
Week 9-10: AI日報生成移植
Week 11-12: 事故報告・領収書機能
Week 13-14: LineWorks連携
Week 15-16: テスト・リリース
★マイルストーン: 日報機能リリース

【Phase 3: ルート検索】
Week 17-18: カレンダー連携
Week 19-20: ルート計算実装
Week 21-22: 通知機能
Week 23-24: テスト・リリース
★マイルストーン: 経路検索リリース

【Phase 4: 出勤簿】
Week 25-28: 出勤簿生成機能
Week 29-30: Web編集・エクスポート
Week 31-32: テスト・リリース
★マイルストーン: 出勤簿リリース
```

## 5-2. 並行運用期間

各フェーズリリース後、**2週間の並行運用期間**を設けます：
- GAS版: 読み取り専用（バックアップ）
- TAYO-LINE版: 新規データ入力
- データ同期: 両方向（検証用）

---

# 6. リスク管理

## 6-1. 技術的リスク

| リスク | 影響度 | 対策 |
|-------|-------|------|
| データ移行失敗 | 高 | 並行運用・段階的移行・バックアップ |
| 認証移行トラブル | 高 | パスワードリセットフロー準備 |
| AI API互換性 | 中 | Gemini API仕様確認・フォールバック準備 |
| Google API制限 | 中 | クォータ監視・キャッシュ戦略 |

## 6-2. 運用リスク

| リスク | 対策 |
|-------|------|
| ユーザートレーニング不足 | 移行マニュアル作成・研修実施 |
| 並行運用時の混乱 | 明確な運用ルール策定 |
| データ不整合 | 自動検証スクリプト開発 |

---

# 7. 移行後の保守・運用

## 7-1. GAS側の扱い

移行完了後のGASプロジェクト扱い：

```
【完全移行後（Phase 4完了後3ヶ月）】
gas-childcare-report: アーカイブ（読み取り専用）
gas-root-serach: アーカイブ
gas-childcare-daily-report: アーカイブ
gas-integrated-system: 停止（TAYO-LINEで代替）
```

## 7-2. データバックアップ戦略

- **Firestore**: 自動バックアップ（Cloud Backup）
- **Firebase Storage**: ライフサイクル管理
- **GAS**: 移行完了後はアーカイブ保持（1年間）

---

# 8. 付録

## 8-1. 移行対象コード詳細リスト

### gas-childcare-report/コード.js

**主要関数（移行必須）：**
- `doGet()` → React Router設定
- `getData()` → Firestoreクエリ
- `generateReportWithWarnings()` → AIサービス
- `saveReport()` → Firestore書き込み
- `verifyLogin()` → Firebase Auth
- `checkSession()` → Firebase Auth
- `sendToLineWorks()` → 通知サービス

**定数（移行必須）：**
- `SPREADSHEET_ID` → Firestore接続情報
- `DEFAULT_PROMPTS` → データベース化
- `PROMPT_KEYS` → 同

### gas-root-serach/main.js

**主要関数（移行必須）：**
- `main()` → Cloud Functionsスケジューラー
- `getCalendarEvents()` → Calendar APIサービス
- `calculateDetailedRoutes()` → ルート計算サービス
- `sendDailyScheduleToLineWorks()` → 通知サービス
- `getCustomerDataFromCsv()` → インポート機能

### gas-childcare-daily-report/attendance.js

**主要関数（移行必須）：**
- `promptAndCreateFiles()` → Web UI + API
- `createStaffAttendanceFiles()` → 出勤簿生成API
- `setupAttendanceSheet()` → データ生成ロジック

## 8-2. テスト計画概要

### 単体テスト
- GAS関数 → TypeScript関数への変換検証
- データ変換ロジックの一致確認

### 統合テスト
- エンドツーエンドフロー検証
- 外部API連携テスト

### 受入テスト
- 現場スタッフによる実運用テスト
- 並行運用期間でのデータ比較

## 8-3. トレーニング計画

1. **管理者向け**: システム管理・設定変更方法
2. **スタッフ向け**: 新UI操作・モバイル利用
3. **バックアップ担当**: データエクスポート方法

---

**本計画書は提案書に基づき、現行GASプロジェクトの実装詳細を反映して作成しました。**
**実際の開発においては、各フェーズ開始時に詳細設計を策定し、必要に応じて本計画を更新します。**

**作成日**: 2026年2月7日
**作成者**: かたひも 松浦 融
