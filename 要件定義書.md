# Google Apps Script 実装要件定義書
## スライドレポート自動生成システム（改訂版）

---

## 目次

1. [プロジェクト概要](#1-プロジェクト概要)
2. [システムアーキテクチャ](#2-システムアーキテクチャ)
3. [データ構造設計](#3-データ構造設計)
4. [スライドテンプレート設計](#4-スライドテンプレート設計)
5. [GAS関数設計](#5-gas関数設計)
6. [Gemini AI連携仕様](#6-gemini-ai連携仕様)
7. [UI/UX仕様](#7-uiux仕様)
8. [エラーハンドリング](#8-エラーハンドリング)
9. [セキュリティ](#9-セキュリティ)
10. [運用・保守](#10-運用保守)

---

## 1. プロジェクト概要

### 1.1 目的

CSVデータから美しいスライドレポートを自動生成し、Gemini AIによるインサイト生成で経営判断を支援するシステムを構築する。

**主要機能：**
- CSVデータのインポートと自動集計
- 4枚構成のモダンデザインスライド自動生成
- Gemini 2.5 Flash によるAIコメント・インサイト生成
- GUI設定画面による簡単セットアップ
- スクリプトプロパティによる設定管理

### 1.2 技術スタック

- **Google Apps Script** (JavaScript ES6)
- **Google Sheets API** (データ管理・集計)
- **Google Slides API** (スライド生成)
- **Google Drive API** (ファイル管理)
- **Gemini API** (AI機能)
- **HTML/CSS/JavaScript** (UI)

### 1.3 成果物

1. メインスクリプト（`Code.gs`）
2. HTML UI（5ファイル）
   - SettingsSidebar.html
   - AISidebar.html
   - MainSidebar.html
   - DummyDataDialog.html
   - HelpDialog.html
3. スライドテンプレート（自動生成）
4. 出力フォルダ（自動生成）

---

## 2. システムアーキテクチャ

### 2.1 全体構成

```
┌─────────────────────────┐
│  CSV Data Import        │
│  (User Upload)          │
└────────┬────────────────┘
         │
         ▼
┌─────────────────────────┐
│  Google Sheets          │
│  - RawSalesData         │
│  - RegionalSales (集計)  │
│  - PersonSales (集計)    │
│  - ProductSales (集計)   │
│  - CategorySales (集計)  │
│  - MonthlySales (集計)   │
└────────┬────────────────┘
         │
         │ データ取得
         ▼
┌─────────────────────────┐
│  Google Apps Script     │
│  - データ集計           │
│  - スライド生成         │
│  - Gemini AI連携        │
└────────┬────────────────┘
         │
         ├─────────────────┐
         │                 │
         ▼                 ▼
┌──────────────┐  ┌──────────────┐
│ Google Slides│  │  Gemini API  │
│ (レポート出力)│  │ (AI生成)     │
└──────────────┘  └──────────────┘
```

### 2.2 処理フロー

```
1. CSVインポート
   ↓
2. RawSalesDataシート作成
   ↓
3. QUERY関数で集計シート自動生成
   - RegionalSales
   - PersonSales
   - ProductSales
   - CategorySales
   - MonthlySales
   ↓
4. グラフ自動生成
   ↓
5. Gemini AIでコメント・インサイト生成（オプション）
   ↓
6. スライド生成・更新
   - テキスト置換
   - チャート挿入
   ↓
7. スクリプトプロパティ更新
```

---

## 3. データ構造設計

### 3.1 RawSalesData（生データ）

**カラム定義：**

| 列 | 名前 | 型 | 必須 | 説明 | 例 |
|----|------|-----|------|------|-----|
| A | Date | date | ✅ | 売上日 | 2025-01-01 |
| B | Region | string | ✅ | 地域 | 関東 |
| C | Person | string | ✅ | 担当者 | 田中太郎 |
| D | Product | string | ✅ | 製品 | 製品A |
| E | Category | string | ✅ | カテゴリ | サブスク |
| F | Quantity | number | ✅ | 数量 | 10 |
| G | UnitPrice | number | ✅ | 単価 | 50000 |
| H | TotalSales | number | ✅ | 売上額 | 500000 |

**データ例：**
```csv
Date,Region,Person,Product,Category,Quantity,UnitPrice,TotalSales
2025-01-01,関東,田中太郎,製品A,サブスク,10,50000,500000
2025-01-02,関西,鈴木花子,製品B,単発,5,100000,500000
2025-01-03,九州,佐藤一郎,製品C,保守,8,30000,240000
```

### 3.2 集計シート（自動生成）

#### RegionalSales

**QUERY式：**
```javascript
=QUERY(RawSalesData!B:H, "SELECT B, SUM(H) WHERE B IS NOT NULL GROUP BY B ORDER BY SUM(H) DESC LABEL B '地域', SUM(H) '売上'", 1)
```

**出力例：**
```
地域    | 売上
--------|----------
関東    | 5200000
関西    | 2800000
九州    | 1100000
```

#### PersonSales

**QUERY式：**
```javascript
=QUERY(RawSalesData!C:H, "SELECT C, SUM(H), COUNT(H) WHERE C IS NOT NULL GROUP BY C ORDER BY SUM(H) DESC LABEL C '担当者', SUM(H) '売上', COUNT(H) '件数'", 1)
```

**追加計算列：**
- D列（平均単価）= `IF(C2>0, B2/C2, 0)`

**出力例：**
```
担当者   | 売上    | 件数 | 平均単価
---------|---------|------|----------
田中太郎 | 3200000 | 15   | 213333
鈴木花子 | 2800000 | 12   | 233333
```

#### ProductSales

**QUERY式：**
```javascript
=QUERY(RawSalesData!D:H, "SELECT D, SUM(H), SUM(F) WHERE D IS NOT NULL GROUP BY D ORDER BY SUM(H) DESC LABEL D '製品', SUM(H) '売上', SUM(F) '販売数'", 1)
```

**追加計算列：**
- D列（平均単価）= `IF(C2>0, B2/C2, 0)`

#### CategorySales

**QUERY式：**
```javascript
=QUERY(RawSalesData!E:H, "SELECT E, SUM(H) WHERE E IS NOT NULL GROUP BY E ORDER BY SUM(H) DESC LABEL E 'カテゴリ', SUM(H) '売上'", 1)
```

#### MonthlySales

**QUERY式：**
```javascript
=QUERY(RawSalesData!A:H, "SELECT YEAR(A) & '-' & TEXT(MONTH(A), '00'), SUM(H) WHERE A IS NOT NULL GROUP BY YEAR(A), MONTH(A) ORDER BY YEAR(A), MONTH(A) LABEL YEAR(A) & '-' & TEXT(MONTH(A), '00') '年月', SUM(H) '売上'", 1)
```

**追加計算列：**
- C列（前月比）= `IF(B3>0, B3-B2, "-")`（3行目以降）
- D列（前月比率）= `IF(B2>0, TEXT((B3/B2-1), "0.0%"), "-")`（3行目以降）

### 3.3 スクリプトプロパティ

| キー | 説明 | 型 | 例 |
|------|------|-----|-----|
| SLIDE_TEMPLATE_ID | スライドテンプレートID | string | 自動設定 |
| CURRENT_SLIDE_ID | 最後に生成したスライドID | string | 自動更新 |
| OUTPUT_FOLDER_ID | 出力先フォルダID | string | 自動設定 |
| GEMINI_API_KEY | Gemini APIキー | string | 手動設定 |
| REPORT_TITLE | レポートタイトル | string | 月次売上レポート |
| PERIOD_TYPE | 期間タイプ | string | monthly |

---

## 4. スライドテンプレート設計

### 4.1 共通仕様

- **フォント：** Arial（日本語対応）
- **テキスト置換形式：** `{{変数名}}`
- **カラースキーム：**
  - プライマリ: `#667eea`（紫青）
  - セカンダリ: `#764ba2`（紫）
  - アクセント1: `#f093fb`（ピンク）
  - アクセント2: `#4facfe`（青）

### 4.2 Slide 1：表紙

**レイアウト：**
- 背景: グラデーション（`#667eea` → `#764ba2`）
- アクセント円: 左上・右下に配置
- タイトル: `{{reportTitle}}`（56pt、白、太字）
- サブタイトル: `{{period}}`（36pt、白）
- タイムスタンプ: `Generated at {{generatedAt}}`（12pt、右下）

### 4.3 Slide 2：サマリー

**要素：**
- ヘッダー帯: `#667eea`
- カード背景: 白（`#ffffff`）
- 内容:
  - 💰 合計売上: `{{totalSales}}`
  - 📈 成長率: `{{totalSalesChange}}`
  - 🏆 トップ地域: `{{topRegion}}` (`{{topRegionSales}}`)
  - 👤 トップ担当者: `{{topPerson}}` (`{{topPersonSales}}`)
  - 💡 AIコメント: `{{aiComment}}`

### 4.4 Slide 3：詳細分析

**要素：**
- ヘッダー帯: `#764ba2`
- 左: 地域別売上グラフ（棒グラフ）
- 右: 担当者別売上グラフ（横棒グラフ）
- チャート挿入位置:
  - 地域: (50, 170, 300, 250)
  - 担当者: (400, 170, 300, 250)

### 4.5 Slide 4：インサイト

**要素：**
- ヘッダー帯: `#f093fb`
- カード背景: 白
- AIインサイト: `{{aiInsight}}`（22pt）

---

## 5. GAS関数設計

### 5.1 セットアップ関数

```javascript
/**
 * 初期セットアップ（テンプレート＆フォルダ作成）
 * - 既存テンプレートがあれば削除して最新版に更新
 * - 出力フォルダは既存のものを使用
 */
function setupInitialEnvironment()

/**
 * スライドテンプレート作成
 * @param {Folder} folder 親フォルダ
 * @return {File} テンプレートファイル
 */
function createSlideTemplate_(folder)

/**
 * スライドテンプレート更新
 * - 既存テンプレートを削除
 * - 新しいテンプレートを作成
 * - スクリプトプロパティを更新
 */
function updateSlideTemplate()
```

### 5.2 データ管理関数

```javascript
/**
 * CSVインポート＆集計
 * @param {string} csvText CSV文字列
 * @return {Object} 結果オブジェクト
 */
function importAndAggregateSalesData(csvText)

/**
 * 集計シート作成（QUERY関数ベース）
 */
function createAggregationSheets_()

/**
 * 集計シート更新
 */
function refreshAggregationSheets()

/**
 * シートデータ取得
 * @param {string} sheetName シート名
 * @return {Array<Object>} データ配列
 */
function getSheetData_(sheetName)

/**
 * レポートデータ取得
 * @param {string} periodType 期間タイプ
 * @param {string} targetDate 対象日付
 * @return {Object} レポートデータ
 */
function getReportData_(periodType, targetDate)
```

### 5.3 スライド生成関数

```javascript
/**
 * スライド生成・更新
 * @param {Object} params パラメータ
 * @return {Object} 結果オブジェクト
 */
function generateOrUpdateSlide(params)

/**
 * 新規スライド作成
 * @param {Object} config 設定
 * @param {Object} data データ
 * @return {Presentation} プレゼンテーション
 */
function createNewSlide_(config, data)

/**
 * 既存スライド更新
 * @param {Presentation} presentation プレゼンテーション
 * @param {Object} data データ
 */
function updateSlide_(presentation, data)

/**
 * データ適用（テキスト置換＋チャート挿入）
 * @param {Presentation} presentation プレゼンテーション
 * @param {Object} data データ
 */
function applyDataToSlide_(presentation, data)

/**
 * チャート挿入
 * @param {Presentation} presentation プレゼンテーション
 */
function insertChartsFromSheet_(presentation)
```

### 5.4 設定管理関数

```javascript
/**
 * スクリプトプロパティ取得
 * @return {Object} 設定オブジェクト
 */
function getScriptProperties_()

/**
 * スクリプトプロパティ保存
 * @param {Object} config 設定オブジェクト
 */
function saveScriptProperties_(config)

/**
 * UI用設定取得（APIキーをマスク）
 * @return {Object} 設定オブジェクト
 */
function getConfigForUI()

/**
 * UI から設定保存
 * @param {Object} config 設定オブジェクト
 * @return {Object} 結果オブジェクト
 */
function saveConfigFromUI(config)
```

---

## 6. Gemini AI連携仕様

### 6.1 使用モデル

- **モデル名**: `gemini-2.0-flash-exp`
- **エンドポイント**: `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent`
- **認証**: APIキー方式

### 6.2 パラメータ設定

```javascript
generationConfig: {
  temperature: 0.7,    // 創造性（0.0-1.0）
  topK: 40,            // トークン選択数
  topP: 0.95,          // 累積確率
  maxOutputTokens: 1024 // 最大出力トークン
}
```

### 6.3 システムプロンプト

```
あなたは営業レポート分析の専門家です。以下のガイドラインに従ってください：
- 簡潔で具体的な分析を提供する
- 数値データに基づいた客観的な評価を行う
- ビジネスインサイトと実行可能な提案を含める
- ポジティブかつ建設的なトーンで記述する
```

### 6.4 主要関数

```javascript
/**
 * Gemini テキスト生成
 * @param {string} prompt ユーザープロンプト
 * @param {string} customPrompt カスタムシステムプロンプト
 * @return {Object} 結果オブジェクト
 */
function generateTextWithGemini(prompt, customPrompt = '')

/**
 * 売上サマリーコメント生成（150文字以内）
 * @return {Object} 結果オブジェクト
 */
function generateAIComment()

/**
 * ビジネスインサイト生成（200文字程度）
 * @return {Object} 結果オブジェクト
 */
function generateAIInsight()

/**
 * カスタムテキスト生成
 * @param {string} userPrompt ユーザープロンプト
 * @param {string} systemPrompt システムプロンプト
 * @return {Object} 結果オブジェクト
 */
function generateCustomText(userPrompt, systemPrompt = '')

/**
 * Gemini接続テスト
 * @return {Object} 結果オブジェクト
 */
function testGeminiConnection()
```

### 6.5 生成内容例

#### 売上サマリーコメント
```
1月の売上は14,500,000円（前月比+15.0%）を達成しました。
トップ地域の関東が5,200,000円、トップ担当者の田中太郎が3,200,000円と好調を維持しています。
```

#### ビジネスインサイト
```
【強み】
関東地域が全体の36%を占め、安定した売上基盤を確立しています。

【課題】
四国・中国地域の売上が低迷（合計8%）しており、テコ入れが必要です。

【提案】
1. 関東の成功事例を他地域に横展開
2. 低迷地域への営業リソース強化
3. トップ担当者のノウハウ共有会を実施
```

---

## 7. UI/UX仕様

### 7.1 メニュー構成

```
📊 スライドレポート
├── ⚙️ 設定
├── ────────────
├── 🎲 ダミーデータ生成
├── 📊 レポート生成
├── ────────────
└── ❓ ヘルプ
```

### 7.2 SettingsSidebar（設定画面）

**セクション構成：**

1. **初期セットアップ**
   - 「⚡ セットアップ実行」ボタン
   - 処理内容:
     - 既存テンプレート削除 → 新規作成
     - 出力フォルダ確認 → 新規作成（必要時）
     - スクリプトプロパティ更新

2. **詳細設定**
   - レポートタイトル（テキスト入力）
   - Gemini APIキー（パスワード入力、マスク表示）
   - 「💾 設定を保存」ボタン
   - 「🔌 Gemini接続テスト」ボタン

3. **テンプレート管理**
   - 「✨ テンプレートを更新」ボタン
   - 確認ダイアログ表示

4. **現在の設定**
   - スライドテンプレートID（読み取り専用）
   - 出力フォルダID（読み取り専用）
   - 現在のスライドID（読み取り専用）
   - 「🔄 スライドIDをリセット」ボタン

### 7.3 AISidebar（AI機能画面）

**セクション構成：**

1. **クイック生成（Gemini 2.5 Flash）**
   - 「💬 売上サマリーコメントを生成」ボタン
   - 「💡 ビジネスインサイトを生成」ボタン

2. **カスタム生成**
   - プロンプト入力欄（テキストエリア）
   - 「🚀 生成」ボタン

3. **結果表示**
   - 成功時: 緑背景、生成テキスト、コピーボタン
   - エラー時: 赤背景、エラーメッセージ

### 7.4 MainSidebar（レポート生成画面）

**セクション構成：**

1. **データ準備**
   - CSVアップロード
   - または「🎲 サンプルデータ生成」ボタン

2. **AI機能（オプション）**
   - 「💬 AIコメント生成」ボタン
   - 「💡 AIインサイト生成」ボタン

3. **スライド生成**
   - 出力モード選択（新規作成/既存更新）
   - 「📊 レポート生成」ボタン

---

## 8. エラーハンドリング

### 8.1 エラー種別

| エラー種別 | 検出方法 | 対応 |
|-----------|---------|------|
| APIキー未設定 | `geminiApiKey`が空 | エラーメッセージ表示 |
| テンプレートID未設定 | `slideTemplateId`が空 | セットアップ促進メッセージ |
| データ不存在 | シートが空 | サンプルデータ生成を提案 |
| Gemini APIエラー | API呼び出し失敗 | エラー詳細をユーザーに表示 |
| 権限不足 | Drive/Slides API権限エラー | 権限付与を促す |

### 8.2 エラーレスポンス形式

```javascript
{
  success: false,
  message: "エラーの説明"
}
```

### 8.3 成功レスポンス形式

```javascript
{
  success: true,
  message: "成功メッセージ",
  // その他の結果データ
}
```

---

## 9. セキュリティ

### 9.1 APIキー管理

- **保存場所**: スクリプトプロパティ（暗号化）
- **表示方法**: マスク表示（`AIxx...xxYZ`形式）
- **変更方法**: フルキー入力時のみ更新
- **アクセス権限**: スクリプト実行権限を持つユーザーのみ

### 9.2 データ保護

- **生データ**: スプレッドシート内に保存
- **アクセス制御**: Google Drive の共有設定に準拠
- **外部送信**: Gemini APIのみ（分析用プロンプトのみ送信）

### 9.3 認証・認可

- **Google OAuth**: Apps Script の標準認証
- **必要なスコープ**:
  - `https://www.googleapis.com/auth/spreadsheets`
  - `https://www.googleapis.com/auth/presentations`
  - `https://www.googleapis.com/auth/drive`
  - `https://www.googleapis.com/auth/script.external_request`

---

## 10. 運用・保守

### 10.1 初期セットアップ手順

1. **スプレッドシート作成**
   - 新規スプレッドシートを作成

2. **スクリプト配置**
   - 「拡張機能」→「Apps Script」
   - `Code.gs`およびHTMLファイルを追加

3. **セットアップ実行**
   - メニュー「📊 スライドレポート」→「⚙️ 設定」
   - 「⚡ セットアップ実行」をクリック

4. **Gemini API設定（オプション）**
   - [Google AI Studio](https://makersuite.google.com/app/apikey)でAPIキー取得
   - 「Gemini APIキー」欄に入力
   - 「💾 設定を保存」
   - 「🔌 Gemini接続テスト」で確認

5. **データインポート**
   - CSV準備またはサンプルデータ生成
   - 集計シート確認

6. **初回レポート生成**
   - 「📊 レポート生成」から実行

### 10.2 定期メンテナンス

#### 月次タスク
- [ ] 新しい月のCSVデータをインポート
- [ ] 集計シート確認
- [ ] レポート生成

#### 四半期タスク
- [ ] スライドテンプレート見直し
- [ ] データ保存容量確認

#### 年次タスク
- [ ] 過去データのアーカイブ
- [ ] システム全体の見直し

### 10.3 トラブルシューティング

| 問題 | 原因 | 解決策 |
|------|------|--------|
| スライド生成エラー | テンプレートIDが無効 | セットアップを再実行 |
| Gemini接続失敗 | APIキーが無効 | APIキーを再設定 |
| グラフが表示されない | データ不足 | データを確認・追加 |
| 集計シートが空 | QUERY式エラー | シートを削除して再生成 |

---

## 付録A: データサンプル

### RawSalesData（3行）
```csv
Date,Region,Person,Product,Category,Quantity,UnitPrice,TotalSales
2025-01-01,関東,田中太郎,製品A,サブスク,10,50000,500000
2025-01-02,関西,鈴木花子,製品B,単発,5,100000,500000
2025-01-03,九州,佐藤一郎,製品C,保守,8,30000,240000
```

---

## 付録B: 参考リンク

- [Google Apps Script リファレンス](https://developers.google.com/apps-script/reference)
- [Google Slides API](https://developers.google.com/slides/api)
- [Google Sheets API](https://developers.google.com/sheets/api)
- [Gemini API ドキュメント](https://ai.google.dev/docs)

---

**作成日**: 2026-01-11
**最終更新**: 2026-01-11
**バージョン**: 2.0
**Geminiモデル**: gemini-2.0-flash-exp
