# CLAUDE.md - GAS スライドレポート自動生成システム

## プロジェクト概要

Google Apps Script (GAS) で CSV データから Google Slides の月次売上レポートを自動生成するシステム。
Gemini AI 連携によるコメント・インサイト自動生成機能付き。

**スクリプトID**: `1nSclBxxqad-tXnTyyoQVeK978AjxfbNCAIrE_k8hZTYMNe0vJZTe-Swe`

## フォルダ構造

```
.
├── CLAUDE.md              # このファイル
├── README.md              # プロジェクト README
├── .clasp.json            # clasp 設定（スクリプトID、rootDir: src）
├── .claspignore           # clasp push 時の除外ファイル
├── .github/               # GitHub Actions
│   ├── SETUP_SECRETS.md     # Secrets セットアップガイド
│   └── workflows/
│       └── deploy-gas.yml   # clasp push 自動デプロイ
├── docs/                  # ドキュメント
│   ├── seminar-overview.md  # セミナー資料
│   ├── context.md           # プロジェクトコンテキスト
│   ├── requirements.md      # 要件定義書
│   ├── setup-guide.md       # セットアップガイド
│   └── new-features.md      # 新機能 README
├── src/                   # GAS ソースコード（clasp rootDir）
│   ├── appsscript.json      # GAS マニフェスト
│   ├── core/
│   │   └── Code.gs          # メインスクリプト（全関数）
│   └── ui/
│       ├── AISidebar.html       # AI アシスタント サイドバー
│       ├── MainSidebar.html     # レポート生成 サイドバー
│       ├── SettingsSidebar.html # 設定 サイドバー
│       └── dialogs/
│           ├── ConfigDialog.html
│           ├── DummyDataDialog.html
│           ├── HelpDialog.html
│           ├── MainDialog.html
│           ├── ReportDialog.html
│           ├── SampleDataDialog.html
│           ├── SettingsDialog.html
│           ├── SetupDialog.html
│           └── TriggerDialog.html
└── data/                  # データファイル
    ├── generate_realistic_data.py  # サンプルデータ生成スクリプト
    └── sales_data_2025.csv         # サンプル CSV データ
```

## 技術スタック

- **Google Apps Script** (V8 ランタイム, ES6+)
- **Google Sheets API** - データ管理・QUERY 関数による集計
- **Google Slides API** - スライド生成・テキスト置換・チャート挿入
- **Google Drive API** - ファイル管理
- **Gemini API** (`gemini-2.0-flash-exp`) - AI コメント・インサイト生成
- **HTML/CSS/JavaScript** - サイドバー・ダイアログ UI

## 開発コマンド

```bash
# clasp ログイン
npx @google/clasp login

# GAS プロジェクトからコードを取得
npx @google/clasp pull

# GAS プロジェクトにコードをプッシュ
npx @google/clasp push

# GAS エディタを開く
npx @google/clasp open
```

## 主要な関数（src/core/Code.gs）

### セットアップ
- `setupInitialEnvironment()` - 初期セットアップ（テンプレート & フォルダ作成）
- `updateSlideTemplate()` - テンプレート更新

### データ管理
- `importAndAggregateSalesData(csvText)` - CSV インポート & 集計
- `refreshAggregationSheets()` - 集計シート更新
- `getReportData_(periodType, targetDate)` - レポートデータ取得

### スライド生成
- `generateOrUpdateSlide(params)` - スライド生成・更新
- `createNewSlide_(config, data)` - 新規スライド作成
- `updateSlide_(presentation, data)` - 既存スライド更新
- `insertChartsFromSheet_(presentation)` - チャート挿入

### Gemini AI
- `generateTextWithGemini(prompt, customPrompt)` - テキスト生成
- `generateAIComment()` - 売上サマリーコメント生成
- `generateAIInsight()` - ビジネスインサイト生成
- `testGeminiConnection()` - 接続テスト

### 設定管理
- `getConfigForUI()` - UI 用設定取得
- `saveConfigFromUI(config)` - UI から設定保存

## コーディング規約

- **関数名**: camelCase（例: `createDataObject`）
- **定数**: UPPER_SNAKE_CASE（例: `FORM_CONFIG`）
- **変数**: camelCase（例: `rowData`）
- **HTML ファイル**: PascalCase（例: `SettingsDialog.html`）
- **プライベート関数**: 末尾アンダースコア（例: `getConfig_()`）
- JSDoc でドキュメント記述

## スクリプトプロパティ

| キー | 説明 |
|------|------|
| `SLIDE_TEMPLATE_ID` | スライドテンプレート ID |
| `CURRENT_SLIDE_ID` | 最後に生成したスライド ID |
| `OUTPUT_FOLDER_ID` | 出力先フォルダ ID |
| `GEMINI_API_KEY` | Gemini API キー |
| `REPORT_TITLE` | レポートタイトル |
| `PERIOD_TYPE` | 期間タイプ（monthly/weekly/yearly） |

## スライド変数（テンプレート置換）

`{{reportTitle}}`, `{{period}}`, `{{generatedAt}}`, `{{totalSales}}`,
`{{totalSalesChange}}`, `{{topRegion}}`, `{{topRegionSales}}`,
`{{topPerson}}`, `{{topPersonSales}}`, `{{aiComment}}`, `{{aiInsight}}`

## 注意事項

- `src/` ディレクトリが clasp の rootDir
- `docs/`, `data/`, `README.md` は clasp push 対象外（`.claspignore` で除外）
- API キーはスクリプトプロパティに保存、コードにハードコーディングしない
- Gemini API 無料枠: 1分あたり 60 リクエスト
