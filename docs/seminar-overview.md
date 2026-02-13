# セミナー資料: GAS + Gemini AI スライドレポート自動生成

## セミナー概要

Google Apps Script と Gemini AI を活用して、CSVデータから月次売上レポートスライドを自動生成するシステムのハンズオンセミナーです。

---

## アジェンダ

### 1. システム概要（イントロ）
- GAS + Google Slides + Gemini AI の連携概要
- デモ: 完成形の紹介

### 2. 環境セットアップ
- スプレッドシートの準備
- Apps Script エディタでのファイル配置
- 初期セットアップ実行（テンプレート & フォルダ自動作成）

### 3. データインポート & 集計
- CSV形式のデータ構造
- QUERY関数による自動集計（地域別・担当者別・製品別・月次）
- グラフ自動生成

### 4. Gemini AI 連携
- APIキーの取得と設定
- 売上サマリーコメント自動生成
- ビジネスインサイト自動生成
- カスタムプロンプト

### 5. スライド自動生成
- 4枚構成のモダンデザイン
  - 表紙 / サマリー / 詳細分析 / インサイト
- テキスト置換 (`{{変数名}}`)
- チャート自動挿入

### 6. 応用 & カスタマイズ
- トリガーによる定期実行
- テンプレートのカスタマイズ
- clasp によるローカル開発

---

## 前提条件

- Google アカウント
- Google Sheets / Slides の基本操作
- （任意）Gemini API キー ([Google AI Studio](https://makersuite.google.com/app/apikey) で取得)

---

## ハンズオン手順

### Step 1: プロジェクトを開く

1. 新しい Google スプレッドシートを作成
2. 「拡張機能」→「Apps Script」でエディタを開く

### Step 2: コードを配置

以下のファイルを Apps Script エディタに追加:

| ファイル | 説明 |
|---------|------|
| `Code.gs` | メインスクリプト（`src/core/Code.gs`） |
| `SettingsSidebar.html` | 設定サイドバー（`src/ui/SettingsSidebar.html`） |
| `AISidebar.html` | AI機能サイドバー（`src/ui/AISidebar.html`） |
| `MainSidebar.html` | レポート生成サイドバー（`src/ui/MainSidebar.html`） |
| `DummyDataDialog.html` | ダミーデータ生成（`src/ui/dialogs/DummyDataDialog.html`） |
| `HelpDialog.html` | ヘルプ（`src/ui/dialogs/HelpDialog.html`） |

### Step 3: 初期セットアップ

1. スプレッドシートを**リロード**
2. メニュー「📊 スライドレポート」→「⚙️ 設定」
3. 「⚡ セットアップ実行」をクリック
4. テンプレートスライドと出力フォルダが自動作成される

### Step 4: データ投入

**方法A: ダミーデータ生成**
- メニュー「📊 スライドレポート」→「🎲 ダミーデータ生成」

**方法B: CSV インポート**
- メニュー「📊 スライドレポート」→「📊 レポート生成」からCSVアップロード

### Step 5: AI コメント生成（Gemini）

1. 設定画面で Gemini API キーを入力 → 保存
2. 「🔌 Gemini接続テスト」で確認
3. レポート生成画面で「💬 AIコメント生成」「💡 AIインサイト生成」

### Step 6: レポート生成

1. メニュー「📊 スライドレポート」→「📊 レポート生成」
2. 出力モード選択: 新規作成 or 既存更新
3. 「📊 レポート生成」ボタンをクリック
4. 生成されたスライドのリンクから確認

---

## システムアーキテクチャ

```
CSV Data → Google Sheets (QUERY集計) → GAS (処理) → Google Slides (出力)
                                          ↓
                                    Gemini API (AI生成)
```

### 処理フロー

```
1. CSVインポート
2. RawSalesData シート作成
3. QUERY関数で集計シート自動生成
   - RegionalSales (地域別)
   - PersonSales (担当者別)
   - ProductSales (製品別)
   - CategorySales (カテゴリ別)
   - MonthlySales (月次推移)
4. グラフ自動生成
5. Gemini AI でコメント・インサイト生成
6. スライド生成・更新 (テキスト置換 + チャート挿入)
```

---

## スライド構成（4枚）

| スライド | 内容 | 主な変数 |
|---------|------|---------|
| 1. 表紙 | タイトル、期間、生成日時 | `{{reportTitle}}`, `{{period}}` |
| 2. サマリー | 合計売上、成長率、トップ地域/担当者 | `{{totalSales}}`, `{{aiComment}}` |
| 3. 詳細分析 | 地域別・担当者別グラフ | チャート自動挿入 |
| 4. インサイト | AI生成のビジネスインサイト | `{{aiInsight}}` |

---

## データ構造

### 入力CSV形式
```csv
Date,Region,Person,Product,Category,Quantity,UnitPrice,TotalSales
2025-01-01,関東,田中太郎,製品A,サブスク,10,50000,500000
```

### 集計シート（自動生成）
- **RegionalSales**: 地域別売上（降順）
- **PersonSales**: 担当者別売上、件数、平均単価
- **ProductSales**: 製品別売上、販売数、平均単価
- **CategorySales**: カテゴリ別売上
- **MonthlySales**: 月次売上推移、前月比

---

## 参考リンク

- [Google Apps Script リファレンス](https://developers.google.com/apps-script/reference)
- [Google Slides API](https://developers.google.com/slides/api)
- [Gemini API ドキュメント](https://ai.google.dev/docs)
- [Google AI Studio (APIキー取得)](https://makersuite.google.com/app/apikey)

---

**作成日**: 2026-02-13
**スクリプトID**: `1nSclBxxqad-tXnTyyoQVeK978AjxfbNCAIrE_k8hZTYMNe0vJZTe-Swe`
