# GitHub Secrets セットアップガイド

GitHub Actions で clasp デプロイを動かすために、以下の Secrets を設定してください。

## 必須 Secrets

### `CLASPRC_JSON_BASE64`

clasp の認証情報（base64 エンコード）。

**取得手順:**

1. ローカルで clasp にログイン:
   ```bash
   npx @google/clasp login
   ```

2. 生成された `~/.clasprc.json` を base64 エンコード:
   ```bash
   # macOS
   cat ~/.clasprc.json | base64

   # Linux
   cat ~/.clasprc.json | base64 -w 0
   ```

3. 出力された文字列を GitHub Secrets に `CLASPRC_JSON_BASE64` として保存

### 設定場所

1. GitHub リポジトリ → **Settings** → **Secrets and variables** → **Actions**
2. **New repository secret** をクリック
3. Name: `CLASPRC_JSON_BASE64`
4. Value: 上記で取得した base64 文字列
5. **Add secret**

## オプション Secrets

### `CLASP_JSON`

リポジトリに `.clasp.json` を含めない場合に使用。

```json
{
  "scriptId": "YOUR_SCRIPT_ID",
  "rootDir": "src"
}
```

### `DISCORD_WEBHOOK_URL`

デプロイ結果を Discord に通知する場合に設定。

**取得手順:**

1. Discord サーバー → チャンネル設定 → **連携サービス** → **ウェブフック**
2. **新しいウェブフック** を作成
3. **ウェブフックURLをコピー**
4. GitHub Secrets に `DISCORD_WEBHOOK_URL` として保存

## 注意事項

- `~/.clasprc.json` にはアクセストークンとリフレッシュトークンが含まれます。絶対にコードにコミットしないでください
- トークンは定期的に期限切れになります。デプロイが認証エラーで失敗した場合は `clasp login` を再実行して Secrets を更新してください
- ワークフロー終了時に認証情報は自動クリーンアップされます
