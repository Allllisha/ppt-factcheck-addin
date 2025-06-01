# PowerPoint FactCheck Add-in をVercelにデプロイする手順

## 前提条件
- Vercelアカウントを持っていること
- Git/GitHubアカウントを持っていること
- Node.js がインストールされていること

## デプロイ手順

### 1. APIキーを環境変数として保護する

`src/config.js` を更新して環境変数を使用するようにします：

```javascript
export const API_CONFIG = {
  JINA_API_TOKEN: process.env.JINA_API_TOKEN || "",
  TAVILY_API_KEY: process.env.TAVILY_API_KEY || "",
  GOOGLE_API_KEY: process.env.GOOGLE_API_KEY || "",
  GOOGLE_SEARCH_ENGINE_ID: process.env.GOOGLE_SEARCH_ENGINE_ID || ""
};
```

### 2. GitHubにプロジェクトをプッシュ

```bash
cd factcheck
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin YOUR_GITHUB_REPO_URL
git push -u origin main
```

### 3. Vercelにデプロイ

1. [Vercel](https://vercel.com) にログイン
2. "New Project" をクリック
3. GitHubリポジトリをインポート
4. プロジェクト設定で以下を確認：
   - Framework Preset: `Other`
   - Build Command: `npm run build`
   - Output Directory: `dist`
   - Install Command: `npm install`

### 4. 環境変数を設定

Vercelダッシュボードで：
1. プロジェクトの Settings → Environment Variables に移動
2. 以下の環境変数を追加：
   - `JINA_API_TOKEN`
   - `TAVILY_API_KEY`
   - `GOOGLE_API_KEY`
   - `GOOGLE_SEARCH_ENGINE_ID`

### 5. manifest.xmlを更新

`manifest-production.xml` の `YOUR-APP-NAME` を実際のVercelアプリ名に置き換えます。
例：`https://ppt-factcheck.vercel.app`

### 6. PowerPointにアドインを追加

1. デプロイ完了後、VercelのURLをコピー
2. `manifest-production.xml` をダウンロード
3. PowerPointを開く
4. Insert → My Add-ins → Upload My Add-in
5. `manifest-production.xml` をアップロード

## 注意事項

- APIキーは必ず環境変数として設定し、コードに直接記載しない
- CORSヘッダーは `vercel.json` で設定済み
- HTTPSは Vercel が自動的に提供
- カスタムドメインを使用する場合は、manifest.xml のURLも更新する

## トラブルシューティング

### アドインが読み込まれない場合
1. ブラウザの開発者ツールでエラーを確認
2. manifest.xml のURLが正しいか確認
3. Vercelのログでビルドエラーがないか確認

### APIが動作しない場合
1. 環境変数が正しく設定されているか確認
2. Vercelダッシュボードで Functions ログを確認