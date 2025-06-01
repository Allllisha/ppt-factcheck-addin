# GitHub リポジトリのセットアップとVercelへのデプロイ

## 1. GitHubでリポジトリを作成

1. [GitHub](https://github.com) にログイン
2. 右上の「+」アイコン → 「New repository」をクリック
3. リポジトリ名を入力（例：`ppt-factcheck-addin`）
4. Private または Public を選択
5. 「Create repository」をクリック（READMEは追加しない）

## 2. ローカルリポジトリをGitHubに接続

GitHubで作成したリポジトリのページに表示されるコマンドを使用：

```bash
cd /Users/anemoto/ppt-factcheck-addin/factcheck
git remote add origin https://github.com/YOUR_USERNAME/ppt-factcheck-addin.git
git branch -M main
git push -u origin main
```

## 3. Vercelにデプロイ

1. [Vercel](https://vercel.com) にログイン
2. 「New Project」をクリック
3. 「Import Git Repository」から先ほど作成したGitHubリポジトリを選択
4. 以下の設定を確認：
   - Framework Preset: `Other`
   - Root Directory: `.` （変更不要）
   - Build Command: `npm run build:vercel`
   - Output Directory: `dist`
   - Install Command: `npm install`

5. 「Environment Variables」をクリックして以下を追加：
   - `JINA_API_TOKEN`
   - `TAVILY_API_KEY`
   - `GOOGLE_API_KEY`
   - `GOOGLE_SEARCH_ENGINE_ID`

6. 「Deploy」をクリック

## 4. デプロイ完了後の設定

1. Vercelのダッシュボードでプロジェクトのドメインを確認（例：`ppt-factcheck-addin.vercel.app`）

2. `manifest-production.xml` を更新：
   ```bash
   # manifest-production.xml内の YOUR-APP-NAME を実際のドメイン名に置き換える
   # 例：https://YOUR-APP-NAME.vercel.app → https://ppt-factcheck-addin.vercel.app
   ```

3. 更新したmanifestをGitHubにプッシュ：
   ```bash
   git add manifest-production.xml
   git commit -m "Update manifest with Vercel domain"
   git push
   ```

## 5. PowerPointでアドインを追加

1. `manifest-production.xml` をダウンロード
2. PowerPointを開く
3. Insert → My Add-ins → Upload My Add-in
4. ダウンロードした `manifest-production.xml` を選択

## トラブルシューティング

### ビルドエラーが発生する場合
- Vercelのビルドログを確認
- 環境変数が正しく設定されているか確認

### アドインが読み込まれない場合
- manifest.xmlのURLが正しいか確認
- ブラウザの開発者ツールでエラーを確認
- CORSエラーが出ていないか確認

### APIが動作しない場合
- Vercelの環境変数が正しく設定されているか確認
- Vercelのファンクションログを確認