# PowerPoint FactCheck Add-in

AI駆動のファクトチェック機能を提供するPowerPointアドインです。プレゼンテーションのテキストを自動的に検証し、正確性を確保します。

## 🚀 機能

- **AIファクトチェック**: Jina AI DeepSearchを使用してテキストの正確性を検証
- **リアルタイムプログレス**: 各文章の処理状況をリアルタイムで表示
- **複数検索エンジン対応**: Tavily、Google Custom Searchによる追加検索
- **視覚的フィードバック**: 結果に基づいてテキストを色分け表示
- **自動修正提案**: 誤りが検出された場合の修正案を提供
- **信頼できるソース**: 政府機関、学術機関、報道機関などの信頼できるソースを優先

## 🛠️ 技術スタック

- **フロントエンド**: HTML5, CSS3, JavaScript (ES6+)
- **Office.js**: PowerPoint API統合
- **AI/検索API**: 
  - Jina AI DeepSearch
  - Tavily Search API
  - Google Custom Search API
- **ビルドツール**: Webpack, Babel
- **スタイリング**: モダンCSS (CSS Grid, Flexbox, CSS Variables)

## 📋 前提条件

- Node.js (v14以上)
- PowerPoint (Windows/Mac/Web)
- 有効なAPIキー:
  - Jina AI API Token
  - Tavily API Key
  - Google Custom Search API Key

## 🔧 セットアップ

1. **リポジトリのクローン**
   ```bash
   git clone <repository-url>
   cd factcheck
   ```

2. **依存関係のインストール**
   ```bash
   npm install
   ```

3. **環境変数の設定**
   `.env`ファイルを作成し、APIキーを設定：
   ```env
   JINA_API_TOKEN=your_jina_token
   TAVILY_API_KEY=your_tavily_key
   GOOGLE_API_KEY=your_google_key
   GOOGLE_SEARCH_ENGINE_ID=your_search_engine_id
   ```

4. **開発サーバーの起動**
   ```bash
   npm start
   ```

## 💡 使用方法

### PowerPointにアドインを追加する方法

#### オンライン版（推奨）
本番環境にデプロイされたアドインを使用する場合：

1. **manifest-production.xml** をダウンロード
2. PowerPointを開く
3. **挿入 (Insert)** タブをクリック
4. **個人用アドイン (My Add-ins)** をクリック
5. 右上の **「カスタムアドインのアップロード」(Upload My Add-in)** をクリック
6. **「参照...」(Browse...)** をクリックして、ダウンロードした **manifest-production.xml** を選択
7. **「アップロード」(Upload)** をクリック
8. アドインがインストールされたら、**ホーム (Home)** タブに **「Show Taskpane」** ボタンが表示されます

#### ローカル開発版
開発環境でテストする場合：

```bash
npm start
```
このコマンドを実行すると、PowerPointが自動的に起動し、アドインがサイドロードされます。

### アドインの使い方

1. **Show Taskpane** ボタンをクリックしてアドインパネルを開く
2. プレゼンテーションにテキストを入力
3. **「ファクトチェック開始」** ボタンをクリック
4. 各テキストの検証結果を確認
5. 必要に応じて修正提案を適用

## 🎨 UI機能

- **折りたたみ可能なセクション**: 実行パネルとデバッグログ
- **プログレス表示**: 全体とファイル単位の進捗表示
- **結果カード**: 信頼性レベル別の色分け表示
- **検索オプション**: エラー時のTavily/Google追加検索

## 🔍 サポートされる検索ソース

### 信頼できるドメイン
- **政府機関**: .gov, .go.jp
- **教育機関**: .edu, .ac.jp
- **学術論文**: Nature, Science, PubMed, Google Scholar
- **百科事典**: Wikipedia, Britannica
- **報道機関**: Reuters, AP News, BBC, NHK
- **国際機関**: WHO, UN, OECD, World Bank

## 📁 プロジェクト構造

```
factcheck/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # メインUI
│   │   ├── taskpane.css       # スタイル
│   │   └── taskpane.js        # メインロジック
│   ├── commands/
│   └── config.js              # API設定
├── assets/                    # アイコンとリソース
├── manifest.xml              # Office アドインマニフェスト
├── webpack.config.js         # ビルド設定
└── package.json              # 依存関係
```

## 🔒 セキュリティ

- APIキーは環境変数で管理
- 機密情報は.gitignoreで除外
- セキュアなHTTPS通信

## 🚧 開発

### ビルド
```bash
npm run build
```

### リント
```bash
npm run lint
```

### デバッグ
開発者ツール（F12）でデバッグログを確認できます。

## 📝 ライセンス

このプロジェクトはMITライセンスの下で公開されています。

## 🤝 コントリビューション

1. フォークする
2. フィーチャーブランチを作成する (`git checkout -b feature/AmazingFeature`)
3. 変更をコミットする (`git commit -m 'Add some AmazingFeature'`)
4. ブランチにプッシュする (`git push origin feature/AmazingFeature`)
5. プルリクエストを開く

## 📞 サポート

問題や質問がある場合は、GitHubのIssuesセクションで報告してください。

---

**Note**: このアドインは教育・研究目的で開発されました。商用利用の際は適切なライセンスを確認してください。