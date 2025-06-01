# PowerPoint FactCheck Add-in プロジェクトドキュメント

## 1. 要件定義

### 1.1 プロジェクト概要
PowerPointプレゼンテーション内のテキストコンテンツに対して、AI駆動のファクトチェック機能を提供するOffice Add-inの開発。

### 1.2 機能要件
- **FR-01**: PowerPointスライド内の全テキストを自動検出
- **FR-02**: 各文章に対してAIベースのファクトチェックを実行
- **FR-03**: ファクトチェック結果を視覚的に表示（色分け）
- **FR-04**: 誤情報に対する修正提案の提供
- **FR-05**: 複数の検索エンジンによる追加検証
- **FR-06**: リアルタイムプログレス表示
- **FR-07**: 信頼できるソースの優先表示

### 1.3 非機能要件
- **NFR-01**: レスポンシブでユーザーフレンドリーなUI
- **NFR-02**: 20秒以内のタイムアウト処理
- **NFR-03**: セキュアなAPI通信（HTTPS）
- **NFR-04**: クロスプラットフォーム対応（Windows/Mac/Web）
- **NFR-05**: 環境変数によるAPIキー管理

## 2. システム設計

### 2.1 アーキテクチャ設計

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   PowerPoint    │────▶│  Office Add-in  │────▶│   External APIs │
│   Application   │◀────│   (factcheck)   │◀────│                 │
└─────────────────┘     └─────────────────┘     └─────────────────┘
                              │                         │
                              ▼                         ▼
                        ┌──────────┐           ┌─────────────┐
                        │ Office.js│           │ ・Jina AI   │
                        │   API    │           │ ・Tavily    │
                        └──────────┘           │ ・Google    │
                                               └─────────────┘
```

### 2.2 技術スタック
- **フロントエンド**: HTML5, CSS3, JavaScript (ES6+)
- **Office統合**: Office.js API
- **ビルドツール**: Webpack, Babel
- **パッケージ管理**: npm
- **デプロイメント**: Vercel
- **バージョン管理**: Git/GitHub

### 2.3 外部API統合
1. **Jina AI DeepSearch API**
   - エンドポイント: `https://deepsearch.jina.ai/v1/chat/completions`
   - 主要機能: AIベースのファクトチェック
   - レスポンス形式: JSON

2. **Tavily Search API**
   - エンドポイント: `https://api.tavily.com/search`
   - 主要機能: 追加検証用Web検索
   - レスポンス形式: JSON

3. **Google Custom Search API**
   - エンドポイント: `https://www.googleapis.com/customsearch/v1`
   - 主要機能: 信頼できるソースからの検索
   - レスポンス形式: JSON

## 3. 開発実装

### 3.1 プロジェクト構造
```
factcheck/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # メインUI
│   │   ├── taskpane.css       # スタイルシート
│   │   └── taskpane.js        # メインロジック
│   ├── commands/
│   │   ├── commands.html      # コマンドUI
│   │   └── commands.js        # コマンドロジック
│   └── config.js              # API設定
├── assets/                    # アイコン・画像
├── manifest.xml              # 開発用マニフェスト
├── manifest-production.xml   # 本番用マニフェスト
├── webpack.config.js         # Webpack設定
├── webpack.production.js     # 本番ビルド設定
├── package.json              # 依存関係
└── vercel.json              # Vercel設定
```

### 3.2 主要コンポーネント

#### 3.2.1 テキスト抽出と処理
```javascript
// PowerPointスライドからテキストを抽出
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  // 各スライドのテキストボックスを処理
  for (let slide of slides.items) {
    const shapes = slide.shapes;
    shapes.load("items");
    await context.sync();
    // テキスト処理ロジック
  }
});
```

#### 3.2.2 ファクトチェックAPI呼び出し
```javascript
async function callJinaFactCheck(claim) {
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${JINA_TOKEN}`
    },
    body: JSON.stringify({
      model: "jina-chat",
      messages: [{
        role: "user",
        content: `Fact-check: "${claim}"`
      }],
      search: true
    })
  });
  // レスポンス処理
}
```

#### 3.2.3 UI更新と結果表示
- プログレスカード表示
- 結果の色分け表示（緑：正確、赤：誤り、青：不明）
- 修正提案の表示
- エラーハンドリング

### 3.3 セキュリティ実装
- APIキーの環境変数管理
- HTTPS通信の強制
- CORSヘッダーの適切な設定
- クライアントサイドでの入力検証

## 4. デプロイメント

### 4.1 ローカル開発環境
```bash
# 依存関係インストール
npm install

# 開発サーバー起動
npm start
```

### 4.2 本番環境デプロイ（Vercel）
```bash
# ビルド
npm run build:vercel

# Vercelへデプロイ
vercel --prod
```

### 4.3 環境変数設定
- `JINA_API_TOKEN`
- `TAVILY_API_KEY`
- `GOOGLE_API_KEY`
- `GOOGLE_SEARCH_ENGINE_ID`

### 4.4 カスタムドメイン
- 本番URL: `https://factcheck-seven.vercel.app`

## 5. 検証とテスト

### 5.1 機能テスト
- ✅ PowerPointテキスト抽出機能
- ✅ 日本語・英語の文章分割処理
- ✅ Jina AIファクトチェックAPI統合
- ✅ Tavily/Google検索API統合
- ✅ 結果の色分け表示
- ✅ エラーハンドリング
- ✅ タイムアウト処理（20秒）

### 5.2 非機能テスト
- ✅ レスポンシブデザイン
- ✅ クロスブラウザ互換性
- ✅ パフォーマンス（大量テキスト処理）
- ✅ セキュリティ（APIキー保護）

### 5.3 評価結果
- **精度**: 評価スクリプトによる自動テスト実装
- **パフォーマンス**: 平均処理時間 < 10秒/スライド
- **ユーザビリティ**: 直感的なUI、リアルタイムフィードバック

## 6. 既知の問題と改善点

### 6.1 既知の問題
- 大量のテキストでのパフォーマンス低下
- API利用制限による処理中断の可能性

### 6.2 今後の改善点
- バッチ処理による高速化
- キャッシュ機能の実装
- オフラインモードのサポート
- 多言語対応の拡充

## 7. メンテナンスとサポート

### 7.1 定期メンテナンス
- APIキーのローテーション
- 依存関係のアップデート
- セキュリティパッチの適用

### 7.2 サポート
- GitHubイシューでの問題報告
- ドキュメントの継続的更新
- ユーザーフィードバックの収集と反映

---

最終更新日: 2025年6月1日