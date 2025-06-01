/**
 * JINA AI ファクトチェック評価スクリプト
 * 
 * このスクリプトは以下の評価指標を計算します：
 * - 精度 (Accuracy)
 * - 適合率 (Precision)
 * - 再現率 (Recall)
 * - F1スコア
 * - 混同行列 (Confusion Matrix)
 * - レスポンス時間
 * - 信頼性スコアの分析
 */

import fs from 'fs';
import path from 'path';
import { API_CONFIG } from '../src/config.js';

// 評価結果を保存するクラス
class EvaluationResults {
  constructor() {
    this.results = [];
    this.metrics = {};
    this.confusionMatrix = {
      truePositive: 0,
      falsePositive: 0,
      trueNegative: 0,
      falseNegative: 0
    };
    this.responseTimes = [];
    this.factualityScores = [];
  }

  addResult(testCase, prediction, responseTime, factuality) {
    const result = {
      id: testCase.id,
      claim: testCase.claim,
      groundTruth: testCase.ground_truth,
      prediction: prediction,
      correct: prediction === testCase.ground_truth,
      responseTime: responseTime,
      factuality: factuality,
      category: testCase.category,
      difficulty: testCase.difficulty,
      language: testCase.language
    };

    this.results.push(result);
    this.responseTimes.push(responseTime);
    
    if (factuality !== null) {
      this.factualityScores.push(factuality);
    }

    // 混同行列の更新
    if (testCase.ground_truth === true && prediction === true) {
      this.confusionMatrix.truePositive++;
    } else if (testCase.ground_truth === false && prediction === false) {
      this.confusionMatrix.trueNegative++;
    } else if (testCase.ground_truth === false && prediction === true) {
      this.confusionMatrix.falsePositive++;
    } else if (testCase.ground_truth === true && prediction === false) {
      this.confusionMatrix.falseNegative++;
    }
  }

  calculateMetrics() {
    const totalTests = this.results.length;
    const correctPredictions = this.results.filter(r => r.correct).length;
    
    // 基本指標
    this.metrics.accuracy = correctPredictions / totalTests;
    
    const { truePositive, falsePositive, trueNegative, falseNegative } = this.confusionMatrix;
    
    this.metrics.precision = truePositive / (truePositive + falsePositive) || 0;
    this.metrics.recall = truePositive / (truePositive + falseNegative) || 0;
    this.metrics.f1Score = 2 * (this.metrics.precision * this.metrics.recall) / 
                          (this.metrics.precision + this.metrics.recall) || 0;
    
    // レスポンス時間統計
    this.metrics.avgResponseTime = this.responseTimes.reduce((a, b) => a + b, 0) / this.responseTimes.length;
    this.metrics.maxResponseTime = Math.max(...this.responseTimes);
    this.metrics.minResponseTime = Math.min(...this.responseTimes);
    
    // 信頼性スコア統計
    if (this.factualityScores.length > 0) {
      this.metrics.avgFactuality = this.factualityScores.reduce((a, b) => a + b, 0) / this.factualityScores.length;
      this.metrics.factualityStdDev = Math.sqrt(
        this.factualityScores.reduce((sq, n) => sq + Math.pow(n - this.metrics.avgFactuality, 2), 0) / 
        this.factualityScores.length
      );
    }

    // カテゴリ別精度
    this.metrics.categoryAccuracy = {};
    const categories = [...new Set(this.results.map(r => r.category))];
    
    categories.forEach(category => {
      const categoryResults = this.results.filter(r => r.category === category);
      const categoryCorrect = categoryResults.filter(r => r.correct).length;
      this.metrics.categoryAccuracy[category] = categoryCorrect / categoryResults.length;
    });

    // 難易度別精度
    this.metrics.difficultyAccuracy = {};
    const difficulties = [...new Set(this.results.map(r => r.difficulty))];
    
    difficulties.forEach(difficulty => {
      const difficultyResults = this.results.filter(r => r.difficulty === difficulty);
      const difficultyCorrect = difficultyResults.filter(r => r.correct).length;
      this.metrics.difficultyAccuracy[difficulty] = difficultyCorrect / difficultyResults.length;
    });

    // 言語別精度
    this.metrics.languageAccuracy = {};
    const languages = [...new Set(this.results.map(r => r.language))];
    
    languages.forEach(language => {
      const languageResults = this.results.filter(r => r.language === language);
      const languageCorrect = languageResults.filter(r => r.correct).length;
      this.metrics.languageAccuracy[language] = languageCorrect / languageResults.length;
    });
  }

  generateReport() {
    const report = {
      summary: {
        totalTests: this.results.length,
        correctPredictions: this.results.filter(r => r.correct).length,
        accuracy: this.metrics.accuracy,
        precision: this.metrics.precision,
        recall: this.metrics.recall,
        f1Score: this.metrics.f1Score
      },
      confusionMatrix: this.confusionMatrix,
      performance: {
        avgResponseTime: this.metrics.avgResponseTime,
        maxResponseTime: this.metrics.maxResponseTime,
        minResponseTime: this.metrics.minResponseTime,
        avgFactuality: this.metrics.avgFactuality,
        factualityStdDev: this.metrics.factualityStdDev
      },
      detailedAnalysis: {
        categoryAccuracy: this.metrics.categoryAccuracy,
        difficultyAccuracy: this.metrics.difficultyAccuracy,
        languageAccuracy: this.metrics.languageAccuracy
      },
      rawResults: this.results
    };

    return report;
  }
}

// JINA APIを呼び出す関数（既存のコードから流用）
async function callJinaFactCheck(claim) {
  const JINA_TOKEN = API_CONFIG.JINA_API_TOKEN;
  const endpoint = "https://deepsearch.jina.ai/v1/chat/completions";

  const body = {
    model: "jina-chat",
    messages: [
      {
        role: "user",
        content: `Please fact-check the following claim and return the result in JSON format with these fields:
- "result": boolean (true if the claim is accurate, false if not)
- "reason": string (explanation in Japanese)
- "factuality": number between 0 and 1
- "references": array of objects with "url", "keyQuote", and "isSupportive" fields

Claim to fact-check: "${claim}"`
      }
    ],
    stream: false,
    temperature: 0,
    search: true
  };

  const startTime = Date.now();

  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);
    
    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${JINA_TOKEN}`
      },
      body: JSON.stringify(body),
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    const responseTime = Date.now() - startTime;

    if (!res.ok) {
      return { hit: false, error: `HTTP ${res.status}`, responseTime };
    }

    const data = await res.json();
    
    let responseData;
    if (data.choices && data.choices[0] && data.choices[0].message) {
      let content = data.choices[0].message.content;
      
      try {
        content = content.replace(/^```json\s*\n?/, '').replace(/\n?```\s*$/, '');
        responseData = JSON.parse(content);
      } catch (e) {
        return { hit: false, error: "JSON parsing failed", responseTime };
      }
    }
    
    if (!responseData) {
      return { hit: false, error: "No valid response data", responseTime };
    }

    return {
      hit: true,
      result: responseData.result ?? "",
      reason: responseData.reason ?? "",
      factuality: responseData.factuality ?? null,
      references: responseData.references ?? [],
      responseTime: responseTime
    };
  } catch (e) {
    const responseTime = Date.now() - startTime;
    if (e.name === 'AbortError') {
      return { hit: false, error: "Timeout", responseTime };
    }
    return { hit: false, error: e.message || String(e), responseTime };
  }
}

// メイン評価関数
async function runEvaluation() {
  console.log("🔍 JINA AI ファクトチェック評価を開始します...\n");

  // テストデータセットを読み込み
  const datasetPath = path.join(process.cwd(), 'evaluation', 'test-dataset.json');
  const dataset = JSON.parse(fs.readFileSync(datasetPath, 'utf8'));
  const testCases = dataset.factcheck_evaluation_dataset.test_cases;

  const evaluator = new EvaluationResults();
  
  console.log(`📋 総テストケース数: ${testCases.length}`);
  console.log("=" * 50);

  // 各テストケースを実行
  for (let i = 0; i < testCases.length; i++) {
    const testCase = testCases[i];
    console.log(`\n[${i + 1}/${testCases.length}] テスト実行中...`);
    console.log(`Claims: "${testCase.claim}"`);
    console.log(`正解: ${testCase.ground_truth}`);
    
    const result = await callJinaFactCheck(testCase.claim);
    
    if (result.hit) {
      console.log(`予測: ${result.result}`);
      console.log(`信頼度: ${result.factuality}`);
      console.log(`レスポンス時間: ${result.responseTime}ms`);
      console.log(`正解: ${result.result === testCase.ground_truth ? '✅' : '❌'}`);
      
      evaluator.addResult(testCase, result.result, result.responseTime, result.factuality);
    } else {
      console.log(`❌ エラー: ${result.error}`);
      console.log(`レスポンス時間: ${result.responseTime}ms`);
      
      // エラーの場合はnullとして扱う
      evaluator.addResult(testCase, null, result.responseTime, null);
    }

    // APIレート制限を考慮して1秒待機
    await new Promise(resolve => setTimeout(resolve, 1000));
  }

  // 評価指標を計算
  evaluator.calculateMetrics();
  
  // レポートを生成
  const report = evaluator.generateReport();
  
  // 結果を表示
  console.log("\n" + "=" * 50);
  console.log("📊 評価結果サマリー");
  console.log("=" * 50);
  console.log(`総テスト数: ${report.summary.totalTests}`);
  console.log(`正解数: ${report.summary.correctPredictions}`);
  console.log(`精度 (Accuracy): ${(report.summary.accuracy * 100).toFixed(2)}%`);
  console.log(`適合率 (Precision): ${(report.summary.precision * 100).toFixed(2)}%`);
  console.log(`再現率 (Recall): ${(report.summary.recall * 100).toFixed(2)}%`);
  console.log(`F1スコア: ${(report.summary.f1Score * 100).toFixed(2)}%`);
  
  console.log("\n📈 パフォーマンス");
  console.log(`平均レスポンス時間: ${report.performance.avgResponseTime.toFixed(0)}ms`);
  console.log(`最大レスポンス時間: ${report.performance.maxResponseTime}ms`);
  console.log(`最小レスポンス時間: ${report.performance.minResponseTime}ms`);
  
  if (report.performance.avgFactuality) {
    console.log(`平均信頼度スコア: ${report.performance.avgFactuality.toFixed(3)}`);
    console.log(`信頼度標準偏差: ${report.performance.factualityStdDev.toFixed(3)}`);
  }

  console.log("\n📊 混同行列");
  console.log(`True Positive: ${report.confusionMatrix.truePositive}`);
  console.log(`False Positive: ${report.confusionMatrix.falsePositive}`);
  console.log(`True Negative: ${report.confusionMatrix.trueNegative}`);
  console.log(`False Negative: ${report.confusionMatrix.falseNegative}`);

  console.log("\n🏷️ カテゴリ別精度");
  Object.entries(report.detailedAnalysis.categoryAccuracy).forEach(([category, accuracy]) => {
    console.log(`${category}: ${(accuracy * 100).toFixed(2)}%`);
  });

  console.log("\n📊 難易度別精度");
  Object.entries(report.detailedAnalysis.difficultyAccuracy).forEach(([difficulty, accuracy]) => {
    console.log(`${difficulty}: ${(accuracy * 100).toFixed(2)}%`);
  });

  console.log("\n🌐 言語別精度");
  Object.entries(report.detailedAnalysis.languageAccuracy).forEach(([language, accuracy]) => {
    console.log(`${language}: ${(accuracy * 100).toFixed(2)}%`);
  });

  // 詳細な結果をJSONファイルに保存
  const resultsPath = path.join(process.cwd(), 'evaluation', `evaluation-results-${Date.now()}.json`);
  fs.writeFileSync(resultsPath, JSON.stringify(report, null, 2));
  console.log(`\n💾 詳細な結果を保存しました: ${resultsPath}`);

  // CSVレポートも生成
  const csvContent = generateCSVReport(report.rawResults);
  const csvPath = path.join(process.cwd(), 'evaluation', `evaluation-results-${Date.now()}.csv`);
  fs.writeFileSync(csvPath, csvContent);
  console.log(`📄 CSVレポートを保存しました: ${csvPath}`);
}

// CSV形式のレポートを生成
function generateCSVReport(results) {
  const headers = [
    'ID', 'Claim', 'Ground Truth', 'Prediction', 'Correct', 
    'Response Time (ms)', 'Factuality', 'Category', 'Difficulty', 'Language'
  ];
  
  const rows = results.map(result => [
    result.id,
    `"${result.claim.replace(/"/g, '""')}"`,
    result.groundTruth,
    result.prediction,
    result.correct,
    result.responseTime,
    result.factuality || '',
    result.category,
    result.difficulty,
    result.language
  ]);

  return [headers.join(','), ...rows.map(row => row.join(','))].join('\n');
}

// エラーハンドリング付きでメイン関数を実行
if (import.meta.url === `file://${process.argv[1]}`) {
  runEvaluation().catch(error => {
    console.error('❌ 評価中にエラーが発生しました:', error);
    process.exit(1);
  });
}

export { runEvaluation, EvaluationResults };