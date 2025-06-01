/**
 * JINA AI ファクトチェック大規模評価スクリプト (100件)
 * 
 * 100件のテストケースで包括的な性能評価を実行
 */

import fs from 'fs';
import path from 'path';
import { API_CONFIG } from '../src/config.js';

// 評価結果を保存するクラス（改良版）
class LargeScaleEvaluationResults {
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
    this.errors = [];
    this.startTime = Date.now();
  }

  addResult(testCase, prediction, responseTime, factuality, error = null) {
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
      language: testCase.language,
      error: error
    };

    this.results.push(result);
    this.responseTimes.push(responseTime);
    
    if (factuality !== null) {
      this.factualityScores.push(factuality);
    }

    if (error) {
      this.errors.push({ id: testCase.id, error: error });
    }

    // 混同行列の更新（エラーは除外）
    if (prediction !== null) {
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
  }

  calculateMetrics() {
    const totalTests = this.results.length;
    const validResults = this.results.filter(r => r.prediction !== null);
    const correctPredictions = validResults.filter(r => r.correct).length;
    
    // 基本指標
    this.metrics.accuracy = correctPredictions / validResults.length;
    this.metrics.errorRate = (totalTests - validResults.length) / totalTests;
    
    const { truePositive, falsePositive, trueNegative, falseNegative } = this.confusionMatrix;
    
    this.metrics.precision = truePositive / (truePositive + falsePositive) || 0;
    this.metrics.recall = truePositive / (truePositive + falseNegative) || 0;
    this.metrics.f1Score = 2 * (this.metrics.precision * this.metrics.recall) / 
                          (this.metrics.precision + this.metrics.recall) || 0;
    this.metrics.specificity = trueNegative / (trueNegative + falsePositive) || 0;
    
    // レスポンス時間統計
    this.metrics.avgResponseTime = this.responseTimes.reduce((a, b) => a + b, 0) / this.responseTimes.length;
    this.metrics.maxResponseTime = Math.max(...this.responseTimes);
    this.metrics.minResponseTime = Math.min(...this.responseTimes);
    this.metrics.medianResponseTime = this.calculateMedian(this.responseTimes);
    this.metrics.responseTimeStdDev = this.calculateStdDev(this.responseTimes);
    
    // 信頼性スコア統計
    if (this.factualityScores.length > 0) {
      this.metrics.avgFactuality = this.factualityScores.reduce((a, b) => a + b, 0) / this.factualityScores.length;
      this.metrics.factualityStdDev = this.calculateStdDev(this.factualityScores);
      this.metrics.factualityMedian = this.calculateMedian(this.factualityScores);
    }

    // カテゴリ別精度
    this.calculateCategoryMetrics();
    
    // 難易度別精度
    this.calculateDifficultyMetrics();

    // 言語別精度
    this.calculateLanguageMetrics();

    // エラー分析
    this.metrics.errorsByCategory = this.analyzeErrorsByCategory();
    this.metrics.errorsByDifficulty = this.analyzeErrorsByDifficulty();

    // 全体評価時間
    this.metrics.totalEvaluationTime = Date.now() - this.startTime;
  }

  calculateMedian(arr) {
    const sorted = [...arr].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  }

  calculateStdDev(arr) {
    const mean = arr.reduce((a, b) => a + b, 0) / arr.length;
    return Math.sqrt(arr.reduce((sq, n) => sq + Math.pow(n - mean, 2), 0) / arr.length);
  }

  calculateCategoryMetrics() {
    this.metrics.categoryAccuracy = {};
    this.metrics.categoryResponseTime = {};
    const categories = [...new Set(this.results.map(r => r.category))];
    
    categories.forEach(category => {
      const categoryResults = this.results.filter(r => r.category === category && r.prediction !== null);
      const categoryCorrect = categoryResults.filter(r => r.correct).length;
      const categoryTimes = categoryResults.map(r => r.responseTime);
      
      this.metrics.categoryAccuracy[category] = categoryCorrect / categoryResults.length;
      this.metrics.categoryResponseTime[category] = categoryTimes.reduce((a, b) => a + b, 0) / categoryTimes.length;
    });
  }

  calculateDifficultyMetrics() {
    this.metrics.difficultyAccuracy = {};
    this.metrics.difficultyResponseTime = {};
    const difficulties = [...new Set(this.results.map(r => r.difficulty))];
    
    difficulties.forEach(difficulty => {
      const difficultyResults = this.results.filter(r => r.difficulty === difficulty && r.prediction !== null);
      const difficultyCorrect = difficultyResults.filter(r => r.correct).length;
      const difficultyTimes = difficultyResults.map(r => r.responseTime);
      
      this.metrics.difficultyAccuracy[difficulty] = difficultyCorrect / difficultyResults.length;
      this.metrics.difficultyResponseTime[difficulty] = difficultyTimes.reduce((a, b) => a + b, 0) / difficultyTimes.length;
    });
  }

  calculateLanguageMetrics() {
    this.metrics.languageAccuracy = {};
    this.metrics.languageResponseTime = {};
    const languages = [...new Set(this.results.map(r => r.language))];
    
    languages.forEach(language => {
      const languageResults = this.results.filter(r => r.language === language && r.prediction !== null);
      const languageCorrect = languageResults.filter(r => r.correct).length;
      const languageTimes = languageResults.map(r => r.responseTime);
      
      this.metrics.languageAccuracy[language] = languageCorrect / languageResults.length;
      this.metrics.languageResponseTime[language] = languageTimes.reduce((a, b) => a + b, 0) / languageTimes.length;
    });
  }

  analyzeErrorsByCategory() {
    const errorsByCategory = {};
    this.errors.forEach(error => {
      const result = this.results.find(r => r.id === error.id);
      if (result) {
        if (!errorsByCategory[result.category]) {
          errorsByCategory[result.category] = 0;
        }
        errorsByCategory[result.category]++;
      }
    });
    return errorsByCategory;
  }

  analyzeErrorsByDifficulty() {
    const errorsByDifficulty = {};
    this.errors.forEach(error => {
      const result = this.results.find(r => r.id === error.id);
      if (result) {
        if (!errorsByDifficulty[result.difficulty]) {
          errorsByDifficulty[result.difficulty] = 0;
        }
        errorsByDifficulty[result.difficulty]++;
      }
    });
    return errorsByDifficulty;
  }

  generateComprehensiveReport() {
    const report = {
      metadata: {
        evaluationDate: new Date().toISOString(),
        totalTestCases: this.results.length,
        successfulTests: this.results.filter(r => r.prediction !== null).length,
        errorCount: this.errors.length,
        totalEvaluationTime: this.metrics.totalEvaluationTime
      },
      summary: {
        accuracy: this.metrics.accuracy,
        precision: this.metrics.precision,
        recall: this.metrics.recall,
        f1Score: this.metrics.f1Score,
        specificity: this.metrics.specificity,
        errorRate: this.metrics.errorRate
      },
      performance: {
        avgResponseTime: this.metrics.avgResponseTime,
        medianResponseTime: this.metrics.medianResponseTime,
        maxResponseTime: this.metrics.maxResponseTime,
        minResponseTime: this.metrics.minResponseTime,
        responseTimeStdDev: this.metrics.responseTimeStdDev,
        avgFactuality: this.metrics.avgFactuality,
        factualityStdDev: this.metrics.factualityStdDev,
        factualityMedian: this.metrics.factualityMedian
      },
      confusionMatrix: this.confusionMatrix,
      detailedAnalysis: {
        categoryAccuracy: this.metrics.categoryAccuracy,
        categoryResponseTime: this.metrics.categoryResponseTime,
        difficultyAccuracy: this.metrics.difficultyAccuracy,
        difficultyResponseTime: this.metrics.difficultyResponseTime,
        languageAccuracy: this.metrics.languageAccuracy,
        languageResponseTime: this.metrics.languageResponseTime
      },
      errorAnalysis: {
        totalErrors: this.errors.length,
        errorsByCategory: this.metrics.errorsByCategory,
        errorsByDifficulty: this.metrics.errorsByDifficulty,
        errorDetails: this.errors
      },
      rawResults: this.results
    };

    return report;
  }

  printProgress(current, total) {
    const percentage = ((current / total) * 100).toFixed(1);
    const progressBar = "█".repeat(Math.floor(current / total * 20)) + "░".repeat(20 - Math.floor(current / total * 20));
    process.stdout.write(`\r[${progressBar}] ${percentage}% (${current}/${total})`);
  }
}

// JINA APIを呼び出す関数（タイムアウト調整版）
async function callJinaFactCheck(claim, timeout = 30000) {
  const JINA_TOKEN = API_CONFIG.JINA_API_TOKEN;
  const endpoint = "https://deepsearch.jina.ai/v1/chat/completions";

  const body = {
    model: "jina-chat",
    messages: [
      {
        role: "user",
        content: `Please fact-check the following claim and return the result in JSON format with these fields:
- "result": boolean (true if the claim is accurate, false if not)
- "reason": string (explanation)
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
    const timeoutId = setTimeout(() => controller.abort(), timeout);
    
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
async function runLargeScaleEvaluation() {
  console.log("🚀 JINA AI ファクトチェック大規模評価を開始します (100件)...\n");

  // 大規模テストデータセットを読み込み
  const datasetPath = path.join(process.cwd(), 'evaluation', 'large-test-dataset.json');
  const dataset = JSON.parse(fs.readFileSync(datasetPath, 'utf8'));
  const testCases = dataset.factcheck_evaluation_dataset.test_cases;

  const evaluator = new LargeScaleEvaluationResults();
  
  console.log(`📋 総テストケース数: ${testCases.length}`);
  console.log(`🌐 言語: ${[...new Set(testCases.map(t => t.language))].join(', ')}`);
  console.log(`📚 カテゴリ: ${[...new Set(testCases.map(t => t.category))].join(', ')}`);
  console.log(`📊 難易度: ${[...new Set(testCases.map(t => t.difficulty))].join(', ')}`);
  console.log("=" .repeat(80));

  // バッチ処理で実行（レート制限対策）
  const batchSize = 5;
  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < testCases.length; i += batchSize) {
    const batch = testCases.slice(i, Math.min(i + batchSize, testCases.length));
    
    console.log(`\n🔄 バッチ ${Math.floor(i / batchSize) + 1}/${Math.ceil(testCases.length / batchSize)} 処理中...`);
    
    const promises = batch.map(async (testCase) => {
      const result = await callJinaFactCheck(testCase.claim);
      
      if (result.hit) {
        evaluator.addResult(testCase, result.result, result.responseTime, result.factuality);
        return { success: true, id: testCase.id };
      } else {
        evaluator.addResult(testCase, null, result.responseTime, null, result.error);
        return { success: false, id: testCase.id, error: result.error };
      }
    });

    const results = await Promise.all(promises);
    
    results.forEach(result => {
      if (result.success) {
        successCount++;
      } else {
        errorCount++;
        console.log(`❌ ID ${result.id}: ${result.error}`);
      }
    });

    // プログレス表示
    evaluator.printProgress(i + batch.length, testCases.length);
    
    // レート制限対策（バッチ間の待機）
    if (i + batchSize < testCases.length) {
      await new Promise(resolve => setTimeout(resolve, 2000));
    }
  }

  console.log("\n\n" + "=".repeat(80));
  console.log("⏱️  評価指標を計算中...");
  
  // 評価指標を計算
  evaluator.calculateMetrics();
  
  // レポートを生成
  const report = evaluator.generateComprehensiveReport();
  
  // 結果を表示
  console.log("\n" + "=".repeat(80));
  console.log("📊 大規模評価結果サマリー (100件)");
  console.log("=".repeat(80));
  
  console.log(`\n📈 基本指標:`);
  console.log(`総テスト数: ${report.metadata.totalTestCases}`);
  console.log(`成功テスト: ${report.metadata.successfulTests}`);
  console.log(`エラー数: ${report.metadata.errorCount}`);
  console.log(`精度 (Accuracy): ${(report.summary.accuracy * 100).toFixed(2)}%`);
  console.log(`適合率 (Precision): ${(report.summary.precision * 100).toFixed(2)}%`);
  console.log(`再現率 (Recall): ${(report.summary.recall * 100).toFixed(2)}%`);
  console.log(`F1スコア: ${(report.summary.f1Score * 100).toFixed(2)}%`);
  console.log(`特異度 (Specificity): ${(report.summary.specificity * 100).toFixed(2)}%`);
  console.log(`エラー率: ${(report.summary.errorRate * 100).toFixed(2)}%`);
  
  console.log(`\n⚡ パフォーマンス:`);
  console.log(`平均レスポンス時間: ${report.performance.avgResponseTime.toFixed(0)}ms`);
  console.log(`中央値レスポンス時間: ${report.performance.medianResponseTime.toFixed(0)}ms`);
  console.log(`最大レスポンス時間: ${report.performance.maxResponseTime}ms`);
  console.log(`最小レスポンス時間: ${report.performance.minResponseTime}ms`);
  console.log(`標準偏差: ${report.performance.responseTimeStdDev.toFixed(0)}ms`);
  
  if (report.performance.avgFactuality) {
    console.log(`平均信頼度スコア: ${report.performance.avgFactuality.toFixed(3)}`);
    console.log(`信頼度中央値: ${report.performance.factualityMedian.toFixed(3)}`);
  }

  console.log(`\n📊 混同行列:`);
  console.log(`True Positive: ${report.confusionMatrix.truePositive}`);
  console.log(`False Positive: ${report.confusionMatrix.falsePositive}`);
  console.log(`True Negative: ${report.confusionMatrix.trueNegative}`);
  console.log(`False Negative: ${report.confusionMatrix.falseNegative}`);

  console.log(`\n🏷️ カテゴリ別精度:`);
  Object.entries(report.detailedAnalysis.categoryAccuracy).forEach(([category, accuracy]) => {
    const avgTime = report.detailedAnalysis.categoryResponseTime[category];
    console.log(`${category}: ${(accuracy * 100).toFixed(1)}% (平均${avgTime.toFixed(0)}ms)`);
  });

  console.log(`\n📊 難易度別精度:`);
  Object.entries(report.detailedAnalysis.difficultyAccuracy).forEach(([difficulty, accuracy]) => {
    const avgTime = report.detailedAnalysis.difficultyResponseTime[difficulty];
    console.log(`${difficulty}: ${(accuracy * 100).toFixed(1)}% (平均${avgTime.toFixed(0)}ms)`);
  });

  console.log(`\n🌐 言語別精度:`);
  Object.entries(report.detailedAnalysis.languageAccuracy).forEach(([language, accuracy]) => {
    const avgTime = report.detailedAnalysis.languageResponseTime[language];
    console.log(`${language}: ${(accuracy * 100).toFixed(1)}% (平均${avgTime.toFixed(0)}ms)`);
  });

  if (report.errorAnalysis.totalErrors > 0) {
    console.log(`\n❌ エラー分析:`);
    console.log(`総エラー数: ${report.errorAnalysis.totalErrors}`);
    console.log(`カテゴリ別エラー:`, report.errorAnalysis.errorsByCategory);
    console.log(`難易度別エラー:`, report.errorAnalysis.errorsByDifficulty);
  }

  console.log(`\n⏱️ 評価時間: ${(report.metadata.totalEvaluationTime / 1000 / 60).toFixed(1)}分`);

  // 詳細な結果をJSONファイルに保存
  const timestamp = Date.now();
  const resultsPath = path.join(process.cwd(), 'evaluation', `large-scale-results-${timestamp}.json`);
  fs.writeFileSync(resultsPath, JSON.stringify(report, null, 2));
  console.log(`\n💾 詳細な結果を保存しました: ${resultsPath}`);

  // CSVレポートも生成
  const csvContent = generateLargeScaleCSVReport(report.rawResults);
  const csvPath = path.join(process.cwd(), 'evaluation', `large-scale-results-${timestamp}.csv`);
  fs.writeFileSync(csvPath, csvContent);
  console.log(`📄 CSVレポートを保存しました: ${csvPath}`);

  // 評価グレードを表示
  displayEvaluationGrade(report);
}

// CSV形式のレポートを生成
function generateLargeScaleCSVReport(results) {
  const headers = [
    'ID', 'Claim', 'Ground Truth', 'Prediction', 'Correct', 
    'Response Time (ms)', 'Factuality', 'Category', 'Difficulty', 'Language', 'Error'
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
    result.language,
    result.error || ''
  ]);

  return [headers.join(','), ...rows.map(row => row.join(','))].join('\n');
}

// 評価グレードを表示
function displayEvaluationGrade(report) {
  const accuracy = report.summary.accuracy;
  const errorRate = report.summary.errorRate;
  const avgResponseTime = report.performance.avgResponseTime;

  let grade = 'D';
  let assessment = '改善が必要';

  if (accuracy >= 0.95 && errorRate <= 0.05 && avgResponseTime <= 5000) {
    grade = 'A+';
    assessment = '優秀 - 商用利用可能レベル';
  } else if (accuracy >= 0.90 && errorRate <= 0.10 && avgResponseTime <= 7000) {
    grade = 'A';
    assessment = '良好 - 実用レベル';
  } else if (accuracy >= 0.85 && errorRate <= 0.15 && avgResponseTime <= 10000) {
    grade = 'B';
    assessment = '標準的 - 改良の余地あり';
  } else if (accuracy >= 0.75 && errorRate <= 0.25) {
    grade = 'C';
    assessment = '基本的 - 大幅改良が必要';
  }

  console.log(`\n${"=".repeat(80)}`);
  console.log(`🏆 総合評価: ${grade} - ${assessment}`);
  console.log(`${"=".repeat(80)}`);
}

// エラーハンドリング付きでメイン関数を実行
if (import.meta.url === `file://${process.argv[1]}`) {
  runLargeScaleEvaluation().catch(error => {
    console.error('❌ 大規模評価中にエラーが発生しました:', error);
    process.exit(1);
  });
}

export { runLargeScaleEvaluation, LargeScaleEvaluationResults };