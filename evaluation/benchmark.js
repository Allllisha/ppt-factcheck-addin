/**
 * JINA AI ファクトチェック ベンチマーク評価
 * 他のファクトチェッカーとの比較用
 */

import fs from 'fs';

// 既知のファクトチェック結果との比較データセット
const benchmarkCases = [
  {
    claim: "The capital of Japan is Tokyo.",
    jina_expected: true,
    snopes_result: true,
    politifact_result: true,
    difficulty: "easy"
  },
  {
    claim: "Humans evolved from monkeys.",
    jina_expected: false, // 正確には共通の祖先から進化
    snopes_result: false,
    politifact_result: false,
    difficulty: "medium"
  },
  {
    claim: "5G networks spread COVID-19.",
    jina_expected: false,
    snopes_result: false,
    politifact_result: false,
    difficulty: "easy"
  },
  {
    claim: "Climate change is primarily caused by human activities.",
    jina_expected: true,
    snopes_result: true,
    politifact_result: true,
    difficulty: "medium"
  },
  {
    claim: "Lightning never strikes the same place twice.",
    jina_expected: false,
    snopes_result: false,
    politifact_result: false,
    difficulty: "medium"
  }
];

// パフォーマンス指標
class PerformanceMetrics {
  constructor() {
    this.responseTimes = [];
    this.accuracyByDifficulty = {};
    this.agreements = {
      with_snopes: 0,
      with_politifact: 0,
      total_comparisons: 0
    };
  }

  addResult(testCase, jinaResult, responseTime) {
    this.responseTimes.push(responseTime);
    
    // 難易度別精度
    if (!this.accuracyByDifficulty[testCase.difficulty]) {
      this.accuracyByDifficulty[testCase.difficulty] = { correct: 0, total: 0 };
    }
    
    this.accuracyByDifficulty[testCase.difficulty].total++;
    if (jinaResult === testCase.jina_expected) {
      this.accuracyByDifficulty[testCase.difficulty].correct++;
    }

    // 他のファクトチェッカーとの一致度
    if (jinaResult === testCase.snopes_result) {
      this.agreements.with_snopes++;
    }
    if (jinaResult === testCase.politifact_result) {
      this.agreements.with_politifact++;
    }
    this.agreements.total_comparisons++;
  }

  getReport() {
    const avgResponseTime = this.responseTimes.reduce((a, b) => a + b, 0) / this.responseTimes.length;
    
    const difficultyAccuracy = {};
    Object.keys(this.accuracyByDifficulty).forEach(difficulty => {
      const stats = this.accuracyByDifficulty[difficulty];
      difficultyAccuracy[difficulty] = (stats.correct / stats.total * 100).toFixed(1);
    });

    return {
      performance: {
        avgResponseTime: avgResponseTime.toFixed(0),
        maxResponseTime: Math.max(...this.responseTimes),
        minResponseTime: Math.min(...this.responseTimes)
      },
      accuracy: difficultyAccuracy,
      agreements: {
        snopes: ((this.agreements.with_snopes / this.agreements.total_comparisons) * 100).toFixed(1),
        politifact: ((this.agreements.with_politifact / this.agreements.total_comparisons) * 100).toFixed(1)
      }
    };
  }
}

// レポート生成
function generateBenchmarkReport(results) {
  const timestamp = new Date().toISOString();
  
  const report = {
    metadata: {
      generated_at: timestamp,
      jina_model: "jina-chat",
      test_cases: benchmarkCases.length,
      purpose: "Benchmark comparison with established fact-checkers"
    },
    results: results.getReport(),
    recommendations: generateRecommendations(results.getReport())
  };

  return report;
}

function generateRecommendations(report) {
  const recommendations = [];
  
  // レスポンス時間の評価
  if (report.performance.avgResponseTime > 5000) {
    recommendations.push("レスポンス時間が長すぎます。タイムアウト設定の見直しをお勧めします。");
  } else if (report.performance.avgResponseTime < 2000) {
    recommendations.push("優秀なレスポンス時間です。");
  }

  // 精度の評価
  Object.entries(report.accuracy).forEach(([difficulty, accuracy]) => {
    if (parseFloat(accuracy) < 70) {
      recommendations.push(`${difficulty}レベルの精度が低いです（${accuracy}%）。追加学習が必要かもしれません。`);
    }
  });

  // 他のファクトチェッカーとの一致度
  if (parseFloat(report.agreements.snopes) < 80) {
    recommendations.push("Snopesとの一致度が低いです。アルゴリズムの調整を検討してください。");
  }

  if (parseFloat(report.agreements.politifact) < 80) {
    recommendations.push("PolitiFactとの一致度が低いです。政治的コンテンツの処理能力を向上させる必要があります。");
  }

  return recommendations;
}

export { benchmarkCases, PerformanceMetrics, generateBenchmarkReport };