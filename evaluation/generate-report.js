/**
 * 評価結果の可視化HTMLレポート生成
 */

import fs from 'fs';
import path from 'path';

function generateHTMLReport(evaluationData) {
  const html = `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>JINA AI ファクトチェック評価レポート</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            color: #333;
            background: #f5f5f5;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px 20px;
            border-radius: 12px;
            margin-bottom: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5rem;
            margin-bottom: 10px;
        }
        
        .header p {
            font-size: 1.1rem;
            opacity: 0.9;
        }
        
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .metric-card {
            background: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            text-align: center;
        }
        
        .metric-value {
            font-size: 2.5rem;
            font-weight: bold;
            color: #667eea;
            margin-bottom: 10px;
        }
        
        .metric-label {
            font-size: 0.9rem;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .chart-container {
            background: white;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            margin-bottom: 30px;
        }
        
        .chart-title {
            font-size: 1.5rem;
            margin-bottom: 20px;
            color: #333;
        }
        
        .confusion-matrix {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 15px;
            margin: 20px 0;
        }
        
        .matrix-cell {
            background: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            border: 2px solid #e1e5e9;
        }
        
        .matrix-cell.true-positive {
            border-color: #28a745;
            background: #d4edda;
        }
        
        .matrix-cell.true-negative {
            border-color: #28a745;
            background: #d4edda;
        }
        
        .matrix-cell.false-positive {
            border-color: #dc3545;
            background: #f8d7da;
        }
        
        .matrix-cell.false-negative {
            border-color: #dc3545;
            background: #f8d7da;
        }
        
        .detailed-results {
            background: white;
            border-radius: 12px;
            padding: 30px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        
        .result-item {
            padding: 15px;
            border-bottom: 1px solid #eee;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .result-item:last-child {
            border-bottom: none;
        }
        
        .result-claim {
            flex: 1;
            margin-right: 20px;
        }
        
        .result-status {
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.8rem;
            font-weight: bold;
        }
        
        .result-status.correct {
            background: #d4edda;
            color: #155724;
        }
        
        .result-status.incorrect {
            background: #f8d7da;
            color: #721c24;
        }
        
        .footer {
            text-align: center;
            padding: 20px;
            color: #666;
            font-size: 0.9rem;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>JINA AI ファクトチェック評価レポート</h1>
            <p>生成日時: ${new Date().toLocaleString('ja-JP')}</p>
        </div>
        
        <div class="metrics-grid">
            <div class="metric-card">
                <div class="metric-value">${(evaluationData.summary.accuracy * 100).toFixed(1)}%</div>
                <div class="metric-label">精度 (Accuracy)</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">${(evaluationData.summary.precision * 100).toFixed(1)}%</div>
                <div class="metric-label">適合率 (Precision)</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">${(evaluationData.summary.recall * 100).toFixed(1)}%</div>
                <div class="metric-label">再現率 (Recall)</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">${(evaluationData.summary.f1Score * 100).toFixed(1)}%</div>
                <div class="metric-label">F1スコア</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">${evaluationData.performance.avgResponseTime.toFixed(0)}ms</div>
                <div class="metric-label">平均レスポンス時間</div>
            </div>
            <div class="metric-card">
                <div class="metric-value">${evaluationData.summary.totalTests}</div>
                <div class="metric-label">総テスト数</div>
            </div>
        </div>

        <div class="chart-container">
            <h2 class="chart-title">混同行列 (Confusion Matrix)</h2>
            <div class="confusion-matrix">
                <div class="matrix-cell true-positive">
                    <h3>${evaluationData.confusionMatrix.truePositive}</h3>
                    <p>True Positive</p>
                </div>
                <div class="matrix-cell false-positive">
                    <h3>${evaluationData.confusionMatrix.falsePositive}</h3>
                    <p>False Positive</p>
                </div>
                <div class="matrix-cell false-negative">
                    <h3>${evaluationData.confusionMatrix.falseNegative}</h3>
                    <p>False Negative</p>
                </div>
                <div class="matrix-cell true-negative">
                    <h3>${evaluationData.confusionMatrix.trueNegative}</h3>
                    <p>True Negative</p>
                </div>
            </div>
        </div>

        <div class="chart-container">
            <h2 class="chart-title">カテゴリ別精度</h2>
            <canvas id="categoryChart" width="400" height="200"></canvas>
        </div>

        <div class="chart-container">
            <h2 class="chart-title">難易度別精度</h2>
            <canvas id="difficultyChart" width="400" height="200"></canvas>
        </div>

        <div class="detailed-results">
            <h2 class="chart-title">詳細結果</h2>
            ${evaluationData.rawResults.map(result => `
                <div class="result-item">
                    <div class="result-claim">
                        <strong>ID ${result.id}:</strong> ${result.claim}
                        <br><small>カテゴリ: ${result.category} | 難易度: ${result.difficulty} | 時間: ${result.responseTime}ms</small>
                    </div>
                    <div class="result-status ${result.correct ? 'correct' : 'incorrect'}">
                        ${result.correct ? '✅ 正解' : '❌ 不正解'}
                    </div>
                </div>
            `).join('')}
        </div>

        <div class="footer">
            <p>このレポートは JINA AI ファクトチェック評価システムによって自動生成されました。</p>
        </div>
    </div>

    <script>
        // カテゴリ別精度チャート
        const categoryCtx = document.getElementById('categoryChart').getContext('2d');
        const categoryData = ${JSON.stringify(evaluationData.detailedAnalysis.categoryAccuracy)};
        
        new Chart(categoryCtx, {
            type: 'bar',
            data: {
                labels: Object.keys(categoryData),
                datasets: [{
                    label: '精度 (%)',
                    data: Object.values(categoryData).map(v => (v * 100).toFixed(1)),
                    backgroundColor: 'rgba(102, 126, 234, 0.6)',
                    borderColor: 'rgba(102, 126, 234, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                scales: {
                    y: {
                        beginAtZero: true,
                        max: 100
                    }
                }
            }
        });

        // 難易度別精度チャート
        const difficultyCtx = document.getElementById('difficultyChart').getContext('2d');
        const difficultyData = ${JSON.stringify(evaluationData.detailedAnalysis.difficultyAccuracy)};
        
        new Chart(difficultyCtx, {
            type: 'doughnut',
            data: {
                labels: Object.keys(difficultyData),
                datasets: [{
                    data: Object.values(difficultyData).map(v => (v * 100).toFixed(1)),
                    backgroundColor: [
                        'rgba(255, 99, 132, 0.6)',
                        'rgba(54, 162, 235, 0.6)',
                        'rgba(255, 205, 86, 0.6)'
                    ]
                }]
            },
            options: {
                responsive: true
            }
        });
    </script>
</body>
</html>`;

  return html;
}

export { generateHTMLReport };