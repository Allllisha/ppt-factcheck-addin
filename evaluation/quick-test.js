/**
 * JINA AI ファクトチェック クイックテスト
 * 少数のテストケースで素早く性能を確認
 */

import { API_CONFIG } from '../src/config.js';

// クイックテスト用のデータセット
const quickTestCases = [
  {
    id: 1,
    claim: "Water boils at 100 degrees Celsius at sea level.",
    expected: true,
    category: "science"
  },
  {
    id: 2,
    claim: "The Earth is flat.",
    expected: false,
    category: "science"
  },
  {
    id: 3,
    claim: "COVID-19 vaccines contain microchips.",
    expected: false,
    category: "health"
  },
  {
    id: 4,
    claim: "水は海面で摂氏100度で沸騰する。",
    expected: true,
    category: "science_japanese"
  },
  {
    id: 5,
    claim: "地球は平らである。",
    expected: false,
    category: "science_japanese"
  }
];

// JINA APIを呼び出す簡略化された関数
async function quickFactCheck(claim) {
  const JINA_TOKEN = API_CONFIG.JINA_API_TOKEN;
  const endpoint = "https://deepsearch.jina.ai/v1/chat/completions";

  const body = {
    model: "jina-chat",
    messages: [{
      role: "user",
      content: `Please fact-check this claim and respond with only "true" or "false": "${claim}"`
    }],
    stream: false,
    temperature: 0,
    search: true
  };

  const startTime = Date.now();

  try {
    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${JINA_TOKEN}`
      },
      body: JSON.stringify(body)
    });

    const responseTime = Date.now() - startTime;

    if (!res.ok) {
      return { error: `HTTP ${res.status}`, responseTime };
    }

    const data = await res.json();
    let content = data.choices?.[0]?.message?.content?.toLowerCase() || "";
    
    // 結果を解析
    let result = null;
    if (content.includes("true") && !content.includes("false")) {
      result = true;
    } else if (content.includes("false") && !content.includes("true")) {
      result = false;
    }

    return { result, content, responseTime };
  } catch (e) {
    return { error: e.message, responseTime: Date.now() - startTime };
  }
}

// クイックテストを実行
async function runQuickTest() {
  console.log("🚀 JINA AI ファクトチェック クイックテスト開始\n");
  
  let correct = 0;
  let total = quickTestCases.length;
  let totalTime = 0;

  for (const testCase of quickTestCases) {
    console.log(`\n🔍 テスト ${testCase.id}: ${testCase.category}`);
    console.log(`Claims: "${testCase.claim}"`);
    console.log(`期待値: ${testCase.expected}`);
    
    const result = await quickFactCheck(testCase.claim);
    
    if (result.error) {
      console.log(`❌ エラー: ${result.error}`);
    } else {
      console.log(`結果: ${result.result}`);
      console.log(`レスポンス: "${result.content}"`);
      console.log(`時間: ${result.responseTime}ms`);
      
      const isCorrect = result.result === testCase.expected;
      console.log(`判定: ${isCorrect ? '✅ 正解' : '❌ 不正解'}`);
      
      if (isCorrect) correct++;
      totalTime += result.responseTime;
    }
    
    // レート制限対策
    await new Promise(resolve => setTimeout(resolve, 1000));
  }

  console.log("\n" + "=".repeat(50));
  console.log("📊 テスト結果サマリー");
  console.log("=".repeat(50));
  console.log(`総テスト数: ${total}`);
  console.log(`正解数: ${correct}`);
  console.log(`精度: ${((correct / total) * 100).toFixed(1)}%`);
  console.log(`平均レスポンス時間: ${(totalTime / total).toFixed(0)}ms`);
  
  if (correct / total >= 0.8) {
    console.log("🎉 良好なパフォーマンスです！");
  } else if (correct / total >= 0.6) {
    console.log("⚠️  改善の余地があります");
  } else {
    console.log("🔴 パフォーマンスが低いです");
  }
}

// メイン実行
if (import.meta.url === `file://${process.argv[1]}`) {
  runQuickTest().catch(console.error);
}