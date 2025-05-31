/* global Office, PowerPoint, document, fetch */

////////////////////////////////////////////////////////////////////////////////
// グローバルフラグ：PowerPoint タスクペーン上かどうかを保持
let isOfficePowerPoint = false;

////////////////////////////////////////////////////////////////////////////////
// デバッグログを画面上の #logArea に追記する関数
function logToScreen(msg) {
  const area = document.getElementById("logArea");
  if (!area) return;
  const time = new Date().toLocaleTimeString();
  const line = document.createElement("div");
  line.textContent = `[${time}] ${msg}`;
  area.appendChild(line);
  area.scrollTop = area.scrollHeight;
}

////////////////////////////////////////////////////////////////////////////////
// Office.js の初期化を待って UI を組み立てる
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    isOfficePowerPoint = true;

    // 「サイドロード中…」を隠して、メイン UI を出す
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    logToScreen("▶ [Office] PowerPoint ホスト上で動作していきます: " + info.host);

    // ボタン押下時に run() を呼ぶ
    document.getElementById("run").onclick = run;
  } else {
    logToScreen("▶ [Office] PowerPoint 以外のホストで動作しています: " + info.host);
  }
});

////////////////////////////////////////////////////////////////////////////////
// run()：FactCheck ボタン押下時に呼ばれるメイン関数
export async function run() {
  // PowerPoint のタスクペーン内で動いていないなら何もしない
  if (!isOfficePowerPoint || typeof PowerPoint === "undefined") {
    logToScreen("× [Office] PowerPoint 環境で動作していないため処理をスキップ");
    return;
  }
  
  // 結果コンテナをクリア
  const resultsContainer = document.getElementById("resultsContainer");
  if (resultsContainer) {
    resultsContainer.innerHTML = "";
  }

  try {
    await PowerPoint.run(async (context) => {
      logToScreen("▶ [Office] run() が呼び出されました");

      // ① 全スライドをロード
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // ② スライドごとに処理
      for (let slideIndex = 0; slideIndex < slides.items.length; slideIndex++) {
        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        logToScreen(`▶ スライド ${slideIndex + 1}：図形数 = ${shapes.items.length}`);

        // ③ 各図形（shape）ごとにチェック
        for (let shapeIndex = 0; shapeIndex < shapes.items.length; shapeIndex++) {
          const shp = shapes.items[shapeIndex];

          // (A) shapeType, placeholderType をログに出す（参考ログ）
          shp.load(["shapeType", "placeholderType"]);
          await context.sync();
          logToScreen(
            `  ■ 図形 ${shapeIndex + 1} の種類: shapeType=${shp.shapeType}, placeholderType=${shp.placeholderType}`
          );

          // (B) textFrame がなければスキップ
          if (!shp.textFrame) {
            logToScreen(`    ● 図形 ${shapeIndex + 1}：textFrame が存在しないためスキップ`);
            continue;
          }

          // (C) textFrame.hasText をロードして「テキストの有無」をチェック
          shp.textFrame.load("hasText");
          await context.sync();
          if (!shp.textFrame.hasText) {
            logToScreen(`    ● 図形 ${shapeIndex + 1}：テキストが空 (hasText=false)`);
            continue;
          }

          // (D) textRange.text をロードして全文を取得
          shp.textFrame.textRange.load("text");
          await context.sync();

          // ① 図形から取得した全文をログに出力
          const fullText = shp.textFrame.textRange.text || "";
          logToScreen(`    ▶ 図形 ${shapeIndex + 1}：text="${fullText}"`);

          // ② 文章を句点で分割して個別の主張に分ける
          const sentences = fullText
            .split(/。/)
            .map((s) => s.trim())
            .filter((s) => s.length > 0)
            .map((s) => s + "。"); // 句点を追加し直す
          
          logToScreen(`      ▶ 分割後 sentences[] = ${sentences.length}個の文章`);

          // 各文章が短すぎる場合（10文字未満）はスキップ
          const validSentences = sentences.filter(s => s.length >= 10);
          
          if (validSentences.length === 0) {
            logToScreen(`      ● 有効な文章がないためスキップ`);
            continue;
          }

          // ③ 各文章を個別にファクトチェック
          let hasError = false;
          let hasFalse = false;
          let hasTrue = false;
          
          for (let sentenceIndex = 0; sentenceIndex < validSentences.length; sentenceIndex++) {
            const sentence = validSentences[sentenceIndex];
            logToScreen(`      ▶ 文章 ${sentenceIndex + 1}/${validSentences.length}: "${sentence}"`);

            // ④ Jina（DeepSearch）API を呼び出す
            const jinaResult = await callJinaFactCheck(sentence);
            logToScreen(`      ▶ Jina API レスポンス: ${JSON.stringify(jinaResult)}`);
            
            // エラーが発生した場合も結果を表示
            if (!jinaResult || !jinaResult.hit) {
              logToScreen(`      ● Jina でヒットしなかったためエラー表示`);
              displayFactCheckError(slideIndex + 1, shapeIndex + 1, sentence, jinaResult?.error);
              hasError = true;
              continue;
            }
            
            // UI に結果を表示（文章番号も含める）
            const sentenceHeader = validSentences.length > 1 ? ` (文${sentenceIndex + 1}/${validSentences.length})` : "";
            displayFactCheckResultWithNumber(slideIndex + 1, shapeIndex + 1, sentence, jinaResult, sentenceHeader);
            
            // 結果を集計
            if (jinaResult.result === false) {
              hasFalse = true;
            } else if (jinaResult.result === true) {
              hasTrue = true;
            }
          }

          // ⑤ 全体の文字色を決定（1つでも誤りがあれば赤、エラーのみなら青、すべて正しければ緑）
          let fontColor;
          if (hasFalse) {
            fontColor = "FF0000"; // 赤（誤りあり）
          } else if (hasError && !hasTrue) {
            fontColor = "0000FF"; // 青（エラーのみ）
          } else if (hasTrue) {
            fontColor = "00FF00"; // 緑（すべて正しい）
          } else {
            fontColor = "0000FF"; // 青（不明）
          }

          // ⑥ textRange.font.color に文字色を設定（全テキストを変色）
          shp.textFrame.textRange.font.color = fontColor;
          logToScreen(`      ▶ 全体の文字色を変更: ${fontColor}`);

          // ※ shape 上の文字が変わったので同期する
          await context.sync();
        } // end for shapes
      } // end for slides

      logToScreen("▶ run() 処理が完了しました");
    }); // end PowerPoint.run
  } catch (error) {
    console.error("× [Office] run() 内エラー:", error);
    Office.context.ui.displayDialogAsync(
      `data:text/html,<html>
         <head><meta charset="utf-8" /></head>
         <body style="font-family:sans-serif; padding:16px;">
           <p>エラーが発生しました：${encodeURIComponent(error.message)}</p>
           <button onclick="Office.context.ui.messageParent('close')">OK</button>
         </body>
       </html>`,
      { height: 40, width: 20 }
    );
  }
}

////////////////////////////////////////////////////////////////////////////////
// displayFactCheckResultWithNumber(): 文章番号付きのファクトチェック結果を表示
function displayFactCheckResultWithNumber(slideNum, shapeNum, claim, result, sentenceHeader) {
  const resultsContainer = document.getElementById("resultsContainer");
  if (!resultsContainer) return;
  
  // 結果のカードを作成
  const resultCard = document.createElement("div");
  resultCard.style.cssText = `
    border: 1px solid #ddd;
    border-radius: 8px;
    padding: 12px;
    margin-bottom: 12px;
    background-color: #f9f9f9;
  `;
  
  // 判定結果に基づいて境界線の色を変更
  if (result.result === true) {
    resultCard.style.borderColor = "#4CAF50"; // 緑
  } else if (result.result === false) {
    resultCard.style.borderColor = "#F44336"; // 赤
  } else {
    resultCard.style.borderColor = "#2196F3"; // 青
  }
  
  // ヘッダー（スライド番号、図形番号、文章番号）
  const header = document.createElement("div");
  header.style.cssText = "font-weight: bold; margin-bottom: 8px;";
  header.textContent = `スライド ${slideNum} - 図形 ${shapeNum}${sentenceHeader}`;
  resultCard.appendChild(header);
  
  // クレーム（チェックした文章）
  const claimDiv = document.createElement("div");
  claimDiv.style.cssText = "font-style: italic; margin-bottom: 8px; color: #555;";
  claimDiv.textContent = `"${claim}"`;
  resultCard.appendChild(claimDiv);
  
  // 判定結果
  const resultDiv = document.createElement("div");
  resultDiv.style.cssText = "margin-bottom: 4px;";
  const resultIcon = result.result === true ? "✅" : result.result === false ? "❌" : "❓";
  const resultText = result.result === true ? "正しい" : result.result === false ? "誤り" : "不明";
  resultDiv.innerHTML = `<strong>判定:</strong> ${resultIcon} ${resultText}`;
  resultCard.appendChild(resultDiv);
  
  // 事実性スコア（主張が事実である確率）
  if (result.factuality !== null) {
    const factualityDiv = document.createElement("div");
    factualityDiv.style.cssText = "margin-bottom: 4px;";
    const percentage = (result.factuality * 100).toFixed(0);
    
    // factuality は「主張が事実である確率」を表す
    // 0% = 完全に誤り、100% = 完全に正しい
    let factualityText = "";
    let interpretation = "";
    let color = "";
    
    if (percentage <= 20) {
      interpretation = "確実に誤り";
      color = "#D32F2F"; // 濃い赤
    } else if (percentage <= 40) {
      interpretation = "おそらく誤り";
      color = "#F44336"; // 赤
    } else if (percentage <= 60) {
      interpretation = "不確実";
      color = "#FF9800"; // オレンジ
    } else if (percentage <= 80) {
      interpretation = "おそらく正しい";
      color = "#8BC34A"; // 薄緑
    } else {
      interpretation = "確実に正しい";
      color = "#4CAF50"; // 緑
    }
    
    factualityText = `<span style="color: ${color};">事実である確率: ${percentage}% (${interpretation})</span>`;
    factualityDiv.innerHTML = factualityText;
    resultCard.appendChild(factualityDiv);
  }
  
  // 理由
  if (result.reason) {
    const reasonDiv = document.createElement("div");
    reasonDiv.style.cssText = "margin-bottom: 8px;";
    reasonDiv.innerHTML = `<strong>理由:</strong> ${result.reason}`;
    resultCard.appendChild(reasonDiv);
  }
  
  // 参照情報
  if (result.references && result.references.length > 0) {
    const referencesDiv = document.createElement("div");
    referencesDiv.innerHTML = "<strong>参照:</strong>";
    referencesDiv.style.cssText = "margin-top: 8px;";
    
    const refList = document.createElement("ul");
    refList.style.cssText = "margin: 4px 0; padding-left: 20px;";
    
    result.references.forEach(ref => {
      const refItem = document.createElement("li");
      refItem.style.cssText = "margin-bottom: 4px; font-size: 12px;";
      
      const supportIcon = ref.isSupportive ? "✅" : "❌";
      refItem.innerHTML = `
        ${supportIcon} <a href="${ref.url}" target="_blank" style="color: #1976D2;">${ref.url}</a><br>
        <span style="color: #666; font-style: italic;">"${ref.keyQuote}"</span>
      `;
      refList.appendChild(refItem);
    });
    
    referencesDiv.appendChild(refList);
    resultCard.appendChild(referencesDiv);
  }
  
  // カードをコンテナに追加
  resultsContainer.appendChild(resultCard);
  
  // スクロールして最新の結果が見えるようにする
  resultsContainer.scrollTop = resultsContainer.scrollHeight;
}

////////////////////////////////////////////////////////////////////////////////
// displayFactCheckError(): ファクトチェックエラーをUIに表示する関数
function displayFactCheckError(slideNum, shapeNum, claim, errorMsg) {
  const resultsContainer = document.getElementById("resultsContainer");
  if (!resultsContainer) return;
  
  // エラーカードを作成
  const errorCard = document.createElement("div");
  errorCard.style.cssText = `
    border: 2px solid #FFC107;
    border-radius: 8px;
    padding: 12px;
    margin-bottom: 12px;
    background-color: #FFF8E1;
  `;
  
  // ヘッダー（スライド番号、図形番号）
  const header = document.createElement("div");
  header.style.cssText = "font-weight: bold; margin-bottom: 8px;";
  header.textContent = `スライド ${slideNum} - 図形 ${shapeNum}`;
  errorCard.appendChild(header);
  
  // クレーム（チェックした文章）
  const claimDiv = document.createElement("div");
  claimDiv.style.cssText = "font-style: italic; margin-bottom: 8px; color: #555;";
  claimDiv.textContent = `"${claim}"`;
  errorCard.appendChild(claimDiv);
  
  // エラーメッセージ
  const errorDiv = document.createElement("div");
  errorDiv.style.cssText = "color: #F57C00; margin-bottom: 4px;";
  errorDiv.innerHTML = `<strong>⚠️ ファクトチェックが実行できませんでした</strong>`;
  errorCard.appendChild(errorDiv);
  
  // エラー詳細
  if (errorMsg) {
    const detailDiv = document.createElement("div");
    detailDiv.style.cssText = "font-size: 12px; color: #666; margin-top: 4px;";
    detailDiv.textContent = `エラー: ${errorMsg}`;
    errorCard.appendChild(detailDiv);
  }
  
  // 一般的なメッセージ
  const messageDiv = document.createElement("div");
  messageDiv.style.cssText = "font-size: 12px; color: #666; margin-top: 8px;";
  messageDiv.innerHTML = `ネットワーク接続またはAPIサービスの問題により、この文章のファクトチェックを完了できませんでした。後でもう一度お試しください。`;
  errorCard.appendChild(messageDiv);
  
  // カードをコンテナに追加
  resultsContainer.appendChild(errorCard);
  
  // スクロールして最新の結果が見えるようにする
  resultsContainer.scrollTop = resultsContainer.scrollHeight;
}

////////////////////////////////////////////////////////////////////////////////
// callJinaFactCheck(): Jina DeepSearch (Grounding) API を叩くユーティリティ
async function callJinaFactCheck(claim) {
  // ■■■■ ご自身で発行した Jina トークンを必ず置き換えてください ■■■■
  const JINA_TOKEN = "jina_a7eab8ce41f9449a921a29da53c223c4c-kXGvKdkdFxOa5H9hBLrMIuxDCI";
  // ■ DeepSearch の Chat Completions エンドポイント ■
  const endpoint = "https://deepsearch.jina.ai/v1/chat/completions";

  // DeepSearch 呼び出し時のリクエストボディ（ファクトチェック用）
  const body = {
    model: "jina-chat", // DeepSearch用のチャットモデル
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
    search: true  // 検索を有効化
  };

  // ログ出力用の関数
  const debugLog = (msg) => {
    console.log(msg);
    const area = document.getElementById("logArea");
    if (area) {
      const time = new Date().toLocaleTimeString();
      const line = document.createElement("div");
      line.textContent = `[${time}] ${msg}`;
      area.appendChild(line);
      area.scrollTop = area.scrollHeight;
    }
  };

  try {
    debugLog(`▶ Jina にリクエスト：URL = ${endpoint}`);
    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${JINA_TOKEN}`
      },
      body: JSON.stringify(body)
    });
    debugLog(`▶ Jina HTTP ステータス: ${res.status}`);

    // JSON を読み取る
    let data;
    try {
      data = await res.json();
    } catch (parseErr) {
      debugLog(`× レスポンス JSON 解析エラー: ${parseErr}`);
      return { hit: false, error: "Jina レスポンス JSON 解析に失敗" };
    }

    if (!res.ok) {
      const errMsg = (data && data.error && data.error.message) 
        ? data.error.message 
        : `ステータスコード ${res.status}`;
      debugLog(`× Jina エラー詳細: ${JSON.stringify(data)}`);
      return { hit: false, error: errMsg };
    }

    debugLog(`▶ Jina レスポンスボディ: ${JSON.stringify(data)}`);
    
    // Response format check - handle both wrapper format and direct format
    let responseData;
    
    // Check if response has the wrapper format with code/status/data
    if (data.code === 200 && data.status === 20000 && data.data) {
      responseData = data.data;
      debugLog(`▶ Wrapper format detected, extracting data: ${JSON.stringify(responseData)}`);
    } 
    // Check if it's the chat completions format
    else if (data.choices && data.choices[0]) {
      const choice = data.choices[0];
      if (choice.message && choice.message.content) {
        // Try to parse content if it's a JSON string
        try {
          if (typeof choice.message.content === 'string') {
            // Remove markdown code block if present
            let content = choice.message.content;
            content = content.replace(/^```json\s*\n?/, '').replace(/\n?```\s*$/, '');
            responseData = JSON.parse(content);
          } else {
            responseData = choice.message.content;
          }
        } catch (e) {
          // If parsing fails, use content as is
          responseData = choice.message.content;
        }
        debugLog(`▶ Chat completions format detected: ${JSON.stringify(responseData)}`);
      }
    }
    // Direct format (factuality, result, reason at top level)
    else if (data.factuality !== undefined || data.result !== undefined) {
      responseData = data;
      debugLog(`▶ Direct format detected: ${JSON.stringify(responseData)}`);
    }
    
    if (!responseData) {
      debugLog(`× Jina レスポンスに有効なデータが含まれない`);
      return { hit: false };
    }

    // 必要なフィールドを取り出し、自作の構造にマッピングする
    //   - hit: true / false
    //   - result: true/false/その他説明文
    //   - reason: 理由説明
    //   - factuality: 0.00～1.00 信頼度（存在すれば）
    //   - references: 根拠ドキュメントリスト（存在すればそのまま）
    return {
      hit: true,
      result: responseData.result ?? "",           // true か false
      reason: responseData.reason ?? "",
      factuality: responseData.factuality ?? null, // 数値スコア
      references: responseData.references ?? []     // 根拠リスト
    };
  } catch (e) {
    debugLog(`× [Office] Jina 呼び出し中に例外発生: ${e}`);
    return { hit: false, error: e.message || String(e) };
  }
}