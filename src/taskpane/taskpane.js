/* global Office, PowerPoint, document, fetch */

// API設定をインポート
import { API_CONFIG } from '../config.js';

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
// ローディング状態を表示する関数
function showLoadingState() {
  const loadingState = document.getElementById("loadingState");
  const resultsContainer = document.getElementById("resultsContainer");
  
  if (loadingState && resultsContainer) {
    // 既存の結果をクリア
    resultsContainer.innerHTML = "";
    
    // ローディング状態を表示
    loadingState.style.display = "flex";
    
    // プログレスバーを0%にリセット
    const progressFill = document.getElementById("progressFill");
    const progressText = document.getElementById("progressText");
    if (progressFill) progressFill.style.width = "0%";
    if (progressText) progressText.textContent = "準備中...";
  }
}

////////////////////////////////////////////////////////////////////////////////
// ローディング状態を隠す関数
function hideLoadingState() {
  const loadingState = document.getElementById("loadingState");
  if (loadingState) {
    loadingState.style.display = "none";
  }
}

////////////////////////////////////////////////////////////////////////////////
// プログレスバーを更新する関数
function updateProgress(percentage, message) {
  const progressFill = document.getElementById("progressFill");
  const progressText = document.getElementById("progressText");
  
  if (progressFill) {
    progressFill.style.width = `${Math.min(100, Math.max(0, percentage))}%`;
  }
  
  if (progressText && message) {
    progressText.textContent = message;
  }
}

////////////////////////////////////////////////////////////////////////////////
// 個別のファクトチェックプログレスカードを追加する関数
function addFactCheckProgressCard(slideNum, shapeNum, sentence, sentenceHeader = "") {
  const resultsContainer = document.getElementById("resultsContainer");
  if (!resultsContainer) return null;
  
  // プログレスカードを作成
  const progressCard = document.createElement("div");
  progressCard.className = "fact-check-progress-card";
  progressCard.style.cssText = `
    border: none;
    border-radius: 16px;
    padding: 20px;
    margin-bottom: 16px;
    background: linear-gradient(135deg, #F0F9FF 0%, #E0F2FE 100%);
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08), 0 2px 8px rgba(0, 0, 0, 0.04);
    position: relative;
    overflow: hidden;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    border-left: 4px solid #3B82F6;
    animation: slideInUp 0.3s ease-out;
  `;
  
  // ヘッダー
  const header = document.createElement("div");
  header.style.cssText = `
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 16px;
    padding-bottom: 12px;
    border-bottom: 1px solid rgba(0, 0, 0, 0.06);
  `;
  
  const locationInfo = document.createElement("span");
  locationInfo.style.cssText = `
    font-size: 12px;
    font-weight: 600;
    color: #1E40AF;
    background: rgba(59, 130, 246, 0.1);
    padding: 4px 8px;
    border-radius: 6px;
  `;
  locationInfo.textContent = `スライド ${slideNum} - テキストボックス ${shapeNum}${sentenceHeader}`;
  
  const statusBadge = document.createElement("span");
  statusBadge.style.cssText = `
    font-size: 14px;
    font-weight: 600;
    padding: 6px 12px;
    border-radius: 20px;
    background: #3B82F6;
    color: white;
    display: flex;
    align-items: center;
    gap: 6px;
  `;
  statusBadge.innerHTML = `<div class="mini-spinner"></div> 分析中`;
  
  header.appendChild(locationInfo);
  header.appendChild(statusBadge);
  progressCard.appendChild(header);
  
  // 文章表示
  const claimDiv = document.createElement("div");
  claimDiv.style.cssText = `
    font-size: 16px;
    line-height: 1.6;
    margin-bottom: 16px;
    color: #1F2937;
    padding: 16px;
    background: rgba(255, 255, 255, 0.6);
    border-radius: 12px;
    border-left: 3px solid #3B82F6;
    font-weight: 500;
  `;
  claimDiv.textContent = `"${sentence}"`;
  progressCard.appendChild(claimDiv);
  
  // プログレス表示
  const progressDiv = document.createElement("div");
  progressDiv.style.cssText = `
    display: flex;
    align-items: center;
    gap: 12px;
    color: #1E40AF;
    font-size: 14px;
    font-weight: 500;
  `;
  progressDiv.innerHTML = `
    <div class="mini-spinner"></div>
    <span>AIがファクトチェックを実行中...</span>
  `;
  progressCard.appendChild(progressDiv);
  
  // ミニスピナーのスタイルを追加（まだ追加されていない場合）
  if (!document.getElementById('mini-spinner-style')) {
    const style = document.createElement('style');
    style.id = 'mini-spinner-style';
    style.textContent = `
      .mini-spinner {
        width: 16px;
        height: 16px;
        border: 2px solid rgba(59, 130, 246, 0.3);
        border-top: 2px solid #3B82F6;
        border-radius: 50%;
        animation: spin 1s linear infinite;
        display: inline-block;
      }
    `;
    document.head.appendChild(style);
  }
  
  // カードをコンテナに追加
  resultsContainer.appendChild(progressCard);
  
  // 最初のプログレスカードの場合は、ローディングを隠して結果セクションまでスクロール
  if (resultsContainer.children.length === 1) {
    hideLoadingState();
    
    const resultsSection = document.querySelector(".results-section");
    if (resultsSection) {
      resultsSection.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }
  
  // スクロールして最新のカードが見えるようにする
  resultsContainer.scrollTop = resultsContainer.scrollHeight;
  
  return progressCard;
}

////////////////////////////////////////////////////////////////////////////////
// プログレスカードを結果カードに置き換える関数
function replaceProgressCardWithResult(progressCard, slideNum, shapeNum, claim, result, sentenceHeader) {
  if (!progressCard || !progressCard.parentNode) return;
  
  const resultsContainer = progressCard.parentNode;
  
  // フェードアウトアニメーション
  progressCard.style.transition = "opacity 0.3s ease-out, transform 0.3s ease-out";
  progressCard.style.opacity = "0";
  progressCard.style.transform = "translateX(-20px)";
  
  setTimeout(() => {
    // プログレスカードの位置を記録
    const progressCardIndex = Array.from(resultsContainer.children).indexOf(progressCard);
    
    // プログレスカードを削除
    progressCard.remove();
    
    // 結果カードを作成（一時的にコンテナの外で）
    const tempContainer = document.createElement('div');
    const originalContainer = resultsContainer;
    
    // displayFactCheckResultWithNumber関数が結果カードを作成できるように、一時的にIDを変更
    tempContainer.id = 'resultsContainer';
    document.body.appendChild(tempContainer);
    
    // 元のコンテナのIDを削除
    originalContainer.id = 'tempResultsContainer';
    
    // 結果カードを作成
    displayFactCheckResultWithNumber(slideNum, shapeNum, claim, result, sentenceHeader);
    
    // 作成された結果カードを取得
    const newResultCard = tempContainer.firstElementChild;
    
    if (newResultCard) {
      // IDを元に戻す
      originalContainer.id = 'resultsContainer';
      tempContainer.remove();
      
      // 結果カードを正しい位置に挿入
      if (progressCardIndex < originalContainer.children.length) {
        originalContainer.insertBefore(newResultCard, originalContainer.children[progressCardIndex]);
      } else {
        originalContainer.appendChild(newResultCard);
      }
      
      // フェードインアニメーション
      newResultCard.style.opacity = "0";
      newResultCard.style.transform = "translateX(20px)";
      newResultCard.style.transition = "opacity 0.3s ease-out, transform 0.3s ease-out";
      
      setTimeout(() => {
        newResultCard.style.opacity = "1";
        newResultCard.style.transform = "translateX(0)";
      }, 50);
    } else {
      // フォールバック: IDを元に戻す
      originalContainer.id = 'resultsContainer';
      tempContainer.remove();
    }
  }, 300);
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
  console.log("[DEBUG] run() function called");
  logToScreen("[DEBUG] run() function called");
  
  // PowerPoint のタスクペーン内で動いていないなら何もしない
  if (!isOfficePowerPoint || typeof PowerPoint === "undefined") {
    logToScreen("× [Office] PowerPoint 環境で動作していないため処理をスキップ");
    return;
  }
  
  // アクションセクションを自動的に閉じる
  const actionSection = document.querySelector(".action-section");
  if (actionSection && actionSection.hasAttribute("open")) {
    actionSection.removeAttribute("open");
  }
  
  // ローディング状態を表示
  showLoadingState();
  updateProgress(5, "スライドを読み込み中...");

  // ファクトチェック結果を収集する配列
  const allFactCheckResults = [];

  try {
    await PowerPoint.run(async (context) => {
      logToScreen("▶ [Office] run() が呼び出されました");

      // ① 全スライドをロード
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      
      updateProgress(15, `${slides.items.length}枚のスライドを発見しました`);
      logToScreen(`▶ 全スライド数: ${slides.items.length}`);

      // ② スライドごとに処理
      for (let slideIndex = 0; slideIndex < slides.items.length; slideIndex++) {
        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();
        
        // スライド進捗を更新（20% - 80%の範囲）
        const slideProgress = 20 + (slideIndex / slides.items.length) * 60;
        updateProgress(slideProgress, `スライド ${slideIndex + 1}/${slides.items.length} を分析中...`);

        logToScreen(`▶ スライド ${slideIndex + 1}：テキストボックス数 = ${shapes.items.length}`);

        // ③ 各テキストボックス（shape）ごとにチェック
        for (let shapeIndex = 0; shapeIndex < shapes.items.length; shapeIndex++) {
          const shp = shapes.items[shapeIndex];

          // (A) shapeType, placeholderType をログに出す（参考ログ）
          shp.load(["shapeType", "placeholderType"]);
          await context.sync();
          logToScreen(
            `  ■ テキストボックス ${shapeIndex + 1} の種類: shapeType=${shp.shapeType}, placeholderType=${shp.placeholderType}`
          );

          // (B) textFrame がなければスキップ
          if (!shp.textFrame) {
            logToScreen(`    ● テキストボックス ${shapeIndex + 1}：textFrame が存在しないためスキップ`);
            continue;
          }

          // (C) textFrame.hasText をロードして「テキストの有無」をチェック
          shp.textFrame.load("hasText");
          await context.sync();
          if (!shp.textFrame.hasText) {
            logToScreen(`    ● テキストボックス ${shapeIndex + 1}：テキストが空 (hasText=false)`);
            continue;
          }

          // (D) textRange.text をロードして全文を取得
          shp.textFrame.textRange.load("text");
          await context.sync();

          // ① テキストボックスから取得した全文をログに出力
          const fullText = shp.textFrame.textRange.text || "";
          logToScreen(`    ▶ テキストボックス ${shapeIndex + 1}：text="${fullText}"`);

          // ② 文章を句点で分割して個別の主張に分ける
          // 日本語と英語の両方に対応した文分割
          let sentences = [];
          
          // 日本語の句点が含まれているかチェック
          if (fullText.includes("。")) {
            // 日本語テキストの場合
            sentences = fullText
              .split(/。/)
              .map((s) => s.trim())
              .filter((s) => s.length > 0)
              .map((s, index, array) => {
                // 最後の文章以外、または元のテキストが句点で終わる場合は句点を追加
                if (index < array.length - 1 || fullText.trim().endsWith("。")) {
                  return s + "。";
                }
                return s; // 最後の文章で元のテキストが句点で終わらない場合はそのまま
              });
          } else {
            // 英語テキストの場合
            // より高度な英語文分割ロジック
            // 略語 (Dr., Mr., Mrs., etc.) を考慮
            const englishText = fullText
              // 略語の後のピリオドを一時的に置換
              .replace(/\b(Dr|Mr|Mrs|Ms|Prof|Sr|Jr)\./g, '$1__ABBR__')
              // 小数点を一時的に置換
              .replace(/(\d)\.(\d)/g, '$1__DECIMAL__$2')
              // その他の一般的な略語
              .replace(/\b(vs|etc|Inc|Ltd|Co|Corp|e\.g|i\.e|cf|al)\./g, '$1__ABBR__');
            
            // ピリオド、感嘆符、疑問符で分割
            const rawSentences = englishText.split(/([.!?]+)\s+/);
            
            // 文と終止符を結合
            for (let i = 0; i < rawSentences.length; i += 2) {
              if (i + 1 < rawSentences.length) {
                const sentence = (rawSentences[i] + rawSentences[i + 1])
                  // 置換したものを元に戻す
                  .replace(/__ABBR__/g, '.')
                  .replace(/__DECIMAL__/g, '.')
                  .trim();
                if (sentence.length > 0) {
                  sentences.push(sentence);
                }
              } else {
                // 最後の要素（終止符がない場合）
                const sentence = rawSentences[i]
                  .replace(/__ABBR__/g, '.')
                  .replace(/__DECIMAL__/g, '.')
                  .trim();
                if (sentence.length > 0) {
                  sentences.push(sentence);
                }
              }
            }
            
            // 文末に終止符がない場合の処理
            if (sentences.length === 0 && fullText.trim().length > 0) {
              sentences = [fullText.trim()];
            }
          }
          
          logToScreen(`      ▶ 分割後 sentences[] = ${sentences.length}個の文章`);
          
          // デバッグ用: 分割された文章を表示
          sentences.forEach((s, idx) => {
            logToScreen(`        - 文章${idx + 1}: "${s}"`);
          });

          // 全ての文章をファクトチェック対象にする（長さ制限を削除）
          const validSentences = sentences.filter(s => s.length > 0);
          
          if (validSentences.length === 0) {
            logToScreen(`      ● 有効な文章がないためスキップ`);
            
            // スキップした場合も結果として表示
            displayFactCheckError(slideIndex + 1, shapeIndex + 1, fullText, "有効な文章がありません");
            
            continue;
          }

          // ③ 各文章を個別にファクトチェック
          let hasError = false;
          let hasFalse = false;
          let hasTrue = false;
          
          // 各文章の結果を保存する配列
          const sentenceResults = [];
          
          for (let sentenceIndex = 0; sentenceIndex < validSentences.length; sentenceIndex++) {
            const sentence = validSentences[sentenceIndex];
            logToScreen(`      ▶ 文章 ${sentenceIndex + 1}/${validSentences.length}: "${sentence}"`);
            
            // 文章レベルの進捗を更新
            updateProgress(slideProgress + (sentenceIndex / validSentences.length) * 10, 
              `「${sentence.substring(0, 30)}...」をファクトチェック中`);

            // 日本語か英語かに応じて表示を変更
            const isJapanese = fullText.includes("。");
            const sentenceHeader = validSentences.length > 1 
              ? isJapanese 
                ? ` (文${sentenceIndex + 1}/${validSentences.length})` 
                : ` (Sentence ${sentenceIndex + 1}/${validSentences.length})`
              : "";

            // プログレスカードを表示
            const progressCard = addFactCheckProgressCard(slideIndex + 1, shapeIndex + 1, sentence, sentenceHeader);

            // ④ Jina（DeepSearch）API を呼び出し
            const factCheckResult = await callJinaFactCheck(sentence);
            logToScreen(`      ▶ ファクトチェック結果: ${JSON.stringify(factCheckResult)}`);
            
            // エラーが発生した場合も結果を表示
            if (!factCheckResult || !factCheckResult.hit) {
              logToScreen(`      ● ファクトチェックでヒットしなかったためエラー表示`);
              
              // プログレスカードを削除してエラーカードを表示
              if (progressCard) {
                progressCard.remove();
              }
              displayFactCheckError(slideIndex + 1, shapeIndex + 1, sentence, factCheckResult?.error);
              hasError = true;
              sentenceResults.push({ sentence, result: "no_check" }); // エラーと区別
              
              // レポート用にエラー結果も保存
              allFactCheckResults.push({
                slideNumber: slideIndex + 1,
                shapeNumber: shapeIndex + 1,
                sentence: sentence,
                result: "error",
                reason: factCheckResult?.error || "ファクトチェックに失敗しました",
                factuality: null,
                references: []
              });
              continue;
            }
            
            // プログレスカードを結果カードに置き換え
            replaceProgressCardWithResult(progressCard, slideIndex + 1, shapeIndex + 1, sentence, factCheckResult, sentenceHeader);
            
            // レポート用に結果を保存
            allFactCheckResults.push({
              slideNumber: slideIndex + 1,
              shapeNumber: shapeIndex + 1,
              sentence: sentence,
              result: factCheckResult.result,
              reason: factCheckResult.reason,
              factuality: factCheckResult.factuality,
              references: factCheckResult.references || []
            });
            
            // 結果を集計
            if (factCheckResult.result === false) {
              hasFalse = true;
              sentenceResults.push({ sentence, result: false });
            } else if (factCheckResult.result === true) {
              hasTrue = true;
              sentenceResults.push({ sentence, result: true });
            } else {
              sentenceResults.push({ sentence, result: "unknown" });
            }
          }

          // ⑤ 各文章を個別に色分けする
          // まず全文を取得し、各文章の位置を特定
          const fullTextForColoring = shp.textFrame.textRange.text;
          let currentPosition = 0;
          
          logToScreen(`      ▶ 色分け処理開始: fullText="${fullTextForColoring}"`);
          
          for (const sentenceResult of sentenceResults) {
            // 文章の開始位置を検索
            logToScreen(`      ▶ 検索中: "${sentenceResult.sentence}" (位置${currentPosition}から)`);
            const sentenceStart = fullTextForColoring.indexOf(sentenceResult.sentence, currentPosition);
            if (sentenceStart === -1) {
              logToScreen(`      ● 文章 "${sentenceResult.sentence}" がテキスト内で見つかりませんでした`);
              logToScreen(`        現在のテキスト: "${fullTextForColoring.substring(currentPosition)}"`);
              continue;
            }
            
            // 文章の終了位置
            const sentenceEnd = sentenceStart + sentenceResult.sentence.length;
            
            // 色を決定
            let fontColor;
            if (sentenceResult.result === false) {
              fontColor = "FF0000"; // 赤（誤り）
            } else if (sentenceResult.result === true) {
              fontColor = "00FF00"; // 緑（正しい）
            } else if (sentenceResult.result === "no_check") {
              fontColor = null; // ファクトチェック不可の場合は色変更しない（黒のまま）
            } else if (sentenceResult.result === "error") {
              fontColor = "0000FF"; // 青（エラー）
            } else {
              fontColor = "0000FF"; // 青（不明）
            }
            
            // 特定の範囲の文字色を変更
            if (fontColor !== null) {
              try {
                // getSubstring で部分文字列を取得し、その色を変更
                const subTextRange = shp.textFrame.textRange.getSubstring(sentenceStart, sentenceEnd - sentenceStart);
                subTextRange.font.color = fontColor;
                logToScreen(`      ▶ 文章 ${sentenceResults.indexOf(sentenceResult) + 1} の色を変更: ${fontColor} (位置: ${sentenceStart}-${sentenceEnd})`);
              } catch (e) {
                logToScreen(`      ● 文章の色変更エラー: ${e.message}`);
              }
            } else {
              logToScreen(`      ▶ 文章 ${sentenceResults.indexOf(sentenceResult) + 1} はファクトチェック不可のため色変更をスキップ`);
            }
            
            // 次の検索開始位置を更新
            currentPosition = sentenceEnd;
          }

          // ※ テキストボックス上の文字が変わったので同期する
          await context.sync();
        } // end for shapes
      } // end for slides

      // ⑥ ファクトチェックレポートスライドを作成（無効化）
      // if (allFactCheckResults.length > 0) {
      //   logToScreen("▶ ファクトチェックレポートを作成中...");
      //   await createFactCheckReportSlide(context, allFactCheckResults);
      //   logToScreen("▶ ファクトチェックレポートを追加しました");
      // }

      // 処理完了
      updateProgress(100, "ファクトチェック完了！");
      setTimeout(() => {
        hideLoadingState();
      }, 1000); // 1秒後にローディングを隠す
      
      logToScreen("▶ run() 処理が完了しました");
    }); // end PowerPoint.run
  } catch (error) {
    // エラー時にもローディングを隠す
    hideLoadingState();
    
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
  
  // 結果のカードを作成（モダンデザイン）
  const resultCard = document.createElement("div");
  resultCard.style.cssText = `
    border: none;
    border-radius: 16px;
    padding: 20px;
    margin-bottom: 16px;
    background: linear-gradient(135deg, #ffffff 0%, #f8fafb 100%);
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08), 0 2px 8px rgba(0, 0, 0, 0.04);
    position: relative;
    overflow: hidden;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    border-left: 4px solid transparent;
  `;
  
  // 判定結果に基づいてアクセントカラーを変更
  let accentColor, backgroundGradient, statusIcon;
  if (result.result === true) {
    accentColor = "#10B981"; // モダンな緑
    backgroundGradient = "linear-gradient(135deg, #ECFDF5 0%, #F0FDF4 100%)";
    statusIcon = "✅";
  } else if (result.result === false) {
    accentColor = "#EF4444"; // モダンな赤
    backgroundGradient = "linear-gradient(135deg, #FEF2F2 0%, #FEF7F7 100%)";
    statusIcon = "❌";
  } else {
    accentColor = "#3B82F6"; // モダンな青
    backgroundGradient = "linear-gradient(135deg, #EFF6FF 0%, #F0F9FF 100%)";
    statusIcon = "❓";
  }
  
  resultCard.style.borderLeftColor = accentColor;
  resultCard.style.background = backgroundGradient;
  
  // ホバー効果
  resultCard.onmouseenter = () => {
    resultCard.style.transform = "translateY(-2px)";
    resultCard.style.boxShadow = "0 8px 30px rgba(0, 0, 0, 0.12), 0 4px 12px rgba(0, 0, 0, 0.08)";
  };
  resultCard.onmouseleave = () => {
    resultCard.style.transform = "translateY(0)";
    resultCard.style.boxShadow = "0 4px 20px rgba(0, 0, 0, 0.08), 0 2px 8px rgba(0, 0, 0, 0.04)";
  };
  
  // ヘッダー（スライド番号、テキストボックス番号、文章番号）
  const header = document.createElement("div");
  header.style.cssText = `
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 16px;
    padding-bottom: 12px;
    border-bottom: 1px solid rgba(0, 0, 0, 0.06);
  `;
  
  const locationInfo = document.createElement("span");
  locationInfo.style.cssText = `
    font-size: 12px;
    font-weight: 600;
    color: #6B7280;
    background: rgba(107, 114, 128, 0.1);
    padding: 4px 8px;
    border-radius: 6px;
  `;
  const sourceInfo = result.source ? ` [${result.source}]` : "";
  locationInfo.textContent = `スライド ${slideNum} - テキストボックス ${shapeNum}${sentenceHeader}${sourceInfo}`;
  
  const statusBadge = document.createElement("span");
  statusBadge.style.cssText = `
    font-size: 14px;
    font-weight: 600;
    padding: 6px 12px;
    border-radius: 20px;
    background: ${accentColor};
    color: white;
    display: flex;
    align-items: center;
    gap: 6px;
  `;
  const resultText = result.result === true ? "正しい" : result.result === false ? "誤り" : "不明";
  statusBadge.innerHTML = `${statusIcon} ${resultText}`;
  
  header.appendChild(locationInfo);
  header.appendChild(statusBadge);
  resultCard.appendChild(header);
  
  // クレーム（チェックした文章）
  const claimDiv = document.createElement("div");
  claimDiv.style.cssText = `
    font-size: 16px;
    line-height: 1.6;
    margin-bottom: 16px;
    color: #1F2937;
    padding: 16px;
    background: rgba(255, 255, 255, 0.6);
    border-radius: 12px;
    border-left: 3px solid ${accentColor};
    font-weight: 500;
  `;
  claimDiv.textContent = `"${claim}"`;
  resultCard.appendChild(claimDiv);
  
  // 事実性スコア（主張が事実である確率）
  if (result.factuality !== null) {
    const factualityDiv = document.createElement("div");
    factualityDiv.style.cssText = "margin-bottom: 4px;";
    
    // factuality を result に基づいて調整
    // result が false の場合、factuality は信頼度を表すので、事実である確率は (1 - factuality)
    let adjustedFactuality;
    if (result.result === false) {
      adjustedFactuality = 1 - result.factuality;
    } else if (result.result === true) {
      adjustedFactuality = result.factuality;
    } else {
      adjustedFactuality = 0.5; // 不明な場合は50%
    }
    
    const percentage = (adjustedFactuality * 100).toFixed(0);
    
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
  
  // 修正ボタンを追加（誤りの場合のみ）
  if (result.result === false) {
    const correctButtonDiv = document.createElement("div");
    correctButtonDiv.style.cssText = "margin-top: 16px; text-align: right;";
    
    const correctButton = document.createElement("button");
    correctButton.innerHTML = "🔧 修正する";
    correctButton.style.cssText = `
      background: linear-gradient(135deg, #FF9800 0%, #F57C00 100%);
      color: white;
      border: none;
      padding: 12px 20px;
      border-radius: 12px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      box-shadow: 0 2px 8px rgba(255, 152, 0, 0.3);
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      display: inline-flex;
      align-items: center;
      gap: 6px;
    `;
    
    // ホバー効果
    correctButton.onmouseenter = () => {
      correctButton.style.transform = "translateY(-1px)";
      correctButton.style.boxShadow = "0 4px 16px rgba(255, 152, 0, 0.4)";
      correctButton.style.background = "linear-gradient(135deg, #F57C00 0%, #E65100 100%)";
    };
    correctButton.onmouseleave = () => {
      correctButton.style.transform = "translateY(0)";
      correctButton.style.boxShadow = "0 2px 8px rgba(255, 152, 0, 0.3)";
      correctButton.style.background = "linear-gradient(135deg, #FF9800 0%, #F57C00 100%)";
    };
    
    // クリックイベント
    correctButton.onclick = () => {
      correctFactCheckResult(slideNum, claim, result);
    };
    
    correctButtonDiv.appendChild(correctButton);
    resultCard.appendChild(correctButtonDiv);
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
  
  // エラーカードを作成（モダンデザイン）
  const errorCard = document.createElement("div");
  errorCard.style.cssText = `
    border: none;
    border-radius: 16px;
    padding: 20px;
    margin-bottom: 16px;
    background: linear-gradient(135deg, #FEF3E2 0%, #FDF8F0 100%);
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08), 0 2px 8px rgba(0, 0, 0, 0.04);
    position: relative;
    overflow: hidden;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    border-left: 4px solid #F59E0B;
  `;
  
  // ホバー効果
  errorCard.onmouseenter = () => {
    errorCard.style.transform = "translateY(-2px)";
    errorCard.style.boxShadow = "0 8px 30px rgba(0, 0, 0, 0.12), 0 4px 12px rgba(0, 0, 0, 0.08)";
  };
  errorCard.onmouseleave = () => {
    errorCard.style.transform = "translateY(0)";
    errorCard.style.boxShadow = "0 4px 20px rgba(0, 0, 0, 0.08), 0 2px 8px rgba(0, 0, 0, 0.04)";
  };
  
  // ヘッダー（スライド番号、テキストボックス番号）
  const header = document.createElement("div");
  header.style.cssText = `
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 16px;
    padding-bottom: 12px;
    border-bottom: 1px solid rgba(0, 0, 0, 0.06);
  `;
  
  const locationInfo = document.createElement("span");
  locationInfo.style.cssText = `
    font-size: 12px;
    font-weight: 600;
    color: #6B7280;
    background: rgba(107, 114, 128, 0.1);
    padding: 4px 8px;
    border-radius: 6px;
  `;
  locationInfo.textContent = `スライド ${slideNum} - テキストボックス ${shapeNum}`;
  
  const statusBadge = document.createElement("span");
  statusBadge.style.cssText = `
    font-size: 14px;
    font-weight: 600;
    padding: 6px 12px;
    border-radius: 20px;
    background: #F59E0B;
    color: white;
    display: flex;
    align-items: center;
    gap: 6px;
  `;
  statusBadge.innerHTML = `⚠️ エラー`;
  
  header.appendChild(locationInfo);
  header.appendChild(statusBadge);
  errorCard.appendChild(header);
  
  // クレーム（チェックした文章）
  const claimDiv = document.createElement("div");
  claimDiv.style.cssText = `
    font-size: 16px;
    line-height: 1.6;
    margin-bottom: 16px;
    color: #1F2937;
    padding: 16px;
    background: rgba(255, 255, 255, 0.6);
    border-radius: 12px;
    border-left: 3px solid #F59E0B;
    font-weight: 500;
  `;
  claimDiv.textContent = `"${claim}"`;
  errorCard.appendChild(claimDiv);
  
  // エラーメッセージ
  const errorDiv = document.createElement("div");
  errorDiv.style.cssText = `
    color: #B45309;
    margin-bottom: 12px;
    font-weight: 600;
    font-size: 14px;
    padding: 12px;
    background: rgba(251, 191, 36, 0.1);
    border-radius: 8px;
    display: flex;
    align-items: center;
    gap: 8px;
  `;
  
  if (errorMsg && errorMsg.includes("有効な文章がありません")) {
    errorDiv.innerHTML = `<span style="font-size: 18px;">📝</span> <span>テキストなし</span>`;
  } else {
    errorDiv.innerHTML = `<span style="font-size: 18px;">📋</span> <span>ファクトチェックの結果がありませんでした</span>`;
  }
  errorCard.appendChild(errorDiv);
  
  // エラー詳細
  if (errorMsg) {
    const detailDiv = document.createElement("div");
    detailDiv.style.cssText = `
      font-size: 12px;
      color: #6B7280;
      margin-bottom: 12px;
      padding: 8px 12px;
      background: rgba(107, 114, 128, 0.05);
      border-radius: 6px;
    `;
    detailDiv.innerHTML = `<strong>詳細:</strong> ${errorMsg}`;
    errorCard.appendChild(detailDiv);
  }
  
  // 一般的なメッセージ
  const messageDiv = document.createElement("div");
  messageDiv.style.cssText = `
    font-size: 13px;
    color: #6B7280;
    line-height: 1.5;
    padding: 12px;
    background: rgba(107, 114, 128, 0.03);
    border-radius: 8px;
    border: 1px solid rgba(107, 114, 128, 0.1);
  `;
  
  if (errorMsg && errorMsg.includes("有効な文章がありません")) {
    messageDiv.innerHTML = `このテキストボックスにはファクトチェック可能なテキストが含まれていません。`;
  } else {
    messageDiv.innerHTML = `この文章についてはファクトチェックを実行できませんでした。内容が複雑すぎるか、信頼できる情報源が見つからなかった可能性があります。`;
  }
  errorCard.appendChild(messageDiv);
  
  // 検索ボタンを追加（有効な文章がない場合以外）
  if (!(errorMsg && errorMsg.includes("有効な文章がありません"))) {
    const searchButtonsDiv = document.createElement("div");
    searchButtonsDiv.style.cssText = "margin-top: 16px; text-align: center; display: flex; gap: 12px; justify-content: center;";
    
    // Tavilyボタン
    const tavilyButton = document.createElement("button");
    tavilyButton.innerHTML = "🔍 Tavilyで検索";
    tavilyButton.style.cssText = `
      background: linear-gradient(135deg, #6366F1 0%, #4F46E5 100%);
      color: white;
      border: none;
      padding: 10px 20px;
      border-radius: 10px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      box-shadow: 0 2px 8px rgba(99, 102, 241, 0.3);
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      display: inline-flex;
      align-items: center;
      gap: 6px;
    `;
    
    // ホバー効果
    tavilyButton.onmouseenter = () => {
      tavilyButton.style.transform = "translateY(-1px)";
      tavilyButton.style.boxShadow = "0 4px 16px rgba(99, 102, 241, 0.4)";
    };
    tavilyButton.onmouseleave = () => {
      tavilyButton.style.transform = "translateY(0)";
      tavilyButton.style.boxShadow = "0 2px 8px rgba(99, 102, 241, 0.3)";
    };
    
    // クリックイベント
    tavilyButton.onclick = async () => {
      await searchWithTavily(slideNum, shapeNum, claim);
    };
    
    // Googleボタン
    const googleButton = document.createElement("button");
    googleButton.innerHTML = "🔍 Googleで検索";
    googleButton.style.cssText = `
      background: linear-gradient(135deg, #34A853 0%, #0F9D58 100%);
      color: white;
      border: none;
      padding: 10px 20px;
      border-radius: 10px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
      box-shadow: 0 2px 8px rgba(52, 168, 83, 0.3);
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      display: inline-flex;
      align-items: center;
      gap: 6px;
    `;
    
    // ホバー効果
    googleButton.onmouseenter = () => {
      googleButton.style.transform = "translateY(-1px)";
      googleButton.style.boxShadow = "0 4px 16px rgba(52, 168, 83, 0.4)";
    };
    googleButton.onmouseleave = () => {
      googleButton.style.transform = "translateY(0)";
      googleButton.style.boxShadow = "0 2px 8px rgba(52, 168, 83, 0.3)";
    };
    
    // クリックイベント
    googleButton.onclick = async () => {
      await searchWithGoogle(slideNum, shapeNum, claim);
    };
    
    searchButtonsDiv.appendChild(tavilyButton);
    searchButtonsDiv.appendChild(googleButton);
    errorCard.appendChild(searchButtonsDiv);
  }
  
  // カードをコンテナに追加
  resultsContainer.appendChild(errorCard);
  
  // スクロールして最新の結果が見えるようにする
  resultsContainer.scrollTop = resultsContainer.scrollHeight;
}


////////////////////////////////////////////////////////////////////////////////
// correctFactCheckResult(): ファクトチェックで誤りと判定された文章を修正する
async function correctFactCheckResult(slideNum, originalClaim, factCheckResult) {
  logToScreen(`▶ 修正処理開始: スライド${slideNum} "${originalClaim}"`);
  
  try {
    // 正しい内容を取得
    const correction = await getCorrectionSuggestion(originalClaim, factCheckResult);
    
    if (!correction) {
      logToScreen("× 修正内容の取得に失敗しました");
      return;
    }
    
    // PowerPoint内のテキストを実際に修正
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      
      if (slideNum <= slides.items.length) {
        const slide = slides.items[slideNum - 1];
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();
        
        // 該当するテキストボックスを見つけて修正
        for (let shapeIndex = 0; shapeIndex < shapes.items.length; shapeIndex++) {
          const shape = shapes.items[shapeIndex];
          
          if (!shape.textFrame) continue;
          
          shape.textFrame.load("hasText");
          await context.sync();
          
          if (!shape.textFrame.hasText) continue;
          
          shape.textFrame.textRange.load("text");
          await context.sync();
          
          const currentText = shape.textFrame.textRange.text;
          
          // 該当する文章が含まれているかチェック
          if (currentText.includes(originalClaim)) {
            const newText = currentText.replace(originalClaim, correction.correctedText);
            shape.textFrame.textRange.text = newText;
            
            // 修正箇所を緑色で強調
            const correctionStart = newText.indexOf(correction.correctedText);
            
            if (correctionStart !== -1) {
              const correctedRange = shape.textFrame.textRange.getSubstring(correctionStart, correction.correctedText.length);
              correctedRange.font.color = "00AA00"; // 濃い緑色
              correctedRange.font.bold = true;
            }
            
            await context.sync();
            logToScreen(`✅ 修正完了: "${correction.correctedText}"`);
            
            // UI上でも修正完了を表示
            displayCorrectionComplete(slideNum, originalClaim, correction);
            return;
          }
        }
      }
      
      logToScreen("× 該当するテキストボックスが見つかりませんでした");
    });
    
  } catch (error) {
    logToScreen(`× 修正処理エラー: ${error.message}`);
    console.error("修正処理エラー:", error);
  }
}

////////////////////////////////////////////////////////////////////////////////
// getCorrectionSuggestion(): 誤った内容に対する正しい修正案を取得
async function getCorrectionSuggestion(originalClaim, factCheckResult) {
  logToScreen("▶ 修正案を取得中...");
  
  const JINA_TOKEN = API_CONFIG.JINA_API_TOKEN;
  const endpoint = "https://deepsearch.jina.ai/v1/chat/completions";

  const body = {
    model: "jina-chat",
    messages: [
      {
        role: "user",
        content: `The following claim has been fact-checked and found to be false:

INCORRECT CLAIM: "${originalClaim}"
FACT-CHECK REASON: "${factCheckResult.reason}"

Please provide a corrected version of this claim that is factually accurate. Return the result in JSON format:

{
  "correctedText": "The factually correct version of the claim",
  "explanation": "Brief explanation of the correction in Japanese"
}

Make the correction concise and maintain similar sentence structure when possible.`
      }
    ],
    stream: false,
    temperature: 0.1,
    search: true
  };

  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 15000);
    
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
    
    if (!res.ok) {
      logToScreen(`× Jina APIエラー: ${res.status}`);
      return null;
    }

    const data = await res.json();
    
    // JSONレスポンスを解析
    let responseData;
    if (data.choices && data.choices[0] && data.choices[0].message) {
      let content = data.choices[0].message.content;
      
      try {
        // JSONコードブロックを削除
        content = content.replace(/^```json\s*\n?/, '').replace(/\n?```\s*$/, '');
        responseData = JSON.parse(content);
      } catch (e) {
        logToScreen("× JSON解析失敗、テキストから抽出を試行");
        // フォールバック: テキストから直接抽出
        const correctedMatch = content.match(/"correctedText":\s*"([^"]+)"/);
        const explanationMatch = content.match(/"explanation":\s*"([^"]+)"/);
        
        if (correctedMatch) {
          responseData = {
            correctedText: correctedMatch[1],
            explanation: explanationMatch ? explanationMatch[1] : "修正されました"
          };
        }
      }
    }
    
    if (responseData && responseData.correctedText) {
      logToScreen(`✅ 修正案取得: "${responseData.correctedText}"`);
      return responseData;
    } else {
      logToScreen("× 有効な修正案が取得できませんでした");
      return null;
    }
    
  } catch (e) {
    if (e.name === 'AbortError') {
      logToScreen("× 修正案取得タイムアウト");
    } else {
      logToScreen(`× 修正案取得エラー: ${e.message}`);
    }
    return null;
  }
}

////////////////////////////////////////////////////////////////////////////////
// displayCorrectionComplete(): 修正完了メッセージを表示
function displayCorrectionComplete(slideNum, originalClaim, correction) {
  const resultsContainer = document.getElementById("resultsContainer");
  if (!resultsContainer) return;
  
  // 修正完了カード（モダンデザイン）
  const correctionCard = document.createElement("div");
  correctionCard.style.cssText = `
    border: none;
    border-radius: 16px;
    padding: 20px;
    margin-bottom: 16px;
    background: linear-gradient(135deg, #ECFDF5 0%, #F0FDF4 100%);
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08), 0 2px 8px rgba(0, 0, 0, 0.04);
    position: relative;
    overflow: hidden;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    border-left: 4px solid #10B981;
    animation: slideInUp 0.5s ease-out;
  `;
  
  // アニメーション定義
  if (!document.getElementById('correction-animations')) {
    const style = document.createElement('style');
    style.id = 'correction-animations';
    style.textContent = `
      @keyframes slideInUp {
        from {
          opacity: 0;
          transform: translateY(20px);
        }
        to {
          opacity: 1;
          transform: translateY(0);
        }
      }
    `;
    document.head.appendChild(style);
  }
  
  // ホバー効果
  correctionCard.onmouseenter = () => {
    correctionCard.style.transform = "translateY(-2px)";
    correctionCard.style.boxShadow = "0 8px 30px rgba(0, 0, 0, 0.12), 0 4px 12px rgba(0, 0, 0, 0.08)";
  };
  correctionCard.onmouseleave = () => {
    correctionCard.style.transform = "translateY(0)";
    correctionCard.style.boxShadow = "0 4px 20px rgba(0, 0, 0, 0.08), 0 2px 8px rgba(0, 0, 0, 0.04)";
  };
  
  // ヘッダー
  const header = document.createElement("div");
  header.style.cssText = `
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 16px;
    padding-bottom: 12px;
    border-bottom: 1px solid rgba(16, 185, 129, 0.2);
  `;
  
  const successIcon = document.createElement("span");
  successIcon.style.cssText = `
    font-size: 24px;
    display: flex;
    align-items: center;
    justify-content: center;
    width: 40px;
    height: 40px;
    background: linear-gradient(135deg, #10B981 0%, #059669 100%);
    border-radius: 12px;
    color: white;
  `;
  successIcon.textContent = "✅";
  
  const titleText = document.createElement("div");
  titleText.style.cssText = `
    font-size: 18px;
    font-weight: 700;
    color: #065F46;
  `;
  titleText.textContent = `修正完了 - スライド ${slideNum}`;
  
  header.appendChild(successIcon);
  header.appendChild(titleText);
  correctionCard.appendChild(header);
  
  // 修正前テキスト
  const originalDiv = document.createElement("div");
  originalDiv.style.cssText = `
    margin-bottom: 16px;
    padding: 12px;
    background: rgba(239, 68, 68, 0.1);
    border-radius: 8px;
    border-left: 3px solid #EF4444;
  `;
  originalDiv.innerHTML = `
    <div style="font-size: 12px; font-weight: 600; color: #991B1B; margin-bottom: 6px;">修正前</div>
    <div style="text-decoration: line-through; color: #6B7280; font-style: italic;">"${originalClaim}"</div>
  `;
  correctionCard.appendChild(originalDiv);
  
  // 修正後テキスト
  const correctedDiv = document.createElement("div");
  correctedDiv.style.cssText = `
    margin-bottom: 16px;
    padding: 12px;
    background: rgba(16, 185, 129, 0.1);
    border-radius: 8px;
    border-left: 3px solid #10B981;
  `;
  correctedDiv.innerHTML = `
    <div style="font-size: 12px; font-weight: 600; color: #065F46; margin-bottom: 6px;">修正後</div>
    <div style="color: #065F46; font-weight: 600;">"${correction.correctedText}"</div>
  `;
  correctionCard.appendChild(correctedDiv);
  
  // 説明（あれば）
  if (correction.explanation) {
    const explanationDiv = document.createElement("div");
    explanationDiv.style.cssText = `
      padding: 12px;
      background: rgba(107, 114, 128, 0.05);
      border-radius: 8px;
      border: 1px solid rgba(107, 114, 128, 0.1);
    `;
    explanationDiv.innerHTML = `
      <div style="font-size: 12px; font-weight: 600; color: #374151; margin-bottom: 4px;">説明</div>
      <div style="font-size: 13px; color: #6B7280; line-height: 1.5;">${correction.explanation}</div>
    `;
    correctionCard.appendChild(explanationDiv);
  }
  
  resultsContainer.appendChild(correctionCard);
  resultsContainer.scrollTop = resultsContainer.scrollHeight;
}

////////////////////////////////////////////////////////////////////////////////
// createFactCheckReportSlide(): ファクトチェック結果のレポートスライドを作成
async function createFactCheckReportSlide(context, allResults) {
  logToScreen(`▶ レポート作成開始: ${allResults.length}件の結果`);
  
  // 新しいスライドを最後に追加
  const newSlide = context.presentation.slides.add();
  await context.sync();
  logToScreen("▶ 新しいスライドを追加しました");

  // タイトルを追加
  const titleShape = newSlide.shapes.addTextBox({
    left: 50,
    top: 50,
    height: 80,
    width: 650
  });
  
  titleShape.textFrame.textRange.text = "ファクトチェック レポート";
  
  // タイトルのフォーマット
  titleShape.textFrame.textRange.font.size = 32;
  titleShape.textFrame.textRange.font.bold = true;
  titleShape.textFrame.textRange.font.color = "1f4e79";
  
  await context.sync();
  logToScreen("▶ タイトルを設定しました");
  
  // 統計情報を計算
  const totalChecks = allResults.length;
  const trueResults = allResults.filter(r => r.result === true).length;
  const falseResults = allResults.filter(r => r.result === false).length;
  const errorResults = allResults.filter(r => r.result === "error").length;
  const unknownResults = totalChecks - trueResults - falseResults - errorResults;

  // サマリーを追加
  const summaryText = `検証項目数: ${totalChecks}件
✅ 正しい: ${trueResults}件
❌ 誤り: ${falseResults}件
❓ 不明: ${unknownResults}件
⚠️ エラー: ${errorResults}件

詳細結果:`;

  logToScreen(`▶ 統計: 合計${totalChecks}件 (正${trueResults}, 誤${falseResults}, 不明${unknownResults}, エラー${errorResults})`);

  const summaryShape = newSlide.shapes.addTextBox({
    left: 50,
    top: 150,
    height: 200,
    width: 650
  });
  
  summaryShape.textFrame.textRange.text = summaryText;
  summaryShape.textFrame.textRange.font.size = 16;
  
  await context.sync();
  logToScreen("▶ サマリーを設定しました");

  // 詳細結果を追加
  let detailText = "";
  let yPos = 380;
  
  allResults.forEach((result, index) => {
    const icon = result.result === true ? "✅" : 
                 result.result === false ? "❌" : 
                 result.result === "error" ? "⚠️" : "❓";
    
    detailText += `${index + 1}. ${icon} スライド${result.slideNumber}-図形${result.shapeNumber}\n`;
    detailText += `   「${result.sentence}」\n`;
    detailText += `   判定: ${result.reason}\n`;
    
    if (result.factuality !== null) {
      const percentage = (result.factuality * 100).toFixed(0);
      detailText += `   信頼度: ${percentage}%\n`;
    }
    
    if (result.references && result.references.length > 0) {
      detailText += `   参考: ${result.references.length}件の情報源\n`;
    }
    
    detailText += "\n";
  });

  // 詳細結果が長すぎる場合は分割
  if (detailText.length > 2000) {
    // 最初の部分のみ表示
    const truncatedText = detailText.substring(0, 1800) + "\n\n... (結果が多いため一部省略)";
    detailText = truncatedText;
  }

  logToScreen(`▶ 詳細テキスト長: ${detailText.length}文字`);

  const detailShape = newSlide.shapes.addTextBox({
    left: 50,
    top: yPos,
    height: 500,
    width: 650
  });
  
  detailShape.textFrame.textRange.text = detailText;
  detailShape.textFrame.textRange.font.size = 12;
  
  await context.sync();
  logToScreen("▶ 詳細結果を設定しました");
  
  // 生成日時を追加
  const now = new Date();
  const timestamp = `生成日時: ${now.toLocaleString('ja-JP')}`;
  
  const timestampShape = newSlide.shapes.addTextBox({
    left: 450,
    top: 900,
    height: 30,
    width: 250
  });
  
  timestampShape.textFrame.textRange.text = timestamp;
  timestampShape.textFrame.textRange.font.size = 10;
  timestampShape.textFrame.textRange.font.color = "666666";

  await context.sync();
  logToScreen("▶ レポートスライド作成完了");
}

////////////////////////////////////////////////////////////////////////////////
// callTavilySearch(): Tavily Search API を使用して検索を実行
async function callTavilySearch(query) {
  // Tavily API キー（config.jsから取得）
  const TAVILY_API_KEY = API_CONFIG.TAVILY_API_KEY;
  
  // デバッグ: APIキーの一部を表示
  console.log(`[DEBUG] Using TAVILY_API_KEY: ${TAVILY_API_KEY.substring(0, 10)}...`);
  logToScreen(`[DEBUG] TAVILY_API_KEY: ${TAVILY_API_KEY.substring(0, 10)}...`);
  
  const endpoint = "https://api.tavily.com/search";
  
  const body = {
    api_key: TAVILY_API_KEY,
    query: query,
    search_depth: "advanced",
    include_answer: true,
    include_raw_content: true,
    max_results: 5,
    include_domains: [
      "wikipedia.org",
      "britannica.com",
      "nature.com",
      "science.org",
      "sciencedirect.com",
      "pubmed.ncbi.nlm.nih.gov",
      "scholar.google.com",
      "reuters.com",
      "apnews.com",
      "bbc.com",
      "nhk.or.jp",
      "go.jp",
      "gov",
      "edu",
      "ac.jp",
      "who.int",
      "un.org",
      "oecd.org",
      "worldbank.org"
    ],
    exclude_domains: []
  };
  
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);
    
    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body),
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    
    if (!res.ok) {
      logToScreen(`× Tavily APIエラー: ${res.status}`);
      return null;
    }
    
    const data = await res.json();
    return data;
    
  } catch (e) {
    if (e.name === 'AbortError') {
      logToScreen("× Tavily タイムアウト (30秒)");
    } else {
      logToScreen(`× Tavily エラー: ${e.message}`);
    }
    return null;
  }
}

////////////////////////////////////////////////////////////////////////////////
// searchWithGoogle(): Googleで検索を実行し、結果を表示
async function searchWithGoogle(slideNum, shapeNum, claim) {
  logToScreen(`▶ Googleで検索開始: "${claim}"`);
  
  // 検索中モーダルを表示
  showSearchingModal("Google");
  
  try {
    const searchResults = await callGoogleSearch(claim);
    
    if (!searchResults || !searchResults.results || searchResults.results.length === 0) {
      hideSearchingModal();
      showNoResultsModal("Google");
      return;
    }
    
    // 検索結果を表示
    hideSearchingModal();
    showSearchResultsModal(slideNum, shapeNum, claim, searchResults, "Google");
    
  } catch (error) {
    hideSearchingModal();
    logToScreen(`× Google検索エラー: ${error.message}`);
    showErrorModal("検索中にエラーが発生しました");
  }
}

////////////////////////////////////////////////////////////////////////////////
// searchWithTavily(): Tavilyで検索を実行し、結果を表示
async function searchWithTavily(slideNum, shapeNum, claim) {
  logToScreen(`▶ Tavilyで検索開始: "${claim}"`);
  
  // 検索中モーダルを表示
  showSearchingModal("Tavily");
  
  try {
    const searchResults = await callTavilySearch(claim);
    
    if (!searchResults || !searchResults.results || searchResults.results.length === 0) {
      hideSearchingModal();
      showNoResultsModal("Tavily");
      return;
    }
    
    // 検索結果を表示
    hideSearchingModal();
    showSearchResultsModal(slideNum, shapeNum, claim, searchResults, "Tavily");
    
  } catch (error) {
    hideSearchingModal();
    logToScreen(`× Tavily検索エラー: ${error.message}`);
    showErrorModal("検索中にエラーが発生しました");
  }
}

////////////////////////////////////////////////////////////////////////////////
// showSearchingModal(): 検索中モーダルを表示
function showSearchingModal(searchEngine = "Tavily") {
  // 既存のモーダルを削除
  const existingModal = document.getElementById("searchModal");
  if (existingModal) existingModal.remove();
  
  const modal = document.createElement("div");
  modal.id = "searchModal";
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.5);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 10000;
  `;
  
  const modalContent = document.createElement("div");
  modalContent.style.cssText = `
    background: white;
    border-radius: 16px;
    padding: 32px;
    text-align: center;
    box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
  `;
  
  modalContent.innerHTML = `
    <div class="loading-spinner" style="margin: 0 auto 20px;"></div>
    <h3 style="margin: 0 0 8px 0; font-size: 18px; color: #1F2937;">${searchEngine}で検索中...</h3>
    <p style="margin: 0; color: #6B7280; font-size: 14px;">信頼できる情報源を探しています</p>
  `;
  
  modal.appendChild(modalContent);
  document.body.appendChild(modal);
}

////////////////////////////////////////////////////////////////////////////////
// hideSearchingModal(): 検索中モーダルを非表示
function hideSearchingModal() {
  const modal = document.getElementById("searchModal");
  if (modal) modal.remove();
}

////////////////////////////////////////////////////////////////////////////////
// getTrustLevelLabel(): 信頼性レベルのラベルを取得
function getTrustLevelLabel(trustLevel) {
  const labels = {
    government: "政府機関",
    academic: "教育機関",
    scientific: "学術論文",
    encyclopedia: "百科事典",
    news: "報道機関",
    international: "国際機関",
    medium: "一般"
  };
  return labels[trustLevel] || "一般";
}

////////////////////////////////////////////////////////////////////////////////
// showSearchResultsModal(): 検索結果を表示するモーダル
function showSearchResultsModal(slideNum, shapeNum, claim, searchResults, searchEngine = "Tavily") {
  // 既存のモーダルを削除
  const existingModal = document.getElementById("searchModal");
  if (existingModal) existingModal.remove();
  
  const modal = document.createElement("div");
  modal.id = "searchModal";
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.5);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 10000;
    padding: 20px;
  `;
  
  const modalContent = document.createElement("div");
  modalContent.style.cssText = `
    background: white;
    border-radius: 16px;
    max-width: 800px;
    width: 100%;
    max-height: 80vh;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
  `;
  
  // ヘッダー
  const header = document.createElement("div");
  header.style.cssText = `
    padding: 24px;
    border-bottom: 1px solid #E5E7EB;
    background: linear-gradient(135deg, #6366F1 0%, #4F46E5 100%);
    color: white;
  `;
  
  header.innerHTML = `
    <div style="display: flex; justify-content: space-between; align-items: center;">
      <div>
        <h2 style="margin: 0 0 8px 0; font-size: 20px;">${searchEngine}検索結果</h2>
        <p style="margin: 0; font-size: 14px; opacity: 0.9;">以下の情報から正しい内容を選択してください</p>
      </div>
      <button id="closeModal" style="
        background: rgba(255, 255, 255, 0.2);
        border: none;
        color: white;
        width: 32px;
        height: 32px;
        border-radius: 8px;
        cursor: pointer;
        font-size: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
      ">×</button>
    </div>
  `;
  
  // 元の文章
  const originalClaim = document.createElement("div");
  originalClaim.style.cssText = `
    padding: 16px 24px;
    background: #FEF3E2;
    border-bottom: 1px solid #E5E7EB;
  `;
  originalClaim.innerHTML = `
    <div style="font-size: 12px; color: #92400E; font-weight: 600; margin-bottom: 4px;">チェック対象の文章:</div>
    <div style="font-size: 14px; color: #1F2937;">"${claim}"</div>
  `;
  
  // 検索結果リスト
  const resultsContainer = document.createElement("div");
  resultsContainer.style.cssText = `
    flex: 1;
    overflow-y: auto;
    padding: 24px;
  `;
  
  // AI回答がある場合は最初に表示
  if (searchResults.answer) {
    const aiAnswer = document.createElement("div");
    aiAnswer.style.cssText = `
      background: linear-gradient(135deg, #EFF6FF 0%, #F0F9FF 100%);
      border: 1px solid #3B82F6;
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 24px;
      cursor: pointer;
      transition: all 0.3s ease;
    `;
    
    aiAnswer.innerHTML = `
      <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 12px;">
        <span style="font-size: 20px;">🤖</span>
        <h3 style="margin: 0; font-size: 16px; color: #1E40AF;">AI統合回答</h3>
      </div>
      <p style="margin: 0; color: #374151; line-height: 1.6;">${searchResults.answer}</p>
      <button class="selectContent" data-content="${searchResults.answer.replace(/"/g, '&quot;')}" style="
        margin-top: 12px;
        background: #3B82F6;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 8px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 600;
      ">この内容を使用</button>
    `;
    
    resultsContainer.appendChild(aiAnswer);
  }
  
  // 各検索結果
  searchResults.results.forEach((result, index) => {
    const resultCard = document.createElement("div");
    resultCard.style.cssText = `
      background: white;
      border: 1px solid #E5E7EB;
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 16px;
      transition: all 0.3s ease;
    `;
    
    // 信頼性に基づいて背景色を設定
    let backgroundColor = "#FFFFFF";
    if (result.trustLevel === "government" || result.trustLevel === "international") {
      backgroundColor = "#EFF6FF"; // 青系
    } else if (result.trustLevel === "academic" || result.trustLevel === "scientific") {
      backgroundColor = "#F0FDF4"; // 緑系
    } else if (result.trustLevel === "encyclopedia") {
      backgroundColor = "#FEF3E2"; // オレンジ系
    }
    
    resultCard.style.backgroundColor = backgroundColor;
    
    resultCard.innerHTML = `
      <div style="display: flex; align-items: start; gap: 12px;">
        <div style="flex: 1;">
          <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
            <span style="font-size: 20px;">${result.trustIcon || '🔍'}</span>
            <h3 style="margin: 0; font-size: 16px; color: #1F2937;">
              ${index + 1}. ${result.title}
            </h3>
          </div>
          <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
            <a href="${result.url}" target="_blank" style="
              color: #3B82F6;
              text-decoration: none;
              font-size: 12px;
            ">${result.displayLink || result.url}</a>
            ${result.trustLevel ? `<span style="
              font-size: 11px;
              padding: 2px 8px;
              border-radius: 12px;
              background: rgba(59, 130, 246, 0.1);
              color: #2563EB;
              font-weight: 600;
            ">${getTrustLevelLabel(result.trustLevel)}</span>` : ''}
          </div>
          <p style="margin: 12px 0; color: #4B5563; line-height: 1.6; font-size: 14px;">
            ${result.content}
          </p>
          <button class="selectContent" data-content="${result.content.replace(/"/g, '&quot;')}" data-url="${result.url}" style="
            background: #10B981;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s ease;
          ">この内容を使用</button>
        </div>
      </div>
    `;
    
    resultsContainer.appendChild(resultCard);
  });
  
  // フッター
  const footer = document.createElement("div");
  footer.style.cssText = `
    padding: 16px 24px;
    border-top: 1px solid #E5E7EB;
    background: #F9FAFB;
    display: flex;
    justify-content: space-between;
    align-items: center;
  `;
  
  footer.innerHTML = `
    <p style="margin: 0; color: #6B7280; font-size: 12px;">
      ${searchResults.results.length}件の検索結果
    </p>
    <button id="cancelButton" style="
      background: #6B7280;
      color: white;
      border: none;
      padding: 10px 20px;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
    ">キャンセル</button>
  `;
  
  modalContent.appendChild(header);
  modalContent.appendChild(originalClaim);
  modalContent.appendChild(resultsContainer);
  modalContent.appendChild(footer);
  modal.appendChild(modalContent);
  document.body.appendChild(modal);
  
  // イベントリスナー
  document.getElementById("closeModal").onclick = () => modal.remove();
  document.getElementById("cancelButton").onclick = () => modal.remove();
  
  // 各選択ボタンのイベント
  modal.querySelectorAll(".selectContent").forEach(button => {
    button.onclick = async () => {
      const content = button.getAttribute("data-content");
      const url = button.getAttribute("data-url");
      
      // 選択された内容でファクトチェック結果を作成
      const tavilyResult = {
        hit: true,
        result: true,
        reason: `${searchEngine}の検索結果から選択された情報`,
        factuality: 0.8,
        references: url ? [{
          url: url,
          keyQuote: content,
          isSupportive: true
        }] : [],
        source: searchEngine
      };
      
      // モーダルを閉じる
      modal.remove();
      
      // 結果を表示
      displayFactCheckResultWithNumber(slideNum, shapeNum, content, tavilyResult, ` (${searchEngine}検索結果)`);
      
      // 修正を適用
      await applySearchCorrection(slideNum, shapeNum, claim, content, searchEngine);
    };
  });
}

////////////////////////////////////////////////////////////////////////////////
// applySearchCorrection(): 検索で選択された内容をPowerPointに適用
async function applySearchCorrection(slideNum, shapeNum, originalText, correctedText, searchEngine = "Tavily") {
  logToScreen(`▶ ${searchEngine}修正適用: スライド${slideNum} "${originalText}" → "${correctedText}"`);
  
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      
      if (slideNum <= slides.items.length) {
        const slide = slides.items[slideNum - 1];
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();
        
        // 該当するテキストボックスを見つけて修正
        for (let shapeIndex = 0; shapeIndex < shapes.items.length; shapeIndex++) {
          const shape = shapes.items[shapeIndex];
          
          if (!shape.textFrame) continue;
          
          shape.textFrame.load("hasText");
          await context.sync();
          
          if (!shape.textFrame.hasText) continue;
          
          shape.textFrame.textRange.load("text");
          await context.sync();
          
          const currentText = shape.textFrame.textRange.text;
          
          // 該当する文章が含まれているかチェック
          if (currentText.includes(originalText)) {
            const newText = currentText.replace(originalText, correctedText);
            shape.textFrame.textRange.text = newText;
            
            // 修正箇所を緑色で強調
            const correctionStart = newText.indexOf(correctedText);
            
            if (correctionStart !== -1) {
              const correctedRange = shape.textFrame.textRange.getSubstring(correctionStart, correctedText.length);
              correctedRange.font.color = "00AA00"; // 濃い緑色
              correctedRange.font.bold = true;
            }
            
            await context.sync();
            logToScreen(`✅ ${searchEngine}修正完了: "${correctedText}"`);
            
            // 修正完了メッセージを表示
            const correction = {
              correctedText: correctedText,
              explanation: `${searchEngine}の検索結果から選択された情報で修正しました`
            };
            displayCorrectionComplete(slideNum, originalText, correction);
            return;
          }
        }
      }
      
      logToScreen("× 該当するテキストボックスが見つかりませんでした");
    });
    
  } catch (error) {
    logToScreen(`× ${searchEngine}修正適用エラー: ${error.message}`);
    console.error(`${searchEngine}修正適用エラー:`, error);
  }
}

////////////////////////////////////////////////////////////////////////////////
// showNoResultsModal(): 検索結果がない場合のモーダル
function showNoResultsModal(searchEngine = "Tavily") {
  const modal = document.createElement("div");
  modal.id = "searchModal";
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.5);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 10000;
  `;
  
  const modalContent = document.createElement("div");
  modalContent.style.cssText = `
    background: white;
    border-radius: 16px;
    padding: 32px;
    text-align: center;
    max-width: 400px;
    box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
  `;
  
  modalContent.innerHTML = `
    <div style="font-size: 48px; margin-bottom: 16px;">😔</div>
    <h3 style="margin: 0 0 12px 0; font-size: 18px; color: #1F2937;">検索結果が見つかりませんでした</h3>
    <p style="margin: 0 0 24px 0; color: #6B7280; font-size: 14px;">
      この文章に関する信頼できる情報源が見つかりませんでした。
    </p>
    <button onclick="document.getElementById('searchModal').remove()" style="
      background: #3B82F6;
      color: white;
      border: none;
      padding: 10px 20px;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
    ">閉じる</button>
  `;
  
  modal.appendChild(modalContent);
  document.body.appendChild(modal);
}

////////////////////////////////////////////////////////////////////////////////
// showErrorModal(): エラーモーダルを表示
function showErrorModal(message) {
  const modal = document.createElement("div");
  modal.id = "searchModal";
  modal.style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.5);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 10000;
  `;
  
  const modalContent = document.createElement("div");
  modalContent.style.cssText = `
    background: white;
    border-radius: 16px;
    padding: 32px;
    text-align: center;
    max-width: 400px;
    box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
  `;
  
  modalContent.innerHTML = `
    <div style="font-size: 48px; margin-bottom: 16px;">❌</div>
    <h3 style="margin: 0 0 12px 0; font-size: 18px; color: #1F2937;">エラーが発生しました</h3>
    <p style="margin: 0 0 24px 0; color: #6B7280; font-size: 14px;">${message}</p>
    <button onclick="document.getElementById('searchModal').remove()" style="
      background: #EF4444;
      color: white;
      border: none;
      padding: 10px 20px;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 600;
    ">閉じる</button>
  `;
  
  modal.appendChild(modalContent);
  document.body.appendChild(modal);
}

////////////////////////////////////////////////////////////////////////////////
// callGoogleSearch(): Google Custom Search API を使用して検索を実行
async function callGoogleSearch(query) {
  // Google Custom Search API の設定（config.jsから取得）
  const GOOGLE_API_KEY = API_CONFIG.GOOGLE_API_KEY;
  const SEARCH_ENGINE_ID = API_CONFIG.GOOGLE_SEARCH_ENGINE_ID;
  const endpoint = `https://www.googleapis.com/customsearch/v1`;
  
  // 信頼できるドメインリスト
  const trustedDomains = [
    "site:wikipedia.org",
    "site:britannica.com",
    "site:nature.com",
    "site:science.org",
    "site:sciencedirect.com",
    "site:pubmed.ncbi.nlm.nih.gov",
    "site:scholar.google.com",
    "site:reuters.com",
    "site:apnews.com",
    "site:bbc.com",
    "site:cnn.com",
    "site:nytimes.com",
    "site:washingtonpost.com",
    "site:nhk.or.jp",
    "site:asahi.com",
    "site:nikkei.com",
    "site:*.go.jp",
    "site:*.gov",
    "site:*.edu",
    "site:*.ac.jp",
    "site:who.int",
    "site:un.org",
    "site:oecd.org",
    "site:worldbank.org",
    "site:imf.org"
  ];
  
  // 信頼できるドメインをOR条件で結合してクエリに追加
  const siteRestriction = trustedDomains.join(" OR ");
  const enhancedQuery = `${query} (${siteRestriction})`;
  
  const params = new URLSearchParams({
    key: GOOGLE_API_KEY,
    cx: SEARCH_ENGINE_ID,
    q: enhancedQuery,
    num: 10, // より多くの結果を取得して信頼できるソースを見つけやすくする
    // 日本語と英語の結果を取得
    lr: "lang_ja|lang_en",
    // セーフサーチを有効化
    safe: "active",
    // 関連性の高い結果を優先
    sort: "relevance"
  });
  
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);
    
    const res = await fetch(`${endpoint}?${params}`, {
      method: "GET",
      signal: controller.signal
    });
    
    clearTimeout(timeoutId);
    
    if (!res.ok) {
      logToScreen(`× Google Search APIエラー: ${res.status}`);
      return null;
    }
    
    const data = await res.json();
    
    // Google Search APIの結果をTavily形式に変換
    const formattedResults = {
      results: [],
      answer: null
    };
    
    if (data.items && data.items.length > 0) {
      formattedResults.results = data.items.map(item => {
        // ドメインから信頼性レベルを判定
        const domain = item.displayLink || "";
        let trustLevel = "medium";
        let trustIcon = "🔍";
        
        if (domain.includes(".gov") || domain.includes(".go.jp")) {
          trustLevel = "government";
          trustIcon = "🏛️";
        } else if (domain.includes(".edu") || domain.includes(".ac.jp")) {
          trustLevel = "academic";
          trustIcon = "🎓";
        } else if (domain.includes("wikipedia.org") || domain.includes("britannica.com")) {
          trustLevel = "encyclopedia";
          trustIcon = "📚";
        } else if (domain.includes("nature.com") || domain.includes("science.org") || 
                   domain.includes("pubmed") || domain.includes("scholar.google")) {
          trustLevel = "scientific";
          trustIcon = "🔬";
        } else if (domain.includes("reuters.com") || domain.includes("apnews.com") || 
                   domain.includes("bbc.com") || domain.includes("nhk.or.jp")) {
          trustLevel = "news";
          trustIcon = "📰";
        } else if (domain.includes("who.int") || domain.includes("un.org") || 
                   domain.includes("worldbank.org")) {
          trustLevel = "international";
          trustIcon = "🌐";
        }
        
        return {
          title: item.title,
          url: item.link,
          content: item.snippet || "",
          displayLink: item.displayLink,
          trustLevel: trustLevel,
          trustIcon: trustIcon
        };
      });
    }
    
    return formattedResults;
    
  } catch (e) {
    if (e.name === 'AbortError') {
      logToScreen("× Google Search タイムアウト (30秒)");
    } else {
      logToScreen(`× Google Search エラー: ${e.message}`);
    }
    return null;
  }
}

////////////////////////////////////////////////////////////////////////////////
// callJinaFactCheck(): Jina DeepSearch (Grounding) API を叩くユーティリティ
async function callJinaFactCheck(claim) {
  // デバッグ: process.envが存在するかチェック
  console.log(`[DEBUG] typeof process:`, typeof process);
  console.log(`[DEBUG] process.env available:`, typeof process !== 'undefined' && process.env !== undefined);
  
  // Jina トークン（config.jsから取得）
  const JINA_TOKEN = API_CONFIG.JINA_API_TOKEN;
  
  // デバッグ: APIキーの一部を表示（セキュリティのため最初の10文字のみ）
  console.log(`[DEBUG] Using JINA_TOKEN: ${JINA_TOKEN.substring(0, 10)}...`);
  logToScreen(`[DEBUG] JINA_TOKEN: ${JINA_TOKEN.substring(0, 10)}...`);
  
  // DeepSearch の Chat Completions エンドポイント
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
    debugLog(`▶ リクエストボディ: ${JSON.stringify(body, null, 2)}`);
    console.log(`[DEBUG] Sending request to Jina API:`, { endpoint, body });
    
    // AbortControllerでタイムアウトを設定（20秒）
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 20000);
    
    debugLog(`▶ fetch 開始...`);
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
    debugLog(`▶ Jina HTTP ステータス: ${res.status}`);
    console.log(`[DEBUG] Jina API response status:`, res.status);

    // JSON を読み取る
    let data;
    try {
      data = await res.json();
    } catch (parseErr) {
      debugLog(`× レスポンス JSON 解析エラー: ${parseErr}`);
      return { hit: false, error: "Jina レスポンス JSON 解析に失敗" };
    }

    if (!res.ok) {
      let errMsg = `ステータスコード ${res.status}`;
      
      // 特定のエラーコードに対する日本語メッセージ
      if (res.status === 402) {
        errMsg = "APIの利用残高が不足しています。Jinaアカウントにチャージしてください。";
      } else if (data && data.error) {
        errMsg = typeof data.error === 'string' ? data.error : (data.error.message || errMsg);
      }
      
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
    const finalResult = {
      hit: true,
      result: responseData.result ?? "",           // true か false
      reason: responseData.reason ?? "",
      factuality: responseData.factuality ?? null, // 数値スコア
      references: responseData.references ?? []     // 根拠リスト
    };
    
    debugLog(`▶ Jina API 最終結果: ${JSON.stringify(finalResult)}`);
    console.log(`[DEBUG] Jina API final result:`, finalResult);
    
    return finalResult;
  } catch (e) {
    if (e.name === 'AbortError') {
      debugLog(`× [Office] Jina タイムアウト (20秒): 処理が遅すぎます`);
      console.error(`[DEBUG] Jina API timeout after 20 seconds`);
      return { hit: false, error: "Jina API タイムアウト (20秒)" };
    }
    debugLog(`× [Office] Jina 呼び出し中に例外発生: ${e}`);
    console.error(`[DEBUG] Jina API call exception:`, e);
    console.error(`[DEBUG] Exception stack:`, e.stack);
    return { hit: false, error: e.message || String(e) };
  }
}













