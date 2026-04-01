// ==========================================
// 設定エリア
// ==========================================
// 個人情報（APIキー・メールアドレス）はスクリプトプロパティで管理
// 設定: 拡張機能 > Apps Script > プロジェクトの設定 > スクリプト プロパティ
// 必要なキー: GEMINI_API_KEY, EMAIL_RECIPIENT
// オプション: GMAIL_QUERY … 未設定なら下記 DEFAULT_GMAIL_QUERY を使用。Gmailのラベル名と一致しないと検索0件になる。
// 既定のGmail検索。ラベルにスペースがある場合はスクリプトプロパティで GMAIL_QUERY を指定（例: label:"10 scholar" label:unread）
const DEFAULT_GMAIL_QUERY = 'label:10_scholar label:unread';
// メール通知: 送信元・送信先はスクリプトプロパティ EMAIL_RECIPIENT で指定
// ※トリガーは通知先アカウントで設定すること（送信元＝実行アカウント）
// レポート出力: シート名「レポート」があればそこに追記。なければアクティブなシート（トリガー実行時は先頭シートになることが多い）
// ==========================================

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = (props.getProperty('GEMINI_API_KEY') || '').trim();
  const emailRecipient = (props.getProperty('EMAIL_RECIPIENT') || '').trim();
  if (!apiKey || !emailRecipient) {
    throw new Error('スクリプトプロパティが未設定です。拡張機能 > Apps Script > プロジェクトの設定 > スクリプト プロパティ で GEMINI_API_KEY と EMAIL_RECIPIENT を設定してください。');
  }
  return { apiKey, emailRecipient };
}

function main() {
  getConfig(); // スクリプトプロパティ未設定時は即エラー
  var gmailQuery = (PropertiesService.getScriptProperties().getProperty('GMAIL_QUERY') || '').trim() || DEFAULT_GMAIL_QUERY;
  var threads = GmailApp.search(gmailQuery, 0, 5);
  if (threads.length === 0) {
    console.log("未読メールなし");
    return;
  }

  // トリガー実行時は「アクティブなシート」が不定のため、名前で指定（なければ従来どおりアクティブシート）
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('レポート') || ss.getActiveSheet();
  const processedThreads = [];

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const message of messages) {
      if (!message.isUnread()) continue;

      const body = message.getPlainBody();
      
      // ★フェーズ1：AI編集長による「選定」
      console.log("AI編集長が審査中...");
      const selectedUrls = selectBestPapers(body);

      if (selectedUrls.length === 0) {
        console.log("→ 審査基準を満たす論文なし、またはAPIエラーのためスキップ。");
        continue;
      }

      // 選ばれた論文だけを処理する
      for (const url of selectedUrls) {
        console.log("★採用論文: " + url);
        
        // ★フェーズ2：全文取得
        const pageContent = fetchUrlContent(url);
        if (!pageContent || pageContent.length < 500) {
          console.log("中身が取れなかったためスキップ");
          continue;
        }

        // ★フェーズ3：AI記者による執筆
        const result = analyzeWithGemini(pageContent);

        if (result) {
          // KEYデータに数値が不足している場合は、数値のみ抽出するフォールバックを実行
          if (!keyEvidenceHasEnoughNumbers(result.key_evidence)) {
            console.log("KEYデータに数値が不足しているため、数値エビデンスを再抽出します");
            const numericalOnly = extractNumericalEvidenceOnly(pageContent);
            if (numericalOnly) {
              result.key_evidence = result.key_evidence
                ? result.key_evidence + "\n\n【数値エビデンス補足】\n" + numericalOnly
                : numericalOnly;
            }
          }
          // 論文の原文URLを正しく解決（Scholarリダイレクトから実際のURLを抽出）
          // ※ Geminiが推測する result.paper_url は誤URLを返すことがあるため使用しない
          const paperUrl = resolvePaperUrl(url) || url;
          // 保存
          sheet.appendRow([
            new Date(), 
            result.title, 
            result.summary_3lines, 
            result.summary_long,
            result.key_evidence,
            credibilityToDisplayString(result.credibility),
            result.japan_context, 
            result.sns_post, // ここに3つの案がまとめて入ります
            paperUrl,
            result.authors || '',
            result.publication_date || ''
          ]);
          // 通知
          sendRichEmail(result, paperUrl);
        }
        
        // 連続アクセス防止の休憩
        Utilities.sleep(10000);
      }
    }
    processedThreads.push(thread);
  }

  if (processedThreads.length > 0) {
    GmailApp.markThreadsRead(processedThreads);
  }
}

// -------------------------------------------------------
// 関数1：AI編集長（選定）
// -------------------------------------------------------
function selectBestPapers(emailBody) {
  const { apiKey } = getConfig();
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
  
  const prompt = `
  あなたは日本の少子化・子育て政策のトップインフルエンサー専属の「編集長」です。
  以下のGoogle Scholar通知メールのテキストから、
  **「日本の実務家・専門家にとって、今読むべき最も価値のある論文」を最大1本だけ** 選んでください。
  
  【厳格な審査基準】
  1. Relevance（適合性）: 日本社会の課題（少子化、貧困、親のメンタルなど）に直結するか。
  2. Counter-intuitive（意外性）: 「定説を覆すデータ」や「直感に反する発見」があるか。
  3. Actionable（実用性）: 具体的な政策提言や数値データがあるか。
  4. Credibility（信憑性）: 信頼できる手法・機関の研究か。（著者が不明瞭なものは除外）

  ※S級の論文がない場合は "[]" (空リスト) を返してください。
  
  出力形式:
  選定した論文のURLのみをJSON配列で返してください。Markdown記号は不要です。
  Example: ["https://scholar.google.com/...."]

  対象メール本文:
  ${emailBody.substring(0, 10000)}
  `;

  return callGeminiAPI(apiUrl, prompt, true); 
}

// -------------------------------------------------------
// 関数2：AI記者（執筆・分析）
// -------------------------------------------------------
function analyzeWithGemini(text) {
  const { apiKey } = getConfig();
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
  
  const prompt = `
  あなたは政策系インフルエンサーの「ゴーストライター」です。
  以下の論文全文を読み込み、私の活動（日本への情報発信）に必要なレポートを作成してください。

  【出力要件】
  1. タイトル: 結論ズバリのキャッチーな日本語タイトル。
  2. 3行要約: 専門用語を使わず3行で簡潔に。
  3. 500文字要約: 論文の背景、手法、結果、考察を詳細に500文字程度で解説。
  4. KEYデータ/エビデンス（厳守）:
     - **必ず具体的な数値を含めること。数値のない記述は1つも入れない。**
     - 各項目は次のいずれかを含むこと: 割合(%), 人数・サンプル数(n=), 効果量(d, OR, RR, HR等), p値, 信頼区間(95%CI), 平均値・標準偏差(M±SD), その他論文に明記された数値。
     - 良い例: 「介入群で抑うつスコアが23%低下（n=156, p<0.01, 95%CI [-0.8, -0.3]）」「効果量d=0.72」「サンプルは3,200人（女性58%）」
     - 悪い例: 「効果があった」「多くの人が改善した」「統計的に有意だった」→ これらは数値がないため禁止。
     - 最低3項目以上、それぞれに数値を入れて箇条書きで出力すること。
  5. 信憑性評価: 以下の3観点で評価すること。
     (1) 掲載媒体と査読: 査読（Peer Review）の有無、インパクトファクター(IF)の有無・水準、ハゲタカジャーナルでないことの確認。
     (2) 著者の属性と所属機関: 著者がその分野の専門家か、信頼できる研究機関（大学・公的機関等）に所属しているか。
     (3) 研究デザインと手法の妥当性: エビデンスレベル（メタ分析/RCT/観察研究等）、サンプルサイズが統計的に十分か。
  6. 日本の現状との比較: 日本のデータとぶつけ、建設的な提言を含める。
  7. 著者情報: 著者名と所属組織名（論文から抽出、不明な場合は「不明」）。
  8. 公開日: 論文の公開年月日（yyyy/mm/dd形式、不明な場合は「不明」）。

  9. SNS投稿案（3パターン）:
     以下の3つの切り口で、それぞれ投稿案を作成し、**1つのテキストにまとめて**出力してください。
     - 【案1：データ重視】数値や客観的事実を前面に出し、知的好奇心を刺激する。
     - 【案2：共感・当事者重視】親や現場の悩みに寄り添い、安心感や気づきを与える。
     - 【案3：建設的提言】社会実装や未来の解決策にフォーカスする。

     **【SNS投稿の絶対厳守ルール】**
     - 特定の個人、属性、集団を攻撃したり、煽ったり、分断を生む表現は**厳禁**。
     - 感情論ではなく、必ず論文のデータ・エビデンスに依拠すること。
     - ハッシュタグ（#）は一切つけないこと。
     - 「です・ます」調、または「だ・である」調のどちらかで統一し、自然な日本語で。

  JSON形式で出力（Markdown記号不要）:
  {
    "title": "...",
    "summary_3lines": "...",
    "summary_long": "...",
    "key_evidence": "...",
    "credibility": "...",
    "japan_context": "...",
    "sns_post": "【案1：データ】\n...\n\n【案2：共感】\n...\n\n【案3：提言】\n...",
    "authors": "著者名・組織名",
    "publication_date": "yyyy/mm/dd"
  }

  論文テキスト:
  ${text.substring(0, 30000)}
  `;

  return callGeminiAPI(apiUrl, prompt, false); 
}

// -------------------------------------------------------
// KEYデータに数値が十分含まれているか簡易チェック（% or n= or p or 95%CI or d= or 数字）
// -------------------------------------------------------
function keyEvidenceHasEnoughNumbers(keyEvidence) {
  if (!keyEvidence || typeof keyEvidence !== 'string') return false;
  var s = keyEvidence;
  // 数値らしきパターン: %, n=, p<, p=, 95%CI, d=, OR=, 小数点・整数
  var hasPercent = /\d+%|%\s*増加|%\s*減少|%\s*低下/.test(s);
  var hasN = /n\s*=\s*\d+|N\s*=\s*\d+|サンプル.*\d+/.test(s);
  var hasP = /p\s*[<>=].*?\d|p\s*<\s*0\.\d+/.test(s);
  var hasCI = /95%\s*CI|信頼区間|CI\s*[\[\(]/.test(s);
  var hasEffect = /d\s*=\s*[\d.]+|OR\s*=|RR\s*=|HR\s*=|効果量/.test(s);
  var hasPlainNumber = (s.match(/\d+\.?\d*/g) || []).length >= 3;
  return (hasPercent || hasN || hasP || hasCI || hasEffect) && hasPlainNumber;
}

// -------------------------------------------------------
// フォールバック：論文テキストから数値エビデンスのみを抽出（KEYデータに数値が足りないとき用）
// -------------------------------------------------------
function extractNumericalEvidenceOnly(paperText) {
  var apiUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + getConfig().apiKey;
  var prompt = `
以下の論文テキストから、**数値・統計が明記されている事実だけ**を箇条書きで抽出してください。
各項目には必ず次のいずれかを含めること: 割合(%), サンプル数(n=), p値, 信頼区間(95%CI), 効果量(d, OR, RR等), 平均・標準偏差(M±SD)。
「効果があった」「有意だった」など数値のない表現は一切含めないでください。5〜10項目、数値付きで出力すること。
出力はプレーンテキストの箇条書きのみ（JSON不要）。

論文テキスト:
${(paperText || '').substring(0, 20000)}
`;
  try {
    var payload = { contents: [{ parts: [{ text: prompt }] }] };
    var response = UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var data = JSON.parse(response.getContentText());
    if (data.error || !data.candidates || !data.candidates[0]) return '';
    var text = data.candidates[0].content.parts[0].text;
    return (text || '').replace(/^[\s\-・]*/gm, '・').trim();
  } catch (e) {
    return '';
  }
}

// -------------------------------------------------------
// 共通：Gemini API呼び出し・エラー処理関数
// -------------------------------------------------------
function callGeminiAPI(apiUrl, prompt, isArrayExpected) {
  try {
    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const response = UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());

    if (data.error) {
      console.error("🚨 Gemini API Error: " + data.error.message);
      return isArrayExpected ? [] : null;
    }
    
    if (!data.candidates || !data.candidates[0]) {
      console.error("🚨 回答が生成されませんでした");
      return isArrayExpected ? [] : null;
    }

    let text = data.candidates[0].content.parts[0].text;
    text = text.replace(/```json/g, "").replace(/```/g, "").trim();
    // Geminiが文字列値内に改行・タブなどをそのまま出すとJSON.parseが落ちるため、制御文字をスペースに置換
    text = text.replace(/[\x00-\x1f\x7f]/g, " ");
    
    const parsed = JSON.parse(text);
    return parsed;

  } catch (e) {
    console.error("プログラム内部エラー: " + e.toString());
    return isArrayExpected ? [] : null;
  }
}

// -------------------------------------------------------
// 関数3：中身を取ってくる
// -------------------------------------------------------
function fetchUrlContent(url) {
  try {
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    return response.getContentText().replace(/<[^>]*>/g, " ").substring(0, 30000);
  } catch (e) { return null; }
}

// -------------------------------------------------------
// 論文の原文URLを正しく解決する（重要：リンク誤り防止）
// Google Scholarの scholar_url?url=... 形式から実際の論文URLを抽出
// メール内のリンクがリダイレクトURLだと、クリック時に別論文に飛ぶことがある
// -------------------------------------------------------
function resolvePaperUrl(url) {
  if (!url || typeof url !== 'string') return url;
  try {
    // scholar_url?url= 形式（ScholarのリダイレクトURL）から実際の論文URLを抽出
    const scholarMatch = url.match(/scholar_url\?url=([^&]+)/);
    if (scholarMatch) {
      return decodeURIComponent(scholarMatch[1]);
    }
    // q= 形式の検索結果リンクの場合、HTMLから取得したURLを使う必要がある
    // ここでは抽出できないため、元のURLを返す
    return url;
  } catch (e) {
    return url;
  }
}

// -------------------------------------------------------
// 関数4：メール通知
// -------------------------------------------------------
function sendRichEmail(result, url) {
  const { emailRecipient } = getConfig();
  const authors = result.authors || '不明';
  const pubDate = result.publication_date || '不明';
  
  const htmlBody = `
    <div style="font-family: sans-serif; max-width: 600px; margin: 0 auto; color: #333;">
      <h2 style="color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 10px; margin-bottom: 20px;">${escapeHtml(result.title)}</h2>
      
      <div style="background-color: #e8f0fe; padding: 12px 15px; border-radius: 8px; margin-bottom: 20px; font-size: 0.9em; color: #1967d2;">
        <strong>執筆者:</strong> ${escapeHtml(authors)} &nbsp;|&nbsp; <strong>公開日:</strong> ${escapeHtml(pubDate)}
      </div>
      
      <div style="background-color: #f1f3f4; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
        <strong style="color: #202124;">■ 3行要約</strong>
        <p style="margin-top: 5px; line-height: 1.6;">${escapeHtml(result.summary_3lines)}</p>
      </div>

      <div style="margin-bottom: 25px;">
        <strong style="color: #202124; border-left: 4px solid #1a73e8; padding-left: 10px; font-size: 1.1em;">詳細要約（500文字）</strong>
        <p style="margin-top: 10px; line-height: 1.8; text-align: justify;">${escapeHtml(result.summary_long)}</p>
      </div>

      <div style="margin-bottom: 25px;">
         <div style="border: 1px solid #e0e0e0; border-radius: 8px; padding: 15px; margin-bottom: 10px;">
           <strong style="color: #d93025;">■ KEYデータ / エビデンス</strong>
           <p style="margin-top: 5px; font-size: 0.95em; white-space: pre-wrap;">${escapeHtml(result.key_evidence)}</p>
         </div>
         <div style="border: 1px solid #e0e0e0; border-radius: 8px; padding: 15px; background-color: #fafafa;">
           <strong style="color: #137333;">■ Credibility Check (信憑性)</strong>
           <p style="margin-top: 5px; font-size: 0.95em; white-space: pre-wrap;">${escapeHtml(credibilityToDisplayString(result.credibility))}</p>
         </div>
      </div>

      <div style="margin-bottom: 25px;">
        <strong style="color: #202124; border-left: 4px solid #d93025; padding-left: 10px; font-size: 1.1em;">日本への示唆・インサイト</strong>
        <p style="margin-top: 10px; line-height: 1.8;">${escapeHtml(result.japan_context)}</p>
      </div>

      <div style="border: 2px dashed #ddd; padding: 15px; border-radius: 8px; background-color: #fff;">
        <strong style="color: #5f6368;">SNS投稿案（3つの切り口）</strong>
        <p style="margin-top: 10px; font-family: sans-serif; background: #f9f9f9; padding: 15px; border-radius: 4px; white-space: pre-wrap; line-height: 1.8; color: #444;">${escapeHtml(result.sns_post)}</p>
      </div>
      
      <div style="margin-top: 30px; text-align: center;">
        <a href="${escapeHtml(url)}" style="background-color: #1a73e8; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold;">論文の原文を読む</a>
      </div>
    </div>
  `;
  
  GmailApp.sendEmail(emailRecipient, `【厳選レポート】${result.title}`, "", {htmlBody: htmlBody});
}

// credibility が Gemini からオブジェクトで返った場合に表示用文字列に変換する
function credibilityToDisplayString(cred) {
  if (cred === null || cred === undefined) return '';
  if (typeof cred === 'string') return cred;
  if (typeof cred === 'object') {
    var parts = [];
    for (var key in cred) {
      if (Object.prototype.hasOwnProperty.call(cred, key)) {
        var val = cred[key];
        if (val !== null && val !== undefined && val !== '') {
          parts.push(key + ': ' + String(val).trim());
        }
      }
    }
    return parts.length ? parts.join('\n\n') : '';
  }
  return String(cred);
}

// HTMLエスケープ（XSS対策）
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
