/**
 * Gemini for Microsoft 365 - Add-in Taskpane Controller
 * 
 * Manages the taskpane UI, chat history, selection toolbar, in-document triggers,
 * and host-adaptive document intelligence for Word, PowerPoint, and Excel.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { askGeminiEnterprise, getActiveProxyUrl, setProxyUrlOverride, getGoogleAccessToken, setGoogleAccessToken } from '../core/geminiClient.js';
import { parseMarkdown } from '../core/markdownParser.js';
import { HostAdapterFactory } from '../adapters/HostAdapterFactory.js';

let currentSessionId = null;
let chatHistoryState = [];
let isProcessingInDocCommand = false;
let hostAdapter = null;
let currentSelectedText = "";

Office.onReady((info) => {
  // Detect active Microsoft Office host (Word, PowerPoint, Excel) dynamically
  hostAdapter = HostAdapterFactory.getAdapter();

  document.getElementById("run").onclick = () => callGeminiProxy();
  const clearSessionBtn = document.getElementById("clearSession");
  if (clearSessionBtn) {
    clearSessionBtn.onclick = resetChatSession;
  }

  const scanBtn = document.getElementById("scanInDoc");
  if (scanBtn) {
    scanBtn.onclick = () => checkForInDocumentCommands(true);
  }

  // Google Enterprise Auth Panel Setup
  const authBtn = document.getElementById("authBtn");
  const authPanel = document.getElementById("authPanel");
  const closeAuthPanel = document.getElementById("closeAuthPanel");
  const googleTokenInput = document.getElementById("googleTokenInput");
  const saveTokenBtn = document.getElementById("saveTokenBtn");
  const clearTokenBtn = document.getElementById("clearTokenBtn");
  const tokenStatusMsg = document.getElementById("tokenStatusMsg");

  if (authBtn && authPanel) {
    const existingToken = getGoogleAccessToken();
    if (existingToken) {
      authBtn.style.color = '#107c41';
      authBtn.style.borderColor = '#107c41';
      authBtn.innerHTML = '🔑 Google (Auth)';
    }

    authBtn.onclick = () => {
      authPanel.style.display = authPanel.style.display === 'none' ? 'block' : 'none';
      if (googleTokenInput) {
        googleTokenInput.value = getGoogleAccessToken();
      }
    };
  }

  if (closeAuthPanel && authPanel) {
    closeAuthPanel.onclick = () => { authPanel.style.display = 'none'; };
  }

  if (saveTokenBtn) {
    saveTokenBtn.onclick = () => {
      const val = googleTokenInput ? googleTokenInput.value.trim() : '';
      setGoogleAccessToken(val);
      if (val) {
        if (tokenStatusMsg) {
          tokenStatusMsg.style.color = '#107c41';
          tokenStatusMsg.innerText = '✅ Google User OAuth Token Saved!';
        }
        if (authBtn) {
          authBtn.style.color = '#107c41';
          authBtn.style.borderColor = '#107c41';
          authBtn.innerHTML = '🔑 Google (Auth)';
        }
      } else {
        if (tokenStatusMsg) {
          tokenStatusMsg.style.color = '#605e5c';
          tokenStatusMsg.innerText = 'Token cleared.';
        }
        if (authBtn) {
          authBtn.style.color = '#323130';
          authBtn.style.borderColor = '#8a8886';
          authBtn.innerHTML = '🔑 Google Auth';
        }
      }
      setTimeout(() => { if (authPanel) authPanel.style.display = 'none'; }, 1000);
    };
  }

  if (clearTokenBtn) {
    clearTokenBtn.onclick = () => {
      setGoogleAccessToken('');
      if (googleTokenInput) googleTokenInput.value = '';
      if (tokenStatusMsg) {
        tokenStatusMsg.style.color = '#605e5c';
        tokenStatusMsg.innerText = 'Token cleared.';
      }
      if (authBtn) {
        authBtn.style.color = '#323130';
        authBtn.style.borderColor = '#8a8886';
        authBtn.innerHTML = '🔑 Google Auth';
      }
    };
  }

  const googleSignInBtn = document.getElementById("googleSignInBtn");
  if (googleSignInBtn) {
    googleSignInBtn.onclick = () => {
      try {
        if (typeof google === 'undefined' || !google.accounts || !google.accounts.oauth2) {
          if (tokenStatusMsg) {
            tokenStatusMsg.style.color = '#a4262c';
            tokenStatusMsg.innerText = 'Google Identity SDK loading... please wait a moment or paste token below.';
          }
          return;
        }

        const clientId = window.localStorage ? (window.localStorage.getItem('google_oauth_client_id') || '36841365232-oauth.apps.googleusercontent.com') : '36841365232-oauth.apps.googleusercontent.com';

        const client = google.accounts.oauth2.initTokenClient({
          client_id: clientId,
          scope: 'https://www.googleapis.com/auth/cloud-platform',
          callback: (tokenResponse) => {
            if (tokenResponse && tokenResponse.access_token) {
              setGoogleAccessToken(tokenResponse.access_token);
              if (googleTokenInput) googleTokenInput.value = tokenResponse.access_token;
              if (tokenStatusMsg) {
                tokenStatusMsg.style.color = '#107c41';
                tokenStatusMsg.innerText = '✅ Successfully Signed in with Google!';
              }
              if (authBtn) {
                authBtn.style.color = '#107c41';
                authBtn.style.borderColor = '#107c41';
                authBtn.innerHTML = '🔑 Google (Auth)';
              }
              setTimeout(() => { if (authPanel) authPanel.style.display = 'none'; }, 1200);
            } else if (tokenResponse && tokenResponse.error) {
              if (tokenStatusMsg) {
                tokenStatusMsg.style.color = '#a4262c';
                tokenStatusMsg.innerText = `Sign in: ${tokenResponse.error_description || tokenResponse.error}`;
              }
            }
          },
          error_callback: (err) => {
            if (tokenStatusMsg) {
              tokenStatusMsg.style.color = '#a4262c';
              tokenStatusMsg.innerText = `Sign in: ${err.message || 'Popup blocked or cancelled'}`;
            }
          }
        });

        client.requestAccessToken();
      } catch (err) {
        console.error('Google Sign in error:', err);
        if (tokenStatusMsg) {
          tokenStatusMsg.style.color = '#a4262c';
          tokenStatusMsg.innerText = `Sign in: ${err.message}`;
        }
      }
    };
  }

  // Target Proxy Endpoint selector setup
  const endpointSelect = document.getElementById("endpointSelect");
  if (endpointSelect) {
    const savedUrl = window.localStorage ? window.localStorage.getItem('gemini_proxy_url') : "";
    if (savedUrl) {
      endpointSelect.value = savedUrl;
    }
    endpointSelect.onchange = (e) => {
      const selectedUrl = e.target.value;
      setProxyUrlOverride(selectedUrl);
      const debugStatus = document.getElementById("debugStatus");
      if (debugStatus) {
        debugStatus.innerText = selectedUrl ? `Target: Override Active` : `${hostAdapter ? hostAdapter.name : 'Office'} Ready`;
      }
    };
  }

  // Adapt UI titles, top banners, and action chips to the active Microsoft product
  adaptUIForHost(hostAdapter.name);

  // Setup Document Intelligence Chips (Feature 1: Full Document Q&A)
  setupDocToolsChips();

  // Setup Selection Quick Toolbar Chips (Feature 3: Inline Rewrite Toolbar)
  setupSelectionChips();

  // Attach selection change handler for in-document detection & selection toolbar
  try {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => handleSelectionChanged()
    );
  } catch (e) {
    console.warn("Could not attach selection handler:", e);
  }
});

function adaptUIForHost(hostName) {
  const tipBannerText = document.getElementById("tipBannerText");
  const docToolsTitle = document.getElementById("docToolsTitle");
  const welcomeBubble = document.getElementById("welcomeSystemBubble");
  const debugStatus = document.getElementById("debugStatus");
  const chipSummarize = document.getElementById("chipSummarize");
  const chipRisks = document.getElementById("chipRisks");
  const chipActionItems = document.getElementById("chipActionItems");
  const chipExecBox = document.getElementById("chipExecBox");

  if (hostName === "PowerPoint") {
    if (tipBannerText) tipBannerText.innerHTML = `💡 Type <code>@gemini &lt;prompt&gt;</code> in Slide`;
    if (docToolsTitle) docToolsTitle.innerHTML = `📊 <strong>Chat with Slides:</strong>`;
    if (debugStatus) debugStatus.innerText = `PowerPoint Ready`;
    if (chipSummarize) chipSummarize.innerHTML = `📊 Summarize Slides`;
    if (chipRisks) chipRisks.innerHTML = `⚠️ Key Risks`;
    if (chipActionItems) chipActionItems.innerHTML = `✅ Action Items`;
    if (chipExecBox) chipExecBox.innerHTML = `🎯 Slide Takeaways`;
    if (welcomeBubble) welcomeBubble.innerHTML = `Type a prompt below, click a <strong>Slide Chip</strong> above, or select shapes in PowerPoint!`;
  } else if (hostName === "Excel") {
    if (tipBannerText) tipBannerText.innerHTML = `💡 Type <code>@gemini &lt;prompt&gt;</code> in Cell`;
    if (docToolsTitle) docToolsTitle.innerHTML = `📈 <strong>Chat with Spreadsheet:</strong>`;
    if (debugStatus) debugStatus.innerText = `Excel Ready`;
    if (chipSummarize) chipSummarize.innerHTML = `📈 Summarize Sheet`;
    if (chipRisks) chipRisks.innerHTML = `⚠️ Data & Formula Risks`;
    if (chipActionItems) chipActionItems.innerHTML = `✅ Action Items`;
    if (chipExecBox) chipExecBox.innerHTML = `💡 Key Metrics Card`;
    if (welcomeBubble) welcomeBubble.innerHTML = `Type a prompt below, click a <strong>Sheet Chip</strong> above, or select cells in Excel!`;
  } else {
    // Word (Default)
    if (tipBannerText) tipBannerText.innerHTML = `💡 Type <code>@gemini &lt;prompt&gt;</code> in Doc`;
    if (docToolsTitle) docToolsTitle.innerHTML = `📄 <strong>Chat with Document:</strong>`;
    if (debugStatus) debugStatus.innerText = `Word Ready`;
    if (chipSummarize) chipSummarize.innerHTML = `📄 Summarize Doc`;
    if (chipRisks) chipRisks.innerHTML = `⚠️ Key Risks`;
    if (chipActionItems) chipActionItems.innerHTML = `✅ Action Items`;
    if (chipExecBox) chipExecBox.innerHTML = `💡 Executive Card`;
    if (welcomeBubble) welcomeBubble.innerHTML = `Type a prompt below, click a <strong>Doc Chip</strong> above, or highlight text in Word!`;
  }

  // Detect Gemini Enterprise StreamAssist Mode
  if (typeof window !== 'undefined' && window.location) {
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('backend') === 'streamassist') {
      const titleEl = document.querySelector(".header h3");
      if (titleEl) {
        titleEl.innerHTML = `✨ Gemini Enterprise <span class="header-badge" style="background:#e8f5e9; color:#1b5e20;">🍔 Wendy's KB</span>`;
      }
      if (welcomeBubble) {
        welcomeBubble.innerHTML = `Connected to <strong>Wendy's Gemini Enterprise Assistant</strong> (HR, Sales, SharePoint, Box, ServiceNow). Answers are grounded on Wendy's internal records with citations!`;
      }
      const promptInput = document.getElementById("promptText");
      if (promptInput) {
        promptInput.placeholder = "Ask Wendy's Enterprise Knowledge Base (HR, Sales, SharePoint)...";
      }
      if (debugStatus) {
        debugStatus.innerText = `Enterprise: Wendy's (${hostName})`;
      }
    }
  }
}

function resetChatSession() {
  currentSessionId = null;
  chatHistoryState = [];
  const historyDiv = document.getElementById("chatHistory");
  if (historyDiv) {
    historyDiv.innerHTML = '<div class="chat-bubble system">Chat session reset. Ready for a new topic!</div>';
  }
  const debugStatus = document.getElementById("debugStatus");
  if (debugStatus) debugStatus.innerText = `${hostAdapter ? hostAdapter.name : 'Office'} Ready (Session Reset)`;
}

// Feature 1: Document Intelligence Quick Chips Handlers
function setupDocToolsChips() {
  const hostName = hostAdapter ? hostAdapter.name : "Word";

  const chipSummarize = document.getElementById("chipSummarize");
  if (chipSummarize) {
    chipSummarize.onclick = () => {
      if (hostName === "PowerPoint") {
        runDocIntelligencePrompt("Summarize these presentation slides thoroughly. Provide an executive overview of all slides, key themes, and a structured summary table.");
      } else if (hostName === "Excel") {
        runDocIntelligencePrompt("Analyze and summarize this spreadsheet data. Provide key patterns, trends, data anomalies, and an executive summary table.");
      } else {
        runDocIntelligencePrompt("Summarize this document thoroughly. Provide an executive overview, key takeaways, and a structured markdown summary table.");
      }
    };
  }

  const chipRisks = document.getElementById("chipRisks");
  if (chipRisks) {
    chipRisks.onclick = () => {
      if (hostName === "PowerPoint") {
        runDocIntelligencePrompt("Analyze these presentation slides and extract all Key Strategic Risks, Challenges, and Operational Gaps with Mitigation Suggestions.");
      } else if (hostName === "Excel") {
        runDocIntelligencePrompt("Analyze this spreadsheet data and extract all Key Financial / Operational Risks, Outliers, and Data Gaps with Mitigation Suggestions.");
      } else {
        runDocIntelligencePrompt("Analyze this document and extract all Key Risks, Ambiguities, and Operational Gaps. Present them in a structured markdown table with Severity and Mitigation Suggestions.");
      }
    };
  }

  const chipActionItems = document.getElementById("chipActionItems");
  if (chipActionItems) {
    chipActionItems.onclick = () => {
      const noun = hostName === "PowerPoint" ? "presentation deck" : (hostName === "Excel" ? "spreadsheet" : "document");
      runDocIntelligencePrompt(`Extract all Action Items, Deliverables, and Next Steps from this ${noun}. Present them in a structured table with Task, Owner, and Priority.`);
    };
  }

  const chipExecBox = document.getElementById("chipExecBox");
  if (chipExecBox) {
    chipExecBox.onclick = () => {
      if (hostName === "PowerPoint") {
        runDocIntelligencePrompt("Generate an Executive Slide Takeaway Callout Box highlighting Strategic Impact and Key Metrics for this deck.");
      } else if (hostName === "Excel") {
        runDocIntelligencePrompt("Generate an Executive Metrics Summary Card highlighting Key Financial / Operational KPIs and totals from this sheet.");
      } else {
        runDocIntelligencePrompt("Generate an Executive Summary Box for this document using a callout block (> [!NOTE] ...) highlighting Strategic Purpose, Key Metrics, and Impact.");
      }
    };
  }
}

// Feature 3: Selection Quick Toolbar Setup
function setupSelectionChips() {
  const selRewrite = document.getElementById("selRewrite");
  if (selRewrite) {
    selRewrite.onclick = () => runSelectionPrompt("Rewrite and polish the selected text to be clearer, punchier, and more compelling while preserving its original meaning.");
  }

  const selProfessional = document.getElementById("selProfessional");
  if (selProfessional) {
    selProfessional.onclick = () => runSelectionPrompt("Elevate the tone of the selected text to executive, professional business prose.");
  }

  const selShorten = document.getElementById("selShorten");
  if (selShorten) {
    selShorten.onclick = () => runSelectionPrompt("Make the selected text significantly more concise and punchy, removing unnecessary fluff.");
  }

  const selBullets = document.getElementById("selBullets");
  if (selBullets) {
    selBullets.onclick = () => runSelectionPrompt("Convert the selected text into clean, high-impact bullet points.");
  }

  const selTable = document.getElementById("selTable");
  if (selTable) {
    selTable.onclick = () => runSelectionPrompt("Convert the information in the selected text into a structured markdown table.");
  }
}

// Handle Word Selection Changes Dynamically
async function handleSelectionChanged() {
  if (!hostAdapter) return;
  const selToolbar = document.getElementById("selectionToolbar");
  const previewText = document.getElementById("selectionPreviewText");
  const wordCountSpan = document.getElementById("selectionWordCount");

  try {
    const text = await hostAdapter.getSelectedText();
    currentSelectedText = text ? text.trim() : "";

    if (currentSelectedText.length > 5) {
      const words = currentSelectedText.split(/\s+/).filter(w => w.length > 0).length;
      if (wordCountSpan) wordCountSpan.innerText = `${words} words`;
      if (previewText) {
        previewText.innerText = `"${currentSelectedText.substring(0, 80)}${currentSelectedText.length > 80 ? '...' : ''}"`;
      }
      if (selToolbar) selToolbar.style.display = "block";
    } else {
      if (selToolbar) selToolbar.style.display = "none";
    }

    // Check for in-document @gemini command
    await checkForInDocumentCommands(false);
  } catch (err) {
    console.warn("Selection change handler error:", err);
  }
}

// Universal In-Document / In-App Command Processor (@gemini <prompt>)
async function checkForInDocumentCommands(forceRun = false) {
  if (isProcessingInDocCommand || !hostAdapter || typeof hostAdapter.checkInDocumentCommands !== 'function') return;

  const debugStatus = document.getElementById("debugStatus");
  const scanBtn = document.getElementById("scanInDoc");
  const loadingText = document.getElementById("loading");

  await hostAdapter.checkInDocumentCommands(forceRun, {
    onStatus: (msg) => {
      if (debugStatus) debugStatus.innerText = msg;
    },
    executePrompt: async (userPrompt) => {
      isProcessingInDocCommand = true;
      if (scanBtn) scanBtn.disabled = true;
      if (loadingText) {
        loadingText.innerText = "⚡ Generating @gemini response...";
        loadingText.style.display = "block";
      }

      appendBubble(`@gemini ${userPrompt}`, "user");

      try {
        const data = await askGeminiEnterprise(userPrompt, chatHistoryState, currentSessionId);
        if (data.sessionId) currentSessionId = data.sessionId;
        if (Array.isArray(data.history)) chatHistoryState = data.history;

        const aiResultText = data.result || "No content returned.";
        appendAssistantBubble(aiResultText);

        return parseMarkdown(aiResultText);
      } catch (err) {
        console.error("In-document command execution error:", err);
        if (debugStatus) debugStatus.innerText = "Error: " + err.message;
        return null;
      } finally {
        isProcessingInDocCommand = false;
        if (scanBtn) scanBtn.disabled = false;
        if (loadingText) {
          loadingText.style.display = "none";
        }
      }
    }
  });
}

async function runSelectionPrompt(instruction) {
  if (!currentSelectedText) {
    currentSelectedText = await hostAdapter.getSelectedText();
  }

  if (!currentSelectedText) {
    appendBubble("Please highlight text in Word first.", "system");
    return;
  }

  const fullPrompt = `Selected Document Text:\n"${currentSelectedText}"\n\nTask: ${instruction}`;
  const displayUserBubble = `✨ ${instruction}\n📌 Context: "${currentSelectedText.substring(0, 70)}..."`;

  await executeGeminiWorkflow(fullPrompt, displayUserBubble);
}

async function runDocIntelligencePrompt(instruction) {
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Reading document context...";
    loadingText.style.display = "block";
  }

  const fullDocText = await hostAdapter.getFullDocumentText();
  
  let fullPrompt = "";
  if (fullDocText && fullDocText.length > 20) {
    fullPrompt = `Full Document Text Context:\n"""\n${fullDocText.substring(0, 100000)}\n"""\n\nUser Instruction: ${instruction}`;
  } else {
    fullPrompt = instruction;
  }

  const displayUserBubble = `📄 [Document Analysis] ${instruction.substring(0, 55)}...`;
  await executeGeminiWorkflow(fullPrompt, displayUserBubble);
}

async function callGeminiProxy(customPrompt = null) {
  const promptInput = document.getElementById("promptText");
  const userText = customPrompt || (promptInput ? promptInput.value.trim() : "");

  let selectedText = await hostAdapter.getSelectedText();

  if (!userText && !selectedText) {
    appendBubble(`Please type a prompt or select content in ${hostAdapter.name} first.`, "system");
    return;
  }

  let fullPrompt = "";
  let displayUserBubble = "";

  if (selectedText && userText) {
    fullPrompt = `Selected Document Context:\n"${selectedText}"\n\nUser Instruction: ${userText}`;
    displayUserBubble = `📌 Context: "${selectedText.substring(0, 70)}${selectedText.length > 70 ? '...' : ''}"\n\n${userText}`;
  } else if (selectedText && !userText) {
    fullPrompt = `Please analyze, summarize, or explain the following selected text:\n"${selectedText}"`;
    displayUserBubble = `📌 Selected Text:\n"${selectedText.substring(0, 90)}${selectedText.length > 90 ? '...' : ''}"`;
  } else {
    fullPrompt = userText;
    displayUserBubble = userText;
  }

  if (promptInput) promptInput.value = "";
  await executeGeminiWorkflow(fullPrompt, displayUserBubble);
}

async function executeGeminiWorkflow(fullPrompt, displayUserBubble) {
  const runButton = document.getElementById("run");
  const loadingText = document.getElementById("loading");
  const historyDiv = document.getElementById("chatHistory");

  if (runButton) runButton.disabled = true;
  if (loadingText) {
    const isEnterprise = typeof window !== 'undefined' && window.location && new URLSearchParams(window.location.search).get('backend') === 'streamassist';
    loadingText.innerText = isEnterprise ? "⚡ Gemini Enterprise is thinking..." : "⚡ Gemini Guru is thinking...";
    loadingText.style.display = "block";
  }

  try {
    appendBubble(displayUserBubble, "user");

    const data = await askGeminiEnterprise(fullPrompt, chatHistoryState, currentSessionId);

    if (data.sessionId) {
      currentSessionId = data.sessionId;
    }
    if (Array.isArray(data.history)) {
      chatHistoryState = data.history;
    }

    const aiResponse = data.result || "No content returned.";
    appendAssistantBubble(aiResponse);

  } catch (error) {
    appendBubble("Error: " + error.message, "system");
  } finally {
    if (runButton) runButton.disabled = false;
    if (loadingText) loadingText.style.display = "none";
    if (historyDiv) historyDiv.scrollTop = historyDiv.scrollHeight;
  }
}

function appendBubble(text, type) {
  const historyDiv = document.getElementById("chatHistory");
  if (!historyDiv) return;
  const bubble = document.createElement("div");
  bubble.className = `chat-bubble ${type}`;
  bubble.innerText = text;
  historyDiv.appendChild(bubble);
  historyDiv.scrollTop = historyDiv.scrollHeight;
}

function appendAssistantBubble(text) {
  const historyDiv = document.getElementById("chatHistory");
  if (!historyDiv) return;
  const bubble = document.createElement("div");
  bubble.className = "chat-bubble assistant";

  // Parse markdown into executive HTML
  const formattedHtml = parseMarkdown(text);

  const textDiv = document.createElement("div");
  textDiv.innerHTML = formattedHtml;
  bubble.appendChild(textDiv);

  // Feature 2 & Multi-turn: Action Toolbar & Refinement Chips
  const actionsContainer = document.createElement("div");
  actionsContainer.className = "response-actions-container";

  // Primary Actions: Replace / Insert / Copy
  const primaryActions = document.createElement("div");
  primaryActions.className = "primary-actions";

  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const isPPT = hostName === "PowerPoint";
  const isExcel = hostName === "Excel";

  // 1. In-Place Replace Button
  const replaceBtn = document.createElement("button");
  replaceBtn.className = "action-btn replace";
  replaceBtn.innerHTML = isPPT ? `🔄 Replace Slides` : (isExcel ? `🔄 Replace in Sheet` : `🔄 Replace in Doc`);
  replaceBtn.title = isPPT ? "Replace existing presentation slides" : "Replace active draft or selection in Word";
  replaceBtn.onclick = async () => {
    await performDocumentInsertion(textDiv.innerHTML, text, "replace_draft");
  };

  // 2. Insert Button
  const insertBtn = document.createElement("button");
  insertBtn.className = "action-btn insert";
  insertBtn.innerHTML = isPPT ? `➕ Insert into Slides` : (isExcel ? `➕ Insert into Sheet` : `➕ Insert at Cursor`);
  insertBtn.title = isPPT ? "Create new presentation slides" : "Insert at current cursor location";
  insertBtn.onclick = async () => {
    await performDocumentInsertion(textDiv.innerHTML, text, "insert_cursor");
  };

  // 3. Copy Button
  const copyBtn = document.createElement("button");
  copyBtn.className = "action-btn copy";
  copyBtn.innerHTML = `📋 Copy`;
  copyBtn.title = "Copy to clipboard";
  copyBtn.onclick = async () => {
    try {
      await navigator.clipboard.writeText(text);
      copyBtn.innerHTML = `✅ Copied!`;
      setTimeout(() => { copyBtn.innerHTML = `📋 Copy`; }, 2000);
    } catch (e) {
      console.warn("Clipboard copy error:", e);
    }
  };

  primaryActions.appendChild(replaceBtn);
  primaryActions.appendChild(insertBtn);
  primaryActions.appendChild(copyBtn);
  actionsContainer.appendChild(primaryActions);

  // Refinement Chips: Quick 1-Click Multi-Turn Prompts
  const chipsLabel = document.createElement("div");
  chipsLabel.className = "refinement-chips-label";
  chipsLabel.innerText = "Refine Draft:";
  actionsContainer.appendChild(chipsLabel);

  const refinementChips = document.createElement("div");
  refinementChips.className = "refinement-chips";

  const chipsData = [
    { label: "📉 Make Shorter", prompt: "Make the above response significantly more concise and punchy for executive reading." },
    { label: "📈 Expand Details", prompt: "Expand the above draft with more in-depth technical, operational, and architectural details." },
    { label: "📊 Format as Table", prompt: "Convert the key findings and aspects of the above response into a structured markdown table." },
    { label: "👔 Executive Tone", prompt: "Rewrite the above response with an authoritative, C-level executive tone." },
    { label: "🔄 Try Again", prompt: "Regenerate the response with a fresh structure and alternative perspective." }
  ];

  chipsData.forEach(item => {
    const chip = document.createElement("button");
    chip.className = "refinement-chip";
    chip.innerText = item.label;
    chip.onclick = () => {
      executeGeminiWorkflow(item.prompt, `${item.label}: "${item.prompt.substring(0, 45)}..."`);
    };
    refinementChips.appendChild(chip);
  });

  actionsContainer.appendChild(refinementChips);
  bubble.appendChild(actionsContainer);

  historyDiv.appendChild(bubble);
  historyDiv.scrollTop = historyDiv.scrollHeight;
}

async function performDocumentInsertion(htmlContent, rawText, mode = "smart") {
  const runButton = document.getElementById("run");
  const loadingText = document.getElementById("loading");

  if (runButton) runButton.disabled = true;
  if (loadingText) {
    loadingText.innerText = hostAdapter?.name === 'PowerPoint' ? "⚡ Creating PowerPoint slides..." : "⚡ Updating document...";
    loadingText.style.display = "block";
  }

  try {
    const insertPromise = hostAdapter.insertContent(htmlContent, rawText);
    const timeoutPromise = new Promise((_, reject) => 
      setTimeout(() => reject(new Error("Operation timed out after 60 seconds")), 60000)
    );

    await Promise.race([insertPromise, timeoutPromise]);

    const debugStatus = document.getElementById("debugStatus");
    if (debugStatus) {
      const host = hostAdapter?.name || 'Office';
      debugStatus.innerText = `Updated in ${host} (${mode === 'replace_draft' ? 'Replaced' : 'Inserted'})`;
    }
  } catch (err) {
    console.error("Document insertion error:", err);
    appendBubble(`🔴 Insertion Notice: ${err.message || err}`, "system");
  } finally {
    if (runButton) runButton.disabled = false;
    if (loadingText) loadingText.style.display = "none";
  }
}

