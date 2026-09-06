/**
 * Gemini for Microsoft 365 - Add-in Taskpane Controller
 * 
 * Manages the taskpane UI, chat history, selection toolbar, in-document triggers,
 * and host-adaptive document intelligence for Word, PowerPoint, and Excel.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { askGeminiEnterprise, getActiveProxyUrl, setProxyUrlOverride } from '../core/geminiClient.js';
import { 
  getOfficeAuthToken, 
  getUserProfile, 
  getLastAuthError, 
  initiateGoogleSignIn, 
  isGoogleTokenValid, 
  getGoogleAccessToken,
  setGoogleAccessToken,
  getGoogleOAuthClientId,
  fetchAppConfig
} from '../core/authService.js';
import { parseMarkdown } from '../core/markdownParser.js';
import { HostAdapterFactory } from '../adapters/HostAdapterFactory.js';
import { initPowerPointDiagnostics } from '../adapters/ppt/pptDiagnostics.js';

let currentSessionId = null;
let chatHistoryState = [];
let isProcessingInDocCommand = false;
let hostAdapter = null;
let currentSelectedText = "";

Office.onReady(async (info) => {
  // Detect active Microsoft Office host (Word, PowerPoint, Excel) dynamically
  hostAdapter = HostAdapterFactory.getAdapter();

  // Pre-fetch dynamic backend configuration (Google OAuth Client ID)
  fetchAppConfig().catch(e => console.warn("Background config fetch failed:", e));

  // Initialize Entra ID & Google Drive Identity in UI
  await initAuthUI();

  // Initialize Collapsible Troubleshooting & Diagnostics Panel
  initTroubleshootPanel();

  // Log Add-in startup and host context to Cloud Logging
  sendDiagnosticLogToCloud(
    `Office Add-in initialized on host '${info.host || 'Unknown'}' (${info.platform || 'Unknown platform'})`,
    "INFO",
    "STARTUP",
    { host: info.host, platform: info.platform }
  );

  // Wire Google Drive 1-click connect button (used in GSuite / Cloud Identity mode)
  const googleDriveBtn = document.getElementById("googleDriveBtn");
  if (googleDriveBtn) {
    googleDriveBtn.onclick = async () => {
      const config = await fetchAppConfig();
      const isWifMode = config?.user_auth_mode === 'wif' || (config?.user_auth_mode === 'auto' && !config?.google_oauth_client_id);
      if (isWifMode) {
        console.log("WIF SSO active: Google OAuth login not required.");
        return;
      }
      console.log("Triggering Google OAuth 3-legged sign-in flow...");
      googleDriveBtn.innerHTML = "⏳ Logging in...";
      const profile = getUserProfile();
      const res = await initiateGoogleSignIn(profile?.email || null, 'select_account');
      if (res.status === 'success') {
        appendBubble("✅ Google login successful! Gemini Enterprise grounding is now active.", "system");
      } else {
        console.warn("Google Sign-In was not completed:", res.error);
      }
      await initAuthUI();
    };
  }

  document.getElementById("run").onclick = () => callGeminiProxy();
  
  const promptText = document.getElementById("promptText");
  if (promptText) {
    promptText.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        callGeminiProxy();
      }
    });
  }

  const clearSessionBtn = document.getElementById("clearSession");
  if (clearSessionBtn) {
    clearSessionBtn.onclick = resetChatSession;
  }

  const scanBtn = document.getElementById("scanInDoc");
  if (scanBtn) {
    scanBtn.onclick = () => checkForInDocumentCommands(true);
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

  if (hostAdapter.name === "PowerPoint") {
    initPowerPointDiagnostics();
  }

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
  // Wire interactive sign-in click on user profile badge
  const userAuthBadge = document.getElementById("userAuthBadge");
  const userStatusDot = document.getElementById("userStatusDot");
  const userProfileBar = document.getElementById("userProfileBar");
  
  const handleAuthClick = async () => {
    console.log("Triggering explicit Office Entra ID SSO sign-in...");
    if (userAuthBadge) userAuthBadge.innerText = "Signing in...";
    const token = await getOfficeAuthToken(true);
    await initAuthUI();
    const lastErr = getLastAuthError();
    if (!token && lastErr) {
      console.warn("Explicit sign-in attempt did not yield token:", lastErr);
      if (lastErr.code === 13007) {
        alert("Office SSO Error 13007: Application ID URI mismatch or client app not authorized in Entra ID.\n\n" +
              "1. In Entra ID App '85fb5428-6249-4131-9eeb-f2436d5d4d8c' -> Expose an API:\n" +
              "   Set App ID URI: api://gemini-frontend-16933400417.us-central1.run.app/85fb5428-6249-4131-9eeb-f2436d5d4d8c\n" +
              "   Authorized Client IDs:\n" +
              "   - 00000002-0000-0ff1-ce00-000000000000 (Office Desktop)\n" +
              "   - ea5a67f6-b6f3-4338-b240-c655ddc3cc8e (Office Web)\n" +
              "   - d3590ed6-52b3-4102-aeff-aad2292ab01c (Office Web / WAC)\n" +
              "2. Restart Office.");
      } else if (lastErr.code === 13001) {
        alert("Office SSO Error 13001: You are not currently signed into Microsoft Office with a corporate Microsoft Entra ID account.");
      } else if (lastErr.code === 13002) {
        alert("Office SSO Error 13002: Sign-in or consent was cancelled.");
      } else if (lastErr.code === 13012) {
        alert("Office SSO Error 13012: SSO API is not supported on this platform version or requires Office restart.");
      } else if (lastErr.code) {
        alert(`Office SSO Code ${lastErr.code}: ${lastErr.message || JSON.stringify(lastErr)}`);
      }
    }
  };

  if (userAuthBadge) userAuthBadge.onclick = handleAuthClick;
  if (userStatusDot) userStatusDot.onclick = handleAuthClick;
  if (userProfileBar) {
    userProfileBar.style.cursor = "pointer";
    userProfileBar.onclick = handleAuthClick;
  }
});

async function initAuthUI() {
  const userAuthBar = document.getElementById("userAuthBar");
  const userEmailText = document.getElementById("userEmailText");
  const userStatusDot = document.getElementById("userStatusDot");
  const googleDriveBtn = document.getElementById("googleDriveBtn");

  try {
    const token = await getOfficeAuthToken();
    const profile = getUserProfile();
    const lastErr = getLastAuthError();
    const config = await fetchAppConfig();
    const isWifMode = config?.user_auth_mode === 'wif' || (config?.user_auth_mode === 'auto' && !config?.google_oauth_client_id);

    // 1. Entra ID / Microsoft 365 status & WIF status
    if (token && profile.is_authenticated) {
      if (userEmailText) {
        userEmailText.innerText = profile.email || profile.name;
        userEmailText.className = "user-email-text";
        userEmailText.title = isWifMode
          ? `WIF SSO Active (Connected as ${profile.email || profile.name})`
          : `Connected to Microsoft 365 as ${profile.email || profile.name} (Tenant: ${profile.tenant_id || 'Entra ID'})`;
      }
      if (userStatusDot) {
        userStatusDot.className = "user-status-dot";
        userStatusDot.title = isWifMode
          ? `WIF SSO Active (Connected as ${profile.email || profile.name})`
          : `Connected to Microsoft 365 as ${profile.email || profile.name}`;
      }
      if (userAuthBar) {
        userAuthBar.title = isWifMode
          ? `WIF SSO Active (Connected as ${profile.email || profile.name})`
          : `Microsoft 365 Identity: ${profile.email || profile.name}`;
      }
    } else {
      if (userEmailText) {
        userEmailText.innerText = profile.email && profile.email !== 'user@organization.com' 
          ? profile.email 
          : "Enterprise Active";
        userEmailText.className = "user-email-text";
        userEmailText.title = "Gemini Enterprise ready (Service Account fallback active)";
      }
      if (userStatusDot) {
        userStatusDot.className = "user-status-dot";
        userStatusDot.title = "Gemini Enterprise active";
      }
      if (userAuthBar) {
        userAuthBar.title = "Gemini Enterprise connected via Cloud Run";
      }
    }

    // 2. Google OAuth button (Used when configured, or badge when in service account mode)
    if (googleDriveBtn) {
      if (isWifMode) {
        googleDriveBtn.style.display = "none";
      } else if (config?.google_oauth_client_id) {
        googleDriveBtn.style.display = "inline-flex";
        if (isGoogleTokenValid()) {
          googleDriveBtn.className = "google-drive-btn connected";
          googleDriveBtn.innerHTML = "✅ Google Connected";
          googleDriveBtn.title = "Gemini Enterprise grounding active (OAuth token valid)";
          googleDriveBtn.style.cursor = "pointer";
        } else {
          googleDriveBtn.className = "google-drive-btn";
          googleDriveBtn.innerHTML = "Login with Google";
          googleDriveBtn.title = "Click to sign into Google for personalized Drive search";
          googleDriveBtn.style.cursor = "pointer";
        }
      } else {
        googleDriveBtn.style.display = "inline-flex";
        googleDriveBtn.className = "google-drive-btn connected";
        googleDriveBtn.innerHTML = "🏢 Enterprise Grounded";
        googleDriveBtn.title = "Grounded in corporate Gemini Enterprise data stores";
        googleDriveBtn.style.cursor = "default";
        googleDriveBtn.onclick = null;
      }
    }
  } catch (err) {
    console.warn("Auth UI init error:", err);
    if (userEmailText) {
      userEmailText.innerText = "Unauthenticated";
      userEmailText.className = "user-email-text offline";
      userEmailText.title = "Auth error: " + (err.message || String(err));
    }
    if (userStatusDot) {
      userStatusDot.className = "user-status-dot offline";
      userStatusDot.title = "Auth error: " + (err.message || String(err));
    }
  } finally {
    updateDiagnosticsPanel().catch(() => {});
  }
}

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
}

function resetChatSession() {
  currentSessionId = null;
  chatHistoryState = [];
  const historyDiv = document.getElementById("chatHistory");
  if (historyDiv) {
    historyDiv.innerHTML = `
      <div class="empty-state" id="emptyState">
        <div class="empty-state-icon">✨</div>
        <div class="empty-state-title" id="emptyStateTitle">Gemini Enterprise Assistant</div>
        <div class="empty-state-sub" id="emptyStateSub">Ask questions, create content, or pick a starter prompt:</div>
        <div class="empty-state-chips">
          <button class="quick-chip" id="chipSummarize">📊 Summarize</button>
          <button class="quick-chip" id="chipRisks">⚠️ Key Risks</button>
          <button class="quick-chip" id="chipActionItems">✅ Action Items</button>
          <button class="quick-chip" id="chipExecBox">🎯 Key Takeaways</button>
        </div>
        <div class="empty-state-hint">💡 Tip: Type <code>@gemini &lt;prompt&gt;</code> directly in your document</div>
      </div>
    `;
    setupDocToolsChips();
    adaptUIForHost(hostAdapter ? hostAdapter.name : "Word");
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
  
  if (hostAdapter.name === "PowerPoint") {
    const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
    fullPrompt = enhancePromptForPowerPoint(fullPrompt);
  }

  await executeGeminiWorkflow(fullPrompt, displayUserBubble);
}

async function executeGeminiWorkflow(fullPrompt, displayUserBubble) {
  const runButton = document.getElementById("run");
  const loadingText = document.getElementById("loading");
  const historyDiv = document.getElementById("chatHistory");

  if (runButton) runButton.disabled = true;
  if (loadingText) {
    loadingText.innerText = "⚡ Gemini Enterprise is thinking...";
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
  const emptyState = document.getElementById("emptyState");
  if (emptyState) emptyState.style.display = "none";

  const historyDiv = document.getElementById("chatHistory");
  if (!historyDiv) return;
  const bubble = document.createElement("div");
  bubble.className = `chat-bubble ${type}`;
  bubble.innerText = text;
  historyDiv.appendChild(bubble);
  historyDiv.scrollTop = historyDiv.scrollHeight;
}

function appendAssistantBubble(text) {
  const emptyState = document.getElementById("emptyState");
  if (emptyState) emptyState.style.display = "none";

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

  // Primary Actions: Insert (Primary) / Copy / Replace
  const primaryActions = document.createElement("div");
  primaryActions.className = "primary-actions";

  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const isPPT = hostName === "PowerPoint";
  const isExcel = hostName === "Excel";

  // 1. Insert Button (Highlighted primary action)
  const insertBtn = document.createElement("button");
  insertBtn.className = "action-btn insert";
  insertBtn.innerHTML = isPPT ? `➕ Insert into Slides` : (isExcel ? `➕ Insert into Sheet` : `➕ Insert at Cursor`);
  insertBtn.title = isPPT ? "Create new presentation slides" : "Insert at current cursor location";
  insertBtn.onclick = async () => {
    await performDocumentInsertion(textDiv.innerHTML, text, "insert_cursor");
  };

  // 2. Copy Button
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

  // 3. In-Place Replace Button
  const replaceBtn = document.createElement("button");
  replaceBtn.className = "action-btn replace";
  replaceBtn.innerHTML = isPPT ? `🔄 Replace Slides` : (isExcel ? `🔄 Replace in Sheet` : `🔄 Replace in Doc`);
  replaceBtn.title = isPPT ? "Replace existing presentation slides" : "Replace active draft or selection in Word";
  replaceBtn.onclick = async () => {
    await performDocumentInsertion(textDiv.innerHTML, text, "replace_draft");
  };

  primaryActions.appendChild(insertBtn);
  primaryActions.appendChild(copyBtn);
  primaryActions.appendChild(replaceBtn);
  actionsContainer.appendChild(primaryActions);

  // Refinement Chips: Collapsible to save vertical screen real estate
  const refineDetails = document.createElement("details");
  refineDetails.className = "refine-details";

  const refineSummary = document.createElement("summary");
  refineSummary.className = "refine-summary";
  refineSummary.innerHTML = `✨ Refine draft ▾`;
  refineDetails.appendChild(refineSummary);

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

  refineDetails.appendChild(refinementChips);
  actionsContainer.appendChild(refineDetails);
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
    await hostAdapter.insertContent(htmlContent, rawText);
    const debugStatus = document.getElementById("debugStatus");
    if (debugStatus) {
      const host = hostAdapter?.name || 'Office';
      debugStatus.innerText = `Updated in ${host} (${mode === 'replace_draft' ? 'Replaced' : 'Inserted'})`;
    }
  } catch (err) {
    console.error("Document insertion error:", err);
    const diagDetail = err.debugInfo?.errorLocation || err.debugInfo?.statement || err.code || "";
    appendBubble(`🔴 Insertion Error: ${err.message || err}${diagDetail ? ` [${diagDetail}]` : ""}`, "system");
  } finally {
    if (runButton) runButton.disabled = false;
    if (loadingText) loadingText.style.display = "none";
  }
}

// ---------------------------------------------------------------------------
// Cloud Logging Dispatcher for Client Diagnostics & Troubleshooting
// ---------------------------------------------------------------------------
let pendingDiagLogs = [];
let diagLogFlushTimeout = null;

export function sendDiagnosticLogToCloud(message, level = "INFO", category = "DIAGNOSTICS", details = {}) {
  const entry = {
    timestamp: new Date().toISOString(),
    level: (level || "INFO").toUpperCase(),
    category: category || "DIAGNOSTICS",
    message: String(message),
    details: details || {}
  };

  pendingDiagLogs.push(entry);

  if (!diagLogFlushTimeout) {
    diagLogFlushTimeout = setTimeout(flushDiagnosticLogsToCloud, 600);
  }
}

async function flushDiagnosticLogsToCloud() {
  diagLogFlushTimeout = null;
  if (pendingDiagLogs.length === 0) return;

  const logsToSend = [...pendingDiagLogs];
  pendingDiagLogs = [];

  const profile = getUserProfile();
  const baseUrl = getActiveProxyUrl().replace(/\/askGeminiEnterprise$/, '');
  const logUrl = `${baseUrl}/api/diagnostics/log`;

  const payload = {
    logs: logsToSend,
    client_context: {
      host: hostAdapter?.name || (typeof Office !== "undefined" && Office.context?.host) || "UnknownHost",
      platform: (typeof Office !== "undefined" && Office.context?.platform) || "UnknownPlatform",
      user_id: profile?.user_id || "anonymous",
      user_email: profile?.email || null,
      tenant_id: profile?.tenant_id || null,
      session_id: currentSessionId || null,
      user_agent: typeof navigator !== "undefined" ? navigator.userAgent : null,
      addin_version: "1.0.0"
    }
  };

  try {
    const res = await fetch(logUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      keepalive: true
    });
    if (!res.ok) {
      console.warn("Cloud log ingestion HTTP error:", res.status);
    }
  } catch (err) {
    console.debug("Failed to dispatch diagnostic logs to Cloud Logging:", err);
  }
}

export function logToDiagBox(msg, isError = false, category = "DIAGNOSTICS", details = {}) {
  const box = document.getElementById("diagLogBox");
  if (box) {
    const time = new Date().toLocaleTimeString();
    box.innerText += `\n[${time}] ${msg}`;
    box.scrollTop = box.scrollHeight;
  }
  // Stream directly to GCP Cloud Logging
  const level = isError ? "ERROR" : "INFO";
  sendDiagnosticLogToCloud(msg, level, category, details);
}
window.logToDiagBox = logToDiagBox;
window.sendDiagnosticLogToCloud = sendDiagnosticLogToCloud;

async function updateDiagnosticsPanel() {
  try {
    // 1. Entra ID SSO diagnostics
    const ssoStatus = document.getElementById("diagSsoStatus");
    const ssoUser = document.getElementById("diagSsoUser");
    const ssoTenant = document.getElementById("diagSsoTenant");
    const profile = getUserProfile();
    const token = await getOfficeAuthToken().catch(() => null);
    const lastErr = getLastAuthError();

    if (token && profile.is_authenticated) {
      if (ssoStatus) {
        ssoStatus.innerText = "✅ Authenticated";
        ssoStatus.style.color = "#107c41";
      }
      if (ssoUser) ssoUser.innerText = profile.email || profile.name || "Signed In";
      if (ssoTenant) ssoTenant.innerText = profile.tenant_id ? (profile.tenant_id.length > 20 ? profile.tenant_id.substring(0, 18) + '...' : profile.tenant_id) : "Default Tenant";
    } else {
      if (ssoStatus) {
        ssoStatus.innerText = lastErr ? `⚠️ ${lastErr.code || lastErr.message || 'Not Authenticated'}` : "Not Authenticated";
        ssoStatus.style.color = "#a4262c";
      }
      if (ssoUser) ssoUser.innerText = profile.email && profile.email !== 'user@organization.com' ? profile.email : "--";
      if (ssoTenant) ssoTenant.innerText = "--";
    }

    // 2. Google OAuth diagnostics
    const googleStatus = document.getElementById("diagGoogleStatus");
    const googleClientId = document.getElementById("diagGoogleClientId");
    const googleExpiry = document.getElementById("diagGoogleExpiry");
    const gToken = getGoogleAccessToken();
    const gClientId = await getGoogleOAuthClientId();

    if (googleClientId) {
      googleClientId.innerText = gClientId ? (gClientId.length > 24 ? gClientId.substring(0, 22) + '...' : gClientId) : 'Not Configured';
      googleClientId.title = gClientId || 'Google OAuth Client ID not configured';
    }

    if (gToken) {
      if (googleStatus) {
        googleStatus.innerText = "✅ Connected";
        googleStatus.style.color = "#107c41";
      }
      const exp = parseInt(window.sessionStorage?.getItem('google_user_token_expiry') || '0', 10);
      const remainingMin = exp ? Math.max(0, Math.round((exp - Date.now()) / 60000)) : 0;
      if (googleExpiry) googleExpiry.innerText = `${remainingMin} min remaining`;
    } else {
      if (googleStatus) {
        googleStatus.innerText = "⚪ Not Connected";
        googleStatus.style.color = "#605e5c";
      }
      if (googleExpiry) googleExpiry.innerText = "--";
    }

    // 3. Backend & Config
    const proxyUrl = document.getElementById("diagProxyUrl");
    const authMode = document.getElementById("diagAuthMode");
    const projId = document.getElementById("diagProjectId");

    const appConfig = await fetchAppConfig();
    if (proxyUrl) {
      const activeUrl = getActiveProxyUrl();
      proxyUrl.innerText = activeUrl.replace('https://', '').split('/')[0];
      proxyUrl.title = activeUrl;
    }
    if (authMode) authMode.innerText = appConfig?.user_auth_mode || "Office SSO + Google OAuth";
    if (projId) projId.innerText = appConfig?.project_id || "agentspace-452714";

    // PowerPoint & Office.js Engine Status
    const pptApiStatus = document.getElementById("diagPptApiStatus");
    if (pptApiStatus) {
      if (typeof Office !== "undefined" && Office.context?.requirements) {
        const v15 = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
        const v14 = Office.context.requirements.isSetSupported("PowerPointApi", "1.4");
        const v13 = Office.context.requirements.isSetSupported("PowerPointApi", "1.3");
        const v12 = Office.context.requirements.isSetSupported("PowerPointApi", "1.2");
        const v11 = Office.context.requirements.isSetSupported("PowerPointApi", "1.1");
        const highest = v15 ? "1.5" : v14 ? "1.4" : v13 ? "1.3" : v12 ? "1.2" : v11 ? "1.1" : "Base";
        pptApiStatus.innerText = `PowerPointApi ${highest} Supported`;
        pptApiStatus.style.color = "#107c41";
      } else {
        const isPPT = typeof PowerPoint !== "undefined";
        pptApiStatus.innerText = isPPT ? "PowerPoint.js Ready" : (Office?.context?.host || "Office Environment");
      }
    }

  } catch (e) {
    console.warn("Diagnostics update error:", e);
  }
}

function initTroubleshootPanel() {
  const panel = document.getElementById("troubleshootPanel");
  if (panel) {
    panel.ontoggle = () => {
      if (panel.open) {
        updateDiagnosticsPanel();
        logToDiagBox("Troubleshooting panel expanded.");
      }
    };
  }

  const btnRefreshSso = document.getElementById("diagRefreshSso");
  if (btnRefreshSso) {
    btnRefreshSso.onclick = async () => {
      logToDiagBox("Forcing Entra ID SSO token refresh...", false, "SSO");
      btnRefreshSso.innerText = "⏳ Refreshing...";
      try {
        const token = await getOfficeAuthToken(true);
        if (token) {
          const profile = getUserProfile();
          logToDiagBox("✅ SSO Token refreshed successfully.", false, "SSO", { email: profile?.email, tenant: profile?.tenant_id });
        } else {
          logToDiagBox("⚠️ Token refresh returned empty.", true, "SSO");
        }
      } catch (err) {
        logToDiagBox(`❌ SSO Refresh error: ${err.message || err}`, true, "SSO", { error: String(err) });
      } finally {
        btnRefreshSso.innerText = "🔄 Refresh SSO Token";
        await initAuthUI();
        await updateDiagnosticsPanel();
      }
    };
  }

  const btnSignInGoogle = document.getElementById("diagSignInGoogle");
  if (btnSignInGoogle) {
    btnSignInGoogle.onclick = async () => {
      logToDiagBox("Launching Google OAuth sign-in flow...", false, "GOOGLE_OAUTH");
      btnSignInGoogle.innerText = "⏳ Signing in...";
      try {
        const profile = getUserProfile();
        const res = await initiateGoogleSignIn(profile?.email || null, 'select_account');
        if (res.status === 'success') {
          logToDiagBox("✅ Google Sign-In succeeded.", false, "GOOGLE_OAUTH", { status: 'success' });
        } else {
          logToDiagBox(`⚠️ Google Sign-In canceled/failed: ${res.error || 'Unknown'}`, true, "GOOGLE_OAUTH", { error: res.error });
        }
      } catch (err) {
        logToDiagBox(`❌ Google Sign-In error: ${err.message || err}`, true, "GOOGLE_OAUTH", { error: String(err) });
      } finally {
        btnSignInGoogle.innerText = "🔑 Sign In with Google";
        await initAuthUI();
        await updateDiagnosticsPanel();
      }
    };
  }

  const btnClearGoogle = document.getElementById("diagClearGoogle");
  if (btnClearGoogle) {
    btnClearGoogle.onclick = async () => {
      setGoogleAccessToken(null);
      logToDiagBox("🗑️ Google access token cleared.", false, "GOOGLE_OAUTH");
      await initAuthUI();
      await updateDiagnosticsPanel();
    };
  }

  const btnPingBackend = document.getElementById("diagPingBackend");
  if (btnPingBackend) {
    btnPingBackend.onclick = async () => {
      const startTime = Date.now();
      btnPingBackend.innerText = "⏳ Pinging...";
      const baseUrl = getActiveProxyUrl().replace(/\/askGeminiEnterprise$/, '');
      const configUrl = `${baseUrl}/api/config`;
      logToDiagBox(`Testing backend connectivity -> ${configUrl}`, false, "BACKEND_PING");
      try {
        const res = await fetch(configUrl, { cache: 'no-store' });
        const latency = Date.now() - startTime;
        if (res.ok) {
          const cfg = await res.json();
          logToDiagBox(`✅ Backend reachable (${latency}ms). Project: ${cfg.project_id || 'OK'}, AuthMode: ${cfg.user_auth_mode || 'OK'}`, false, "BACKEND_PING", { latency, config: cfg });
        } else {
          logToDiagBox(`⚠️ Backend HTTP ${res.status} (${latency}ms)`, true, "BACKEND_PING", { latency, status: res.status });
        }
      } catch (err) {
        logToDiagBox(`❌ Connectivity failed: ${err.message || err}`, true, "BACKEND_PING", { error: String(err) });
      } finally {
        btnPingBackend.innerText = "📡 Test Connection";
        await updateDiagnosticsPanel();
      }
    };
  }

  // PowerPoint Diagnostic & Live Test Actions
  const btnDiagCheckApi = document.getElementById("btnDiagCheckApi");
  if (btnDiagCheckApi) {
    btnDiagCheckApi.onclick = async () => {
      logToDiagBox("Checking PowerPoint Office.js environment...", false, "POWERPOINT_API");
      try {
        const hasOffice = typeof Office !== "undefined";
        const hasPPT = typeof PowerPoint !== "undefined";
        const host = (hasOffice && Office.context?.host) || "Unknown";
        const platform = (hasOffice && Office.context?.platform) || "Unknown";

        let versions = {};
        if (hasOffice && Office.context?.requirements) {
          versions = {
            v11: Office.context.requirements.isSetSupported("PowerPointApi", "1.1"),
            v12: Office.context.requirements.isSetSupported("PowerPointApi", "1.2"),
            v13: Office.context.requirements.isSetSupported("PowerPointApi", "1.3"),
            v14: Office.context.requirements.isSetSupported("PowerPointApi", "1.4"),
            v15: Office.context.requirements.isSetSupported("PowerPointApi", "1.5")
          };
          logToDiagBox(`Host: ${host}, Platform: ${platform}, PowerPointApi: 1.1=${versions.v11}, 1.2=${versions.v12}, 1.3=${versions.v13}, 1.4=${versions.v14}, 1.5=${versions.v15}`, false, "POWERPOINT_API", { host, platform, versions });
        } else {
          logToDiagBox(`Host: ${host}, Platform: ${platform}, PPT Namespace: ${hasPPT}`, false, "POWERPOINT_API", { host, platform, hasPPT });
        }
      } catch (e) {
        logToDiagBox(`❌ API Check Error: ${e.message}`, true, "POWERPOINT_API", { error: String(e) });
      }
    };
  }

  const btnDiagTestSlide = document.getElementById("btnDiagTestSlide");
  if (btnDiagTestSlide) {
    btnDiagTestSlide.onclick = async () => {
      logToDiagBox("Calling PowerPoint.run(presentation.slides.add())...", false, "POWERPOINT_TEST");
      const startTime = Date.now();
      try {
        if (typeof PowerPoint === "undefined") {
          throw new Error("PowerPoint namespace is not available in current host.");
        }
        await PowerPoint.run(async (context) => {
          context.presentation.slides.add();
          logToDiagBox("Slide add queued. Calling context.sync()...", false, "POWERPOINT_TEST");
          await context.sync();
        });
        const elapsed = Date.now() - startTime;
        logToDiagBox(`✅ Successfully added 1 blank slide via PowerPoint.run! (${elapsed}ms)`, false, "POWERPOINT_TEST", { test: "add_slide", elapsed });
      } catch (e) {
        logToDiagBox(`❌ Slide Add Error: ${e.message} (code: ${e.code || "N/A"})`, true, "POWERPOINT_TEST", { error: String(e), code: e.code });
      }
    };
  }

  const btnDiagTestText = document.getElementById("btnDiagTestText");
  if (btnDiagTestText) {
    btnDiagTestText.onclick = async () => {
      logToDiagBox("Testing Add Slide + Title Textbox + Body Textbox via getCount()...", false, "POWERPOINT_TEST");
      const startTime = Date.now();
      try {
        if (typeof PowerPoint === "undefined") {
          throw new Error("PowerPoint namespace is not available in current host.");
        }
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.add();
          await context.sync();

          const countResult = slides.getCount();
          await context.sync();

          const slideCount = countResult.value;
          logToDiagBox(`Slide count: ${slideCount}. Fetching slide at index ${slideCount - 1}...`, false, "POWERPOINT_TEST");
          const slide = slides.getItemAt(slideCount - 1);

          slide.shapes.addTextBox("🧪 Diagnostic Test Title", {
            left: 50,
            top: 40,
            width: 650,
            height: 55
          });

          slide.shapes.addTextBox("• Bullet item 1: Official Microsoft Office.js pattern\n• Bullet item 2: Direct geometry textbox creation verified", {
            left: 50,
            top: 110,
            width: 650,
            height: 300
          });

          await context.sync();
        });
        const elapsed = Date.now() - startTime;
        logToDiagBox(`✅ Successfully created Slide with Title & Body Textbox! (${elapsed}ms)`, false, "POWERPOINT_TEST", { test: "add_slide_textbox", elapsed });
      } catch (e) {
        logToDiagBox(`❌ Slide Text Error: ${e.message} (code: ${e.code || "N/A"})`, true, "POWERPOINT_TEST", { error: String(e), code: e.code });
      }
    };
  }

  const btnClearLog = document.getElementById("diagClearLog");
  if (btnClearLog) {
    btnClearLog.onclick = (e) => {
      e.stopPropagation();
      const box = document.getElementById("diagLogBox");
      if (box) box.innerText = "[Ready] Log cleared.";
    };
  }
}

