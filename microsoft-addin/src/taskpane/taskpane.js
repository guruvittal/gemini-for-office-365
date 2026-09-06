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
let userClearedSelection = false;

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

  // Setup Transform Doc to Deck Feature
  initDocToDeckFeature();

  // Setup Insert Executive Visual Feature
  initExecutiveVisualFeature();

  // Wire Tip Banner Dismiss button
  const dismissTipBtn = document.getElementById("dismissTipBtn");
  if (dismissTipBtn) {
    dismissTipBtn.onclick = () => {
      const tipBanner = document.getElementById("tipBanner");
      if (tipBanner) tipBanner.style.display = "none";
    };
  }

  // Wire Context Attachment Pill Clear button
  const clearSelectionBtn = document.getElementById("clearSelectionBtn");
  if (clearSelectionBtn) {
    clearSelectionBtn.onclick = () => {
      userClearedSelection = true;
      currentSelectedText = "";
      renderAdaptiveActionChips(false);
    };
  }

  // Attach selection change handler for in-document detection & adaptive toolbar
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
      const errHint = lastErr ? (lastErr.message || lastErr.code || JSON.stringify(lastErr)) : "Not connected to Microsoft 365";
      if (userEmailText) {
        userEmailText.innerText = profile.email && profile.email !== 'user@organization.com' ? profile.email : "Not Connected";
        userEmailText.className = "user-email-text offline";
        userEmailText.title = isWifMode ? `WIF SSO Inactive (${errHint})` : `SSO: ${errHint}`;
      }
      if (userStatusDot) {
        userStatusDot.className = "user-status-dot offline";
        userStatusDot.title = isWifMode ? `WIF SSO Inactive (${errHint})` : `SSO: ${errHint}`;
      }
      if (userAuthBar) {
        userAuthBar.title = isWifMode ? `WIF SSO Inactive (${errHint})` : `SSO: ${errHint}`;
      }
    }

    // 2. Google OAuth button (Used only in GSuite / Cloud Identity mode, hidden in WIF mode)
    if (googleDriveBtn) {
      if (isWifMode) {
        // In WIF mode, user identity is purely Microsoft Entra ID exchanged silently with Google STS.
        // No secondary badge or button is shown.
        googleDriveBtn.style.display = "none";
      } else {
        googleDriveBtn.style.display = "inline-flex";
        if (isGoogleTokenValid()) {
          googleDriveBtn.className = "google-drive-btn connected";
          googleDriveBtn.innerHTML = "✅ Gemini Connected";
          googleDriveBtn.title = "Gemini Enterprise grounding active (OAuth token valid)";
          googleDriveBtn.style.cursor = "pointer";
        } else {
          googleDriveBtn.className = "google-drive-btn";
          googleDriveBtn.innerHTML = "Login with Google";
          googleDriveBtn.title = "Click to sign into Google for Gemini Enterprise grounding";
          googleDriveBtn.style.cursor = "pointer";
        }
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

  if (hostName === "PowerPoint") {
    if (tipBannerText) tipBannerText.innerHTML = `💡 Type <code>@gemini &lt;prompt&gt;</code> in Slide`;
    if (docToolsTitle) docToolsTitle.innerHTML = `📊 <strong>Chat with Slides:</strong>`;
    if (debugStatus) debugStatus.innerText = `PowerPoint Ready`;
    if (welcomeBubble) welcomeBubble.innerHTML = `Type a prompt below, click an action chip above, or select slide content!`;
  } else if (hostName === "Excel") {
    if (tipBannerText) tipBannerText.innerHTML = `💡 Type <code>@gemini &lt;prompt&gt;</code> in Cell`;
    if (docToolsTitle) docToolsTitle.innerHTML = `📈 <strong>Chat with Spreadsheet:</strong>`;
    if (debugStatus) debugStatus.innerText = `Excel Ready`;
    if (welcomeBubble) welcomeBubble.innerHTML = `Type a prompt below, click an action chip above, or select cells in Excel!`;
  } else {
    // Word (Default)
    if (tipBannerText) tipBannerText.innerHTML = `💡 Type <code>@gemini &lt;prompt&gt;</code> in Doc`;
    if (docToolsTitle) docToolsTitle.innerHTML = `📄 <strong>Chat with Document:</strong>`;
    if (debugStatus) debugStatus.innerText = `Word Ready`;
    if (welcomeBubble) welcomeBubble.innerHTML = `Type a prompt below, click an action chip above, or highlight text in Word!`;
  }

  renderAdaptiveActionChips(false);
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

// Adaptive Action Bar & Chips Renderer
function renderAdaptiveActionChips(isSelected = false, slideCount = 1, words = 0, isTextSelection = false) {
  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const docToolsTitle = document.getElementById("docToolsTitle");
  const debugStatus = document.getElementById("debugStatus");
  const container = document.getElementById("docToolsChipsContainer");
  const selectionPill = document.getElementById("selectionAttachmentPill");
  const selectionPillLabel = document.getElementById("selectionPillLabel");

  if (!container) return;

  if (hostName === "PowerPoint") {
    if (isSelected) {
      const labelText = isTextSelection 
        ? `Selected Text (${words} words)` 
        : (slideCount > 1 ? `${slideCount} Slides (${words} words)` : `Selected Content (${words} words)`);
      const barTitle = isTextSelection 
        ? `✨ <strong>Selected Text:</strong>` 
        : `✨ <strong>Selected Slides (${slideCount}):</strong>`;

      if (docToolsTitle) docToolsTitle.innerHTML = barTitle;
      if (debugStatus) debugStatus.innerHTML = `<span style="color:#0078d4; font-weight:600;">${labelText}</span>`;

      if (selectionPill) {
        if (selectionPillLabel) selectionPillLabel.innerText = `Attached: ${labelText}`;
        selectionPill.style.display = "flex";
      }

      container.innerHTML = `
        <button class="quick-chip" id="chipSummarize" style="background-color:#e0f2fe; color:#0369a1; border-color:#bae6fd;">📊 Summarize Slides</button>
        <button class="quick-chip" id="chipExecVisual" style="background-color:#ecfdf5; color:#047857; border-color:#a7f3d0; font-weight:600;">✨ Insert Executive Visual</button>
        <button class="quick-chip" id="chipRisks" style="background-color:#fef3c7; color:#b45309; border-color:#fde68a;">⚠️ Key Risks</button>
        <button class="quick-chip" id="chipRewrite" style="background-color:#f3e8ff; color:#7e22ce; border-color:#e9d5ff;">🪄 Rewrite Slide</button>
        <button class="quick-chip" id="chipDocToDeck" style="background-color:#eef2ff; color:#4338ca; border-color:#c7d2fe; font-weight:600;">📄 Doc to Deck</button>
        <button class="quick-chip" id="chipShorten">📉 Shorten</button>
        <button class="quick-chip" id="chipTable">📊 Table</button>
      `;

      const cSummarize = document.getElementById("chipSummarize");
      if (cSummarize) cSummarize.onclick = () => runPowerPointSlideAction("takeaways");

      const cExecVisual = document.getElementById("chipExecVisual");
      if (cExecVisual) cExecVisual.onclick = () => openExecutiveVisualModal();

      const cRisks = document.getElementById("chipRisks");
      if (cRisks) cRisks.onclick = () => runPowerPointSlideAction("risks");

      const cRewrite = document.getElementById("chipRewrite");
      if (cRewrite) cRewrite.onclick = () => runSelectionPrompt("Rewrite and elevate the selected slide content into clear, high-impact executive presentation prose.");

      const cDocToDeck = document.getElementById("chipDocToDeck");
      if (cDocToDeck) {
        cDocToDeck.onclick = () => {
          const fileInput = document.getElementById("docToDeckFileInput");
          if (fileInput) fileInput.click();
        };
      }

      const cShorten = document.getElementById("chipShorten");
      if (cShorten) cShorten.onclick = () => runSelectionPrompt("Make this slide content significantly more concise and punchy, removing unnecessary fluff.");

      const cTable = document.getElementById("chipTable");
      if (cTable) cTable.onclick = () => runSelectionPrompt("Convert this slide's metrics and structured data into a clean, executive markdown table.");

    } else {
      if (docToolsTitle) docToolsTitle.innerHTML = `📊 <strong>Chat with Slides:</strong>`;
      if (debugStatus) debugStatus.innerText = `PowerPoint Ready`;

      if (selectionPill) selectionPill.style.display = "none";

      container.innerHTML = `
        <button class="quick-chip" id="chipDocToDeck" style="background-color:#eef2ff; color:#4338ca; border-color:#c7d2fe; font-weight:600;">📄 Doc to Deck</button>
        <button class="quick-chip" id="chipExecVisual" style="background-color:#ecfdf5; color:#047857; border-color:#a7f3d0; font-weight:600;">✨ Insert Executive Visual</button>
        <button class="quick-chip" id="chipSummarize" style="background-color:#e0f2fe; color:#0369a1; border-color:#bae6fd;">📊 Summarize Slides</button>
      `;

      const cDocToDeck = document.getElementById("chipDocToDeck");
      if (cDocToDeck) {
        cDocToDeck.onclick = () => {
          const fileInput = document.getElementById("docToDeckFileInput");
          if (fileInput) fileInput.click();
        };
      }

      const cExecVisual = document.getElementById("chipExecVisual");
      if (cExecVisual) cExecVisual.onclick = () => openExecutiveVisualModal();

      const cSum = document.getElementById("chipSummarize");
      if (cSum) cSum.onclick = () => runPowerPointSlideAction("takeaways");
    }
  } else if (hostName === "Excel") {
    if (isSelected) {
      if (docToolsTitle) docToolsTitle.innerHTML = `✨ <strong>Selected Cells:</strong>`;
      if (debugStatus) debugStatus.innerHTML = `<span style="color:#0078d4; font-weight:600;">${words} words</span>`;
      if (selectionPill) {
        if (selectionPillLabel) selectionPillLabel.innerText = `Selected Cells (${words} words)`;
        selectionPill.style.display = "flex";
      }
      container.innerHTML = `
        <button class="quick-chip" id="selRewrite">🪄 Rewrite</button>
        <button class="quick-chip" id="selShorten">📉 Shorten</button>
        <button class="quick-chip" id="selTable">📊 Table</button>
      `;
      setupSelectionChips();
    } else {
      if (docToolsTitle) docToolsTitle.innerHTML = `📈 <strong>Chat with Spreadsheet:</strong>`;
      if (debugStatus) debugStatus.innerText = `Excel Ready`;
      if (selectionPill) selectionPill.style.display = "none";

      container.innerHTML = `
        <button class="quick-chip" id="chipSummarize">📈 Summarize Sheet</button>
        <button class="quick-chip" id="chipRisks">⚠️ Data & Formula Risks</button>
        <button class="quick-chip" id="chipActionItems">✅ Action Items</button>
        <button class="quick-chip" id="chipExecBox">💡 Key Metrics Card</button>
      `;
      setupDocToolsChips();
    }
  } else {
    // Word
    if (isSelected) {
      if (docToolsTitle) docToolsTitle.innerHTML = `✨ <strong>Selected Text:</strong>`;
      if (debugStatus) debugStatus.innerHTML = `<span style="color:#0078d4; font-weight:600;">${words} words</span>`;
      if (selectionPill) {
        if (selectionPillLabel) selectionPillLabel.innerText = `Selected Text (${words} words)`;
        selectionPill.style.display = "flex";
      }
      container.innerHTML = `
        <button class="quick-chip" id="selRewrite">🪄 Rewrite</button>
        <button class="quick-chip" id="selProfessional">💼 Professional</button>
        <button class="quick-chip" id="selShorten">📉 Shorten</button>
        <button class="quick-chip" id="selBullets">🎯 Bullet Points</button>
        <button class="quick-chip" id="selTable">📊 Table</button>
      `;
      setupSelectionChips();
    } else {
      if (docToolsTitle) docToolsTitle.innerHTML = `📄 <strong>Chat with Document:</strong>`;
      if (debugStatus) debugStatus.innerText = `Word Ready`;
      if (selectionPill) selectionPill.style.display = "none";

      container.innerHTML = `
        <button class="quick-chip" id="chipSummarize">📄 Summarize Doc</button>
        <button class="quick-chip" id="chipRisks">⚠️ Key Risks</button>
        <button class="quick-chip" id="chipActionItems">✅ Action Items</button>
        <button class="quick-chip" id="chipExecBox">💡 Executive Card</button>
      `;
      setupDocToolsChips();
    }
  }
}

// Fallback / Initial Chip Bindings for Word & Excel
function setupDocToolsChips() {
  const hostName = hostAdapter ? hostAdapter.name : "Word";

  const chipSummarize = document.getElementById("chipSummarize");
  if (chipSummarize) {
    chipSummarize.onclick = () => {
      if (hostName === "PowerPoint") {
        runPowerPointSlideAction("summarize");
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
        runPowerPointSlideAction("risks");
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
      if (hostName === "PowerPoint") {
        runPowerPointSlideAction("action_items");
      } else {
        const noun = hostName === "Excel" ? "spreadsheet" : "document";
        runDocIntelligencePrompt(`Extract all Action Items, Deliverables, and Next Steps from this ${noun}. Present them in a structured table with Task, Owner, and Priority.`);
      }
    };
  }

  const chipExecBox = document.getElementById("chipExecBox");
  if (chipExecBox) {
    chipExecBox.onclick = () => {
      if (hostName === "PowerPoint") {
        runPowerPointSlideAction("takeaways");
      } else if (hostName === "Excel") {
        runDocIntelligencePrompt("Generate an Executive Metrics Summary Card highlighting Key Financial / Operational KPIs and totals from this sheet.");
      } else {
        runDocIntelligencePrompt("Generate an Executive Summary Box for this document using a callout block (> [!NOTE] ...) highlighting Strategic Purpose, Key Metrics, and Impact.");
      }
    };
  }

  const chipDocToDeck = document.getElementById("chipDocToDeck");
  if (chipDocToDeck) {
    chipDocToDeck.onclick = () => {
      const fileInput = document.getElementById("docToDeckFileInput");
      if (fileInput) fileInput.click();
    };
  }
}

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

// Handle Host Selection Changes Dynamically (Adapts Top Bar and Shows Attachment Pill)
async function handleSelectionChanged() {
  if (!hostAdapter) return;

  try {
    const text = await hostAdapter.getSelectedText();
    const newText = text ? text.trim() : "";

    // If selection content changed, reset explicit dismissal flag
    if (newText !== currentSelectedText) {
      userClearedSelection = false;
    }

    if (userClearedSelection) {
      return;
    }

    currentSelectedText = newText;

    if (currentSelectedText.length > 5) {
      const words = currentSelectedText.split(/\s+/).filter(w => w.length > 0).length;
      let slideCount = 1;
      let isTextSelection = false;
      if (hostAdapter.name === "PowerPoint" && typeof hostAdapter.getSelectedSlidesText === "function") {
        try {
          const slides = await hostAdapter.getSelectedSlidesText();
          if (slides && slides.length > 0) {
            slideCount = slides.length;
            if (slides.length === 1 && slides[0].isTextSelection) {
              isTextSelection = true;
            }
          }
        } catch (e) {}
      }
      renderAdaptiveActionChips(true, slideCount, words, isTextSelection);
    } else {
      renderAdaptiveActionChips(false);
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
    appendBubble(hostAdapter?.name === "PowerPoint" ? "Please select text or a shape on the slide first." : "Please highlight text in Word first.", "system");
    return;
  }

  let fullPrompt = `Selected Content:\n"""\n${currentSelectedText}\n"""\n\nTask: ${instruction}`;
  if (hostAdapter.name === "PowerPoint") {
    const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
    fullPrompt = enhancePromptForPowerPoint(fullPrompt);
  }

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

// Dedicated PowerPoint Slide Intelligence Action (Targets Highlighted / Selected Slides)
async function runPowerPointSlideAction(actionType) {
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Reading highlighted slide(s)...";
    loadingText.style.display = "block";
  }

  let slides = [];
  try {
    if (hostAdapter && typeof hostAdapter.getSelectedSlidesText === "function") {
      slides = await hostAdapter.getSelectedSlidesText();
    }
  } catch (e) {
    console.warn("Could not retrieve selected slides:", e);
  }

  let slideContext = "";
  let slideLabel = "current slide";

  if (slides && slides.length > 0) {
    slideLabel = slides.length === 1 ? "highlighted slide" : `${slides.length} highlighted slides`;
    slideContext = slides.map((s, idx) => `[Highlighted Slide ${idx + 1}]:\n${s.text}`).join("\n\n---\n\n");
  } else {
    // Fall back to full presentation text if no specific slide is highlighted
    const fullDocText = await hostAdapter.getFullDocumentText();
    if (fullDocText && fullDocText.length > 10) {
      slideContext = fullDocText;
      slideLabel = "presentation deck";
    }
  }

  let taskInstruction = "";
  let displayBubble = "";

  switch (actionType) {
    case "risks":
      taskInstruction = `Based on the context from the ${slideLabel} provided below, extract and analyze all Key Strategic & Operational Risks, Blockers, Gaps, and Dependencies. Create content for a new PowerPoint slide titled "⚠️ Key Risks & Mitigations" with high-impact bullet points, severity assessments, and concrete mitigation actions.`;
      displayBubble = `⚠️ [Key Risks] Generating new risk analysis slide from ${slideLabel}...`;
      break;
    case "summarize":
    case "takeaways":
      taskInstruction = `Based on the context from the ${slideLabel} provided below, create EXACTLY ONE executive summary slide.
CRITICAL CONSTRAINT: You must output ONLY ONE SINGLE SLIDE. Do NOT generate multiple slides.
Create content for a new PowerPoint slide titled "📊 Executive Slide Summary" highlighting high-impact findings, core metrics, and strategic implications.`;
      displayBubble = `📊 [Summarize Slides] Generating executive summary slide from ${slideLabel}...`;
      break;
    case "action_items":
      taskInstruction = `Based on the context from the ${slideLabel} provided below, extract all Action Items, Deliverables, Next Steps, and Ownership. Create content for a new PowerPoint slide titled "✅ Action Items & Next Steps" with actionable task bullets, suggested owners, and priority levels.`;
      displayBubble = `✅ [Action Items] Generating action items slide from ${slideLabel}...`;
      break;
  }

  let fullPrompt = "";
  if (slideContext && slideContext.length > 5) {
    fullPrompt = `Context from ${slideLabel}:\n"""\n${slideContext.substring(0, 50000)}\n"""\n\nTask: ${taskInstruction}`;
  } else {
    fullPrompt = taskInstruction;
  }

  // Apply PowerPoint slide generation formatting rules
  const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
  fullPrompt = enhancePromptForPowerPoint(fullPrompt);

  await executeGeminiWorkflow(fullPrompt, displayBubble);
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

// Feature: Transform Doc to Deck (Direct document attachments to Discovery Engine streamAssist)
function initDocToDeckFeature() {
  const fileInput = document.getElementById("docToDeckFileInput");
  if (!fileInput) return;

  fileInput.addEventListener("change", async (event) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    const loadingText = document.getElementById("loading");
    if (loadingText) {
      loadingText.innerText = `⚡ Uploading ${files.length} document(s) to Gemini Enterprise...`;
      loadingText.style.display = "block";
    }

    try {
      const attachments = [];
      const fileNames = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        fileNames.push(file.name);

        const base64Data = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => {
            const result = reader.result;
            const base64 = typeof result === 'string' && result.includes(',') ? result.split(',')[1] : result;
            resolve(base64);
          };
          reader.onerror = (err) => reject(err);
          reader.readAsDataURL(file);
        });

        let mimeType = file.type;
        if (!mimeType) {
          const ext = file.name.split('.').pop().toLowerCase();
          if (ext === 'pdf') mimeType = 'application/pdf';
          else if (ext === 'docx') mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
          else if (ext === 'doc') mimeType = 'application/msword';
          else if (ext === 'txt') mimeType = 'text/plain';
          else if (ext === 'md') mimeType = 'text/markdown';
          else mimeType = 'application/octet-stream';
        }

        attachments.push({
          fileName: file.name,
          mimeType: mimeType,
          fileContents: base64Data
        });
      }

      const displayUserBubble = `📄 [Doc to Deck]\nAttached: ${fileNames.join(', ')}`;
      const prompt = `Extract the key takeaways from the attached document(s) and create an executive presentation slide deck with no more than 5 slides summarizing the key insights.

PowerPoint Slide Deck Requirements:
1. Provide EXACTLY ONE presentation deck with no more than 5 slides. DO NOT output multiple alternative options.
2. DO NOT output conversational preamble or filler (e.g. "Here is...", "Sure!"). Output the slide deck content directly.
3. For each slide, structure with rich executive visual hierarchy:
   - "## Slide <N>: <Emoji> <Punchy Slide Title (3-4 words max)>"
   - "Subtitle: <Crisp Subtitle / Strategic Context>"
   - Diversify slide visual formats across the deck to create an engaging executive presentation flow:
     * Qualitative / Strategic Slides (Overview, Vision, Operational Pillars, Roadmap): Use 3 to 4 high-impact bullet points with bold lead-ins (• **Strategic Pillar:** Detailed description). Do NOT include tables on these slides.
     * Quantitative / Comparative Slides (Financials, KPIs, Comparisons): When presenting dense metrics or side-by-side comparisons, use a clean Markdown table (| Metric / Dimension | Detail | Implication |). Limit tables to at most 1 or 2 slides in the entire deck.
   - Conclude with a clear strategic takeaway line: "Takeaway: <Executive takeaway sentence>"
4. Do NOT output internal design metadata like "Visual Concept:", "Color:", or font sizes.`;

      await executeGeminiWorkflow(prompt, displayUserBubble, attachments);
    } catch (err) {
      console.error("Error reading documents for Transform Doc to Deck:", err);
      appendBubble("Error reading documents: " + err.message, "system");
    } finally {
      event.target.value = "";
      if (loadingText) loadingText.style.display = "none";
    }
  });
}

// Feature: Insert Executive Visual (3-Column Metric Grid or Before/After Comparison)
function initExecutiveVisualFeature() {
  const modal = document.getElementById("execVisualModal");
  const closeBtn = document.getElementById("closeExecVisualModal");
  const cancelBtn = document.getElementById("cancelExecVisualBtn");
  const generateBtn = document.getElementById("generateExecVisualBtn");
  const promptInput = document.getElementById("execVisualPrompt");
  const optMetric = document.getElementById("optMetricGrid");
  const optBeforeAfter = document.getElementById("optBeforeAfter");

  if (!modal) return;

  let selectedType = "metric_grid_3col";

  if (optMetric) {
    optMetric.onclick = () => {
      selectedType = "metric_grid_3col";
      optMetric.classList.add("selected");
      if (optBeforeAfter) optBeforeAfter.classList.remove("selected");
    };
  }

  if (optBeforeAfter) {
    optBeforeAfter.onclick = () => {
      selectedType = "before_after";
      optBeforeAfter.classList.add("selected");
      if (optMetric) optMetric.classList.remove("selected");
    };
  }

  const closeModal = () => {
    modal.style.display = "none";
  };

  if (closeBtn) closeBtn.onclick = closeModal;
  if (cancelBtn) cancelBtn.onclick = closeModal;

  if (generateBtn) {
    generateBtn.onclick = async () => {
      closeModal();
      const customTopic = promptInput ? promptInput.value.trim() : "";
      if (promptInput) promptInput.value = "";
      await runGenerateExecutiveVisual(selectedType, customTopic);
    };
  }
}

function openExecutiveVisualModal() {
  const modal = document.getElementById("execVisualModal");
  if (modal) {
    modal.style.display = "flex";
    const promptInput = document.getElementById("execVisualPrompt");
    if (promptInput) promptInput.focus();
  }
}

async function runGenerateExecutiveVisual(visualType, customTopic = "") {
  let slideContext = "";
  try {
    if (typeof hostAdapter.getSelectedSlidesText === "function") {
      slideContext = await hostAdapter.getSelectedSlidesText();
    }
  } catch (_) {}

  const topicContext = customTopic || (slideContext && slideContext.length > 20 
    ? `the following slide context:\n"""\n${slideContext.substring(0, 30000)}\n"""` 
    : "enterprise transformation, operational performance, and strategic execution");

  let prompt = "";
  let displayBubble = "";

  if (visualType === "metric_grid_3col") {
    displayBubble = `✨ [Executive Visual] Generating 3-Column Metric Grid...`;
    prompt = `You are an elite executive presentation designer. Create a high-impact PowerPoint 3-Column Metric Grid visual slide based on: ${topicContext}.

You MUST return a JSON object inside a \`\`\`json markdown code block adhering strictly to this schema:
\`\`\`json
{
  "visualType": "metric_grid_3col",
  "title": "Short Impactful Slide Title (e.g. FY25 Key Performance Indicators)",
  "subtitle": "Clear Context Subtitle (e.g. Accelerating enterprise revenue and cloud operational scale)",
  "cards": [
    {
      "metric": "+42%",
      "title": "Revenue Growth",
      "subtitle": "Year-over-Year ARR",
      "bullets": [
        "Driven by cloud AI enterprise adoption",
        "Net retention rate expanded to 124%"
      ]
    },
    {
      "metric": "$1.4B",
      "title": "Global Bookings",
      "subtitle": "Record Annual High",
      "bullets": [
        "Closed 38 Fortune 500 deals",
        "Pipeline grew 3.2x across key verticals"
      ]
    },
    {
      "metric": "99.98%",
      "title": "Operational Uptime",
      "subtitle": "Mission-Critical SLA",
      "bullets": [
        "Sub-second p99 latency globally",
        "Automated failover across multi-region clusters"
      ]
    }
  ]
}
\`\`\`
Ensure metrics are short and punchy (e.g., +35%, $4.2B, 10x, 99.9%). Each card should have exactly 2-3 concise takeaway bullets.`;
  } else {
    displayBubble = `✨ [Executive Visual] Generating Before/After Comparison...`;
    prompt = `You are an elite executive presentation designer. Create a high-impact PowerPoint Before/After Comparison visual slide based on: ${topicContext}.

You MUST return a JSON object inside a \`\`\`json markdown code block adhering strictly to this schema:
\`\`\`json
{
  "visualType": "before_after",
  "title": "Short Impactful Slide Title (e.g. Operational Modernization & AI Transformation)",
  "subtitle": "Clear Context Subtitle (e.g. Transitioning from legacy bottlenecks to intelligent automation)",
  "before": {
    "title": "Legacy / Current State Challenges",
    "bullets": [
      "Fragmented data silos causing 3-day reporting delays",
      "High human operational overhead and error-prone reconciliation",
      "Static quarterly projections lacking predictive insights"
    ]
  },
  "after": {
    "title": "Gemini AI / Future State Transformation",
    "bullets": [
      "Sub-second unified intelligence across enterprise data",
      "Automated presentation drafting & 99.9% data accuracy",
      "Continuous predictive forecasting with actionable real-time alerts"
    ]
  }
}
\`\`\`
Ensure each side has 3-4 clear, parallel, high-contrast bullet points.`;
  }

  // Apply PowerPoint prompt enhancer rules
  const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
  prompt = enhancePromptForPowerPoint(prompt);

  await executeGeminiWorkflow(prompt, displayBubble);
}

async function executeGeminiWorkflow(fullPrompt, displayUserBubble, attachments = null) {
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

    const data = await askGeminiEnterprise(fullPrompt, chatHistoryState, currentSessionId, true, attachments);

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
  const hasSelection = isPPT && currentSelectedText && currentSelectedText.trim().length > 0;
  const replaceBtn = document.createElement("button");
  replaceBtn.className = "action-btn replace";
  replaceBtn.innerHTML = isPPT ? (hasSelection ? `🔄 Replace in Slide` : `🔄 Replace Slide`) : (isExcel ? `🔄 Replace in Sheet` : `🔄 Replace in Doc`);
  replaceBtn.title = isPPT ? (hasSelection ? "Replace selected text in slide" : "Replace active slide content") : "Replace active draft or selection in Word";
  replaceBtn.onclick = async () => {
    await performDocumentInsertion(textDiv.innerHTML, text, "replace_draft");
  };

  // 2. Insert Button
  const insertBtn = document.createElement("button");
  insertBtn.className = "action-btn insert";
  insertBtn.innerHTML = isPPT ? (hasSelection ? `➕ Insert as New Slide` : `➕ Insert Slide`) : (isExcel ? `➕ Insert into Sheet` : `➕ Insert at Cursor`);
  insertBtn.title = isPPT ? "Insert generated slide into presentation" : "Insert at current cursor location";
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
    { label: "📉 Make Shorter", prompt: "Make the above response significantly more concise and punchy for executive reading. Provide exactly ONE finalized version with no conversational preamble or multiple options." },
    { label: "📈 Expand Details", prompt: "Expand the above draft with more in-depth technical, operational, and architectural details." },
    { label: "📊 Format as Table", prompt: "Convert the key findings and aspects of the above response into a structured markdown table." },
    { label: "👔 Executive Tone", prompt: "Rewrite the above response with an authoritative, C-level executive tone." },
    { label: "🔄 Try Again", prompt: "Regenerate the response with a fresh structure and alternative perspective." }
  ];

  chipsData.forEach(item => {
    const chip = document.createElement("button");
    chip.className = "refinement-chip";
    chip.innerText = item.label;
    chip.onclick = async () => {
      let chipPrompt = item.prompt;
      if (hostAdapter?.name === "PowerPoint") {
        const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
        chipPrompt = enhancePromptForPowerPoint(chipPrompt);
      }
      executeGeminiWorkflow(chipPrompt, `${item.label}: "${item.prompt.substring(0, 45)}..."`);
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
    loadingText.innerText = hostAdapter?.name === 'PowerPoint' ? (mode === 'replace_draft' ? "⚡ Replacing slide content..." : "⚡ Creating PowerPoint slides...") : "⚡ Updating document...";
    loadingText.style.display = "block";
  }

  try {
    const isPPT = hostAdapter?.name === "PowerPoint";
    if (isPPT) {
      await hostAdapter.insertContent(htmlContent, rawText, { mode: mode === "replace_draft" ? "replace" : "insert" });
    } else {
      await hostAdapter.insertContent(htmlContent, mode);
    }
    const debugStatus = document.getElementById("debugStatus");
    if (debugStatus) {
      const host = hostAdapter?.name || 'Office';
      debugStatus.innerText = `Updated in ${host} (${mode === 'replace_draft' ? 'Replaced' : 'Inserted'})`;
    }
  } catch (err) {
    console.error("Document insertion error:", err);
    appendBubble(`🔴 Insertion Error: ${err.message || err}`, "system");
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

