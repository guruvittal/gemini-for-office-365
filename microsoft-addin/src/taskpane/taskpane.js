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

/**
 * Clears prior chat interaction history and empties the prompt textarea above the send button.
 * Invoked whenever a user clicks any top action chip/button.
 */
export function clearChatAndInput() {
  const historyDiv = document.getElementById("chatHistory");
  if (historyDiv) {
    historyDiv.innerHTML = "";
  }
  const promptInput = document.getElementById("promptText");
  if (promptInput) {
    promptInput.value = "";
    promptInput.style.height = "auto";
  }
  console.log("🧹 [Chat] Cleared previous chat interactions and reset input box.");
}

Office.onReady(async (info) => {
  // Detect active Microsoft Office host (Word, PowerPoint, Excel) dynamically
  hostAdapter = HostAdapterFactory.getAdapter(info);

  // IMMEDIATELY adapt UI for the active host (PowerPoint, Word, Excel)
  // Ensures "Chat with Slides" loads instantly on startup without waiting for Entra ID SSO login
  adaptUIForHost(hostAdapter.name);

  if (hostAdapter.name === "PowerPoint") {
    initPowerPointDiagnostics();
  }

  // Pre-fetch dynamic backend configuration (Google OAuth Client ID)
  fetchAppConfig().catch(e => console.warn("Background config fetch failed:", e));

  // Initialize Entra ID & Google Drive Identity in UI
  await initAuthUI();

  // Initialize Collapsible Troubleshooting & Diagnostics Panel
  initTroubleshootPanel();

  // Wire capture listener on top buttons to clear chat history and input box on click
  const docToolsContainer = document.getElementById("docToolsChipsContainer");
  if (docToolsContainer) {
    docToolsContainer.addEventListener("click", (e) => {
      if (e.target.closest(".quick-chip")) {
        clearChatAndInput();
      }
    }, true);
  }

  // Wire click-to-zoom on preview images in chat history
  const chatHistoryDiv = document.getElementById("chatHistory");
  if (chatHistoryDiv) {
    chatHistoryDiv.addEventListener("click", (e) => {
      const insertNewBtn = e.target.closest(".img-action-btn-insert-new");
      const insertCurrentBtn = e.target.closest(".img-action-btn-insert-current");
      if (insertNewBtn) {
        e.stopPropagation();
        const src = insertNewBtn.getAttribute("data-img-src");
        if (src) {
          performDocumentInsertion(`<img src="${src}" />`, `![Image](${src})`, "insert_cursor", { imageOnly: true });
        }
        return;
      }
      if (insertCurrentBtn) {
        e.stopPropagation();
        const src = insertCurrentBtn.getAttribute("data-img-src");
        if (src) {
          performDocumentInsertion(`<img src="${src}" />`, `![Image](${src})`, "insert_current_slide", { imageOnly: true });
        }
        return;
      }

      const zoomBtn = e.target.closest(".img-zoom-btn, .img-action-btn-zoom");
      const clickedImg = e.target.tagName === "IMG" ? e.target : e.target.closest("img");
      if (zoomBtn) {
        e.stopPropagation();
        let src = zoomBtn.getAttribute("data-img-src");
        let alt = zoomBtn.getAttribute("data-img-alt") || "Generated Visual";
        if (!src) {
          const card = zoomBtn.closest(".office-visual-image-card, .rendered-chart-container, .chat-bubble");
          const img = card ? card.querySelector("img") : null;
          if (img && img.src) {
            src = img.src;
            alt = img.alt || alt;
          }
        }
        if (src) {
          openImageZoomModal(src, alt);
        }
      } else if (clickedImg && clickedImg.src) {
        e.stopPropagation();
        openImageZoomModal(clickedImg.src, clickedImg.alt);
      }
    });
  }

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

  // Setup Document Intelligence Chips (Feature 1: Full Document Q&A)
  setupDocToolsChips();

  // Setup Selection Quick Toolbar Chips (Feature 3: Inline Rewrite Toolbar)
  setupSelectionChips();

  // Setup Transform Doc to Deck Feature
  initDocToDeckFeature();

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
  if (typeof window !== "undefined") {
    window.__isSummarizeSlidesAction = false;
    window.__lastUserPrompt = "";
  }
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
        <button class="quick-chip" id="chipGenImage" style="background-color:#fdf2f8; color:#be185d; border-color:#fbcfe8; font-weight:600;">🎨 Create Image</button>
        <button class="quick-chip" id="chipCreateChart" style="background-color:#f0fdf4; color:#15803d; border-color:#bbf7d0; font-weight:600;">📈 Create Chart</button>
        <button class="quick-chip" id="chipRisks" style="background-color:#fef3c7; color:#b45309; border-color:#fde68a;">⚠️ Key Risks</button>
        <button class="quick-chip" id="chipRewrite" style="background-color:#f3e8ff; color:#7e22ce; border-color:#e9d5ff;">🪄 Rewrite Slide</button>
        <button class="quick-chip" id="chipDocToDeck" style="background-color:#eef2ff; color:#4338ca; border-color:#c7d2fe; font-weight:600;">📄 Doc to Deck</button>
        <button class="quick-chip" id="chipShorten">📉 Shorten</button>
        <button class="quick-chip" id="chipTable">📊 Table</button>
      `;

      const cSummarize = document.getElementById("chipSummarize");
      if (cSummarize) cSummarize.onclick = () => runPowerPointSlideAction("takeaways");

      const cGenImage = document.getElementById("chipGenImage");
      if (cGenImage) {
        cGenImage.onclick = () => handleCreateImageClick();
      }

      const cCreateChart = document.getElementById("chipCreateChart");
      if (cCreateChart) {
        cCreateChart.onclick = () => handleCreateChartClick();
      }

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
        <button class="quick-chip" id="chipSummarize" style="background-color:#e0f2fe; color:#0369a1; border-color:#bae6fd;">📊 Summarize Slides</button>
        <button class="quick-chip" id="chipGenImage" style="background-color:#fdf2f8; color:#be185d; border-color:#fbcfe8; font-weight:600;">🎨 Create Image</button>
        <button class="quick-chip" id="chipCreateChart" style="background-color:#f0fdf4; color:#15803d; border-color:#bbf7d0; font-weight:600;">📈 Create Chart</button>
      `;

      const cDocToDeck = document.getElementById("chipDocToDeck");
      if (cDocToDeck) {
        cDocToDeck.onclick = () => {
          const fileInput = document.getElementById("docToDeckFileInput");
          if (fileInput) fileInput.click();
        };
      }

      const cSum = document.getElementById("chipSummarize");
      if (cSum) cSum.onclick = () => runPowerPointSlideAction("takeaways");

      const cGenImage = document.getElementById("chipGenImage");
      if (cGenImage) {
        cGenImage.onclick = () => handleCreateImageClick();
      }

      const cCreateChart = document.getElementById("chipCreateChart");
      if (cCreateChart) {
        cCreateChart.onclick = () => handleCreateChartClick();
      }
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

// Dynamic host selection polling is completely disabled to protect pointer event performance.
// Document and slide selections are read strictly on-demand when user clicks an action chip or submits a prompt.
function handleSelectionChanged() {
  return;
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
        appendAssistantBubble(aiResultText, data, userPrompt);

        const hasStructuredChart = /```(?:json|chart|pie|bar|column|line|doughnut|donut)?\s*\{[\s\S]*?(?:chartType|chart_type|pie|bar|column|line|doughnut)[\s\S]*?```/i.test(aiResultText) ||
                                   /\{\s*"(?:chartType|chart_type)"\s*:\s*"(?:pie|bar|line|doughnut|donut|column)"/i.test(aiResultText);
        const hasDistinctNonChartImageIntent = /(?:photo|photograph|portrait|illustration|logo|camera|scenery|picture of|image of a)/i.test((userPrompt || '').toLowerCase());

        let fullAiText = aiResultText;
        if (data && Array.isArray(data.images) && data.images.length > 0) {
          if (!hasStructuredChart || hasDistinctNonChartImageIntent) {
            for (const imgUrl of data.images) {
              if (imgUrl && !fullAiText.includes(imgUrl)) {
                fullAiText += `\n\n![Generated Image](${imgUrl})\n\n`;
              }
            }
          }
        }
        return parseMarkdown(fullAiText, { hasDistinctNonChartImageIntent });
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
  clearChatAndInput();
  const freshText = await hostAdapter.getSelectedText();
  if (freshText && freshText.trim().length > 0) {
    currentSelectedText = freshText.trim();
  }

  if (!currentSelectedText) {
    appendBubble(hostAdapter?.name === "PowerPoint" ? "Please select text or a slide first." : "Please highlight text in Word first.", "system");
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

export function isSubstantiveText(text) {
  if (!text || typeof text !== "string") return false;
  const cleaned = text
    .replace(/\(Visual\s*\/\s*Slide\s*content\)/gi, "")
    .replace(/\[Slide\s*\d+\]:?/gi, "")
    .replace(/[-•*#_`~|\s\r\n\t]/g, "")
    .trim();
  return /[a-zA-Z0-9]/.test(cleaned) && cleaned.length >= 3;
}

// Dedicated Image Generation Action from Selection or User Query
async function handleCreateImageClick() {
  clearChatAndInput();
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Checking selection...";
    loadingText.style.display = "block";
  }

  let selectedText = "";
  let slideCount = 0;

  try {
    if (hostAdapter && typeof hostAdapter.getSelectedText === "function") {
      const raw = await hostAdapter.getSelectedText();
      selectedText = isSubstantiveText(raw) ? raw.trim() : "";
    }
    if (hostAdapter && typeof hostAdapter.getSelectedSlidesText === "function") {
      const slides = await hostAdapter.getSelectedSlidesText();
      if (slides) slideCount = slides.length;
    }
  } catch (err) {
    console.warn("Could not read selection for image generation:", err);
  } finally {
    if (loadingText) loadingText.style.display = "none";
  }

  // 1. If substantive text is selected AND not multiple slides: generate image for that selected content
  if (selectedText && selectedText.length > 0 && slideCount <= 1) {
    const prompt = `Create a high-quality, professional image visual illustration representing the following selected content:

"""
${selectedText}
"""

CRITICAL INSTRUCTIONS:
1. Trigger the native image generation tool directly to produce the visual image.
2. DO NOT output raw JSON, visual_request schemas, or code blocks.
3. Generate an impactful, visually stunning illustration for this presentation topic.`;
    const snippet = selectedText.replace(/\s+/g, " ").trim().substring(0, 75);
    const displayUserBubble = `🎨 Generate image for selected text:\n"${snippet}${selectedText.length > 75 ? '...' : ''}"`;
    await executeGeminiWorkflow(prompt, displayUserBubble, null, { isolateSession: true });
    return;
  }

  // 2. If selected content is blank/null OR multiple slides are selected: ask the user what the image should be created for
  const reasonText = slideCount > 1
    ? `${slideCount} slides are currently selected.`
    : `No specific text is currently selected.`;

  const html = `
    <div>
      <div style="font-weight:600; color:#be185d; margin-bottom:4px;">🎨 Create Image</div>
      <div style="font-size:12px; margin-bottom:6px;">${reasonText} What should the image be created for?</div>
      <div style="font-size:11px; color:#605e5c; margin-bottom:8px;">Choose a suggested concept below or describe what you want in the input box:</div>
      <div style="display:flex; flex-wrap:wrap; gap:5px; margin-bottom:4px;">
        <button class="quick-chip img-opt-btn" style="cursor:pointer; background:#fdf2f8; color:#be185d; border-color:#fbcfe8;">🏬 Retail Storefront</button>
        <button class="quick-chip img-opt-btn" style="cursor:pointer; background:#fdf2f8; color:#be185d; border-color:#fbcfe8;">📦 Supply Chain Logistics</button>
        <button class="quick-chip img-opt-btn" style="cursor:pointer; background:#fdf2f8; color:#be185d; border-color:#fbcfe8;">👥 Executive Strategy Team</button>
        <button class="quick-chip img-opt-btn" style="cursor:pointer; background:#fdf2f8; color:#be185d; border-color:#fbcfe8;">🌱 Corporate Sustainability</button>
      </div>
    </div>
  `;

  appendInteractiveBubble(html, (bubble) => {
    const optButtons = bubble.querySelectorAll(".img-opt-btn");
    optButtons.forEach((btn) => {
      btn.addEventListener("click", async () => {
        const concept = btn.innerText.replace(/^[^\w]+/, "").trim();
        optButtons.forEach((b) => (b.disabled = true));
        const prompt = `Create a high-quality, professional image visual illustration of: ${concept}.

CRITICAL INSTRUCTIONS:
1. Trigger the native image generation tool directly to produce the visual image.
2. DO NOT output raw JSON, visual_request schemas, or code blocks.`;
        const displayBubble = `🎨 Generate image: "${concept}"`;
        await executeGeminiWorkflow(prompt, displayBubble, null, { isolateSession: true });
      });
    });

    const promptInput = document.getElementById("promptText");
    if (promptInput) {
      promptInput.value = "Create an image of ";
      promptInput.focus();
      if (promptInput.setSelectionRange) {
        promptInput.setSelectionRange(promptInput.value.length, promptInput.value.length);
      }
    }
  });
}

// Dedicated Chart Creation Action from Selection or User Guidance
async function handleCreateChartClick() {
  clearChatAndInput();
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Checking selection for chart...";
    loadingText.style.display = "block";
  }

  let selectedText = "";
  let slideCount = 0;

  try {
    if (hostAdapter && typeof hostAdapter.getSelectedText === "function") {
      const raw = await hostAdapter.getSelectedText();
      selectedText = isSubstantiveText(raw) ? raw.trim() : "";
    }
    if (hostAdapter && typeof hostAdapter.getSelectedSlidesText === "function") {
      const slides = await hostAdapter.getSelectedSlidesText();
      if (slides) slideCount = slides.length;
    }
  } catch (err) {
    console.warn("Could not read selection for chart:", err);
  } finally {
    if (loadingText) loadingText.style.display = "none";
  }

  // Case 1: Specific text IS selected (and not multiple slides)
  // Requirement: "If the text is selected, ask the user what type of chart and pass both to streamassist"
  if (selectedText && selectedText.length > 0 && slideCount <= 1) {
    const preview = selectedText.replace(/\s+/g, " ").trim().substring(0, 80);
    const html = `
      <div>
        <div style="font-weight:600; color:#15803d; margin-bottom:4px;">📈 Create Chart for Selected Content</div>
        <div style="font-size:11.5px; color:#605e5c; margin-bottom:8px; font-style:italic;">"${preview}${selectedText.length > 80 ? '...' : ''}"</div>
        <div style="font-size:12px; margin-bottom:8px;">What type of chart would you like to create?</div>
        <div style="display:flex; flex-wrap:wrap; gap:6px;">
          <button class="quick-chip chart-opt-btn" data-type="doughnut" style="background:#eff6fc; color:#0369a1; border-color:#bae6fd; font-weight:600; cursor:pointer;">🍩 Doughnut / Pie</button>
          <button class="quick-chip chart-opt-btn" data-type="bar" style="background:#f0fdf4; color:#15803d; border-color:#bbf7d0; font-weight:600; cursor:pointer;">📊 Bar Chart</button>
          <button class="quick-chip chart-opt-btn" data-type="column" style="background:#fef3c7; color:#b45309; border-color:#fde68a; font-weight:600; cursor:pointer;">📶 Column Chart</button>
          <button class="quick-chip chart-opt-btn" data-type="line" style="background:#f3e8ff; color:#7e22ce; border-color:#e9d5ff; font-weight:600; cursor:pointer;">📈 Line Chart</button>
        </div>
      </div>
    `;

    appendInteractiveBubble(html, (bubble) => {
      const buttons = bubble.querySelectorAll(".chart-opt-btn");
      buttons.forEach((btn) => {
        btn.addEventListener("click", async () => {
          const chartType = btn.getAttribute("data-type") || "bar";
          buttons.forEach((b) => (b.disabled = true));
          btn.style.borderColor = "#15803d";
          btn.style.backgroundColor = "#dcfce7";

          const prompt = `Create a high-resolution ${chartType} chart for PowerPoint visualizing the following selected data and context:
"""
${selectedText}
"""

Requirements:
1. Output a structured JSON code block with the exact chart specification:
\`\`\`json
{
  "chartType": "${chartType}",
  "title": "<Short Title (2-5 words max)>",
  "data": [
    { "label": "<Category/Metric>", "value": <NumericValue> }
  ]
}
\`\`\`
2. CRITICAL CONSTRAINT FOR TITLE:
   - The chart title MUST be short and punchy (maximum 2 to 5 words, e.g. "Coffee Output by State", "Q4 Revenue Share").
   - NEVER create long titles, never include parenthetical subtitles in the title, and never exceed 5 words!
3. Include 2 concise, executive takeaway bullet points analyzing the data under the chart.
4. CRITICAL SCOPE CONTRACT: Output content for EXACTLY ONE single slide. DO NOT regenerate, repeat, or expand upon previous presentation slides from earlier in this conversation under any circumstances.`;

          const displayBubble = `📈 [Create Chart] Generating ${chartType} chart from selected text...`;
          await executeGeminiWorkflow(prompt, displayBubble);
        });
      });
    });
    return;
  }

  // Case 2: No text is selected OR multiple slides are selected
  // Requirement: "If no text is selected or multiple slides selected, ask the user what type of chart (you can provide some options) and also ask about what."
  const notice = slideCount > 1
    ? `${slideCount} slides are currently selected.`
    : `No specific text is currently selected.`;

  const html = `
    <div>
      <div style="font-weight:600; color:#15803d; margin-bottom:4px;">📈 Create Chart</div>
      <div style="font-size:11.5px; color:#605e5c; margin-bottom:8px;">${notice} What type of chart would you like to create, and what data or topic should it visualize?</div>
      
      <div style="font-weight:600; font-size:11.5px; margin-bottom:4px; color:#323130;">1. Select Chart Type:</div>
      <div style="display:flex; flex-wrap:wrap; gap:6px; margin-bottom:10px;">
        <button class="quick-chip chart-type-btn" data-type="doughnut" style="background:#eff6fc; color:#0369a1; border-color:#bae6fd; font-weight:600; cursor:pointer;">🍩 Doughnut / Pie</button>
        <button class="quick-chip chart-type-btn" data-type="bar" style="background:#f0fdf4; color:#15803d; border-color:#bbf7d0; font-weight:600; cursor:pointer;">📊 Bar Chart</button>
        <button class="quick-chip chart-type-btn" data-type="column" style="background:#fef3c7; color:#b45309; border-color:#fde68a; font-weight:600; cursor:pointer;">📶 Column Chart</button>
        <button class="quick-chip chart-type-btn" data-type="line" style="background:#f3e8ff; color:#7e22ce; border-color:#e9d5ff; font-weight:600; cursor:pointer;">📈 Line Chart</button>
      </div>

      <div style="font-weight:600; font-size:11.5px; margin-bottom:4px; color:#323130;">2. Enter or Paste Data (e.g. from Table):</div>
      <textarea class="chart-data-input" placeholder="Paste table rows, numbers, or describe metrics (e.g. 'Revenue: Q1 10M, Q2 15M, Q3 25M')..." style="width:100%; box-sizing:border-box; border:1px solid #c7e0f4; border-radius:4px; padding:6px; font-size:11.5px; font-family:inherit; resize:vertical; min-height:50px; margin-bottom:6px;"></textarea>
      <div style="margin-bottom:10px;">
        <button class="chart-generate-btn" style="background:#15803d; color:#ffffff; border:none; border-radius:4px; padding:5px 12px; font-size:11.5px; font-weight:600; cursor:pointer; display:inline-flex; align-items:center; gap:4px;">✨ Generate Chart from Data</button>
      </div>

      <div style="font-weight:600; font-size:11.5px; margin-bottom:4px; color:#323130;">Or choose a quick topic:</div>
      <div style="display:flex; flex-wrap:wrap; gap:5px; margin-bottom:4px;">
        <button class="quick-chip chart-topic-btn" style="cursor:pointer;">🛒 Retail Market Share</button>
        <button class="quick-chip chart-topic-btn" style="cursor:pointer;">💰 Quarterly Revenue Growth</button>
        <button class="quick-chip chart-topic-btn" style="cursor:pointer;">📉 Expense & Cost Breakdown</button>
        <button class="quick-chip chart-topic-btn" style="cursor:pointer;">⚡ Regional Performance</button>
      </div>
    </div>
  `;

  appendInteractiveBubble(html, (bubble) => {
    let currentType = "bar";
    const typeButtons = bubble.querySelectorAll(".chart-type-btn");
    const topicButtons = bubble.querySelectorAll(".chart-topic-btn");
    const generateBtn = bubble.querySelector(".chart-generate-btn");
    const dataInput = bubble.querySelector(".chart-data-input");

    typeButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        currentType = btn.getAttribute("data-type") || "bar";
        typeButtons.forEach((b) => {
          b.style.borderColor = "";
          b.style.fontWeight = "600";
          b.style.backgroundColor = "";
        });
        btn.style.borderColor = "#15803d";
        btn.style.backgroundColor = "#dcfce7";

        const promptInput = document.getElementById("promptText");
        if (promptInput) {
          promptInput.value = `Create a ${currentType} chart of `;
          promptInput.focus();
          if (promptInput.setSelectionRange) {
            promptInput.setSelectionRange(promptInput.value.length, promptInput.value.length);
          }
        }
      });
    });

    if (generateBtn && dataInput) {
      generateBtn.addEventListener("click", async () => {
        const dataText = dataInput.value.trim();
        if (!dataText) {
          dataInput.focus();
          dataInput.style.borderColor = "#dc2626";
          return;
        }
        typeButtons.forEach((b) => (b.disabled = true));
        topicButtons.forEach((b) => (b.disabled = true));
        generateBtn.disabled = true;

        const prompt = `Create a high-resolution ${currentType} chart for PowerPoint visualizing the following data and metrics:
${dataText}

Requirements:
1. Output a structured JSON code block with the exact chart specification:
\`\`\`json
{
  "chartType": "${currentType}",
  "title": "<Short Title (2-5 words max)>",
  "data": [
    { "label": "<Category>", "value": <NumericValue> }
  ]
}
\`\`\`
2. CRITICAL CONSTRAINT FOR TITLE:
   - The chart title MUST be short and punchy (maximum 2 to 5 words).
   - NEVER create long titles, never include parenthetical subtitles in the title, and never exceed 5 words!
3. Include 2 concise, executive takeaway bullet points analyzing the data under the chart.
4. CRITICAL SCOPE CONTRACT: Output content for EXACTLY ONE single slide. DO NOT regenerate, repeat, or expand upon previous presentation slides from earlier in this conversation under any circumstances.`;

        const displayBubble = `📈 [Create Chart] Generating ${currentType} chart from provided data...`;
        await executeGeminiWorkflow(prompt, displayBubble);
      });
    }

    topicButtons.forEach((btn) => {
      btn.addEventListener("click", async () => {
        const topic = btn.innerText.replace(/^[^\w]+/, "").trim();
        typeButtons.forEach((b) => (b.disabled = true));
        topicButtons.forEach((b) => (b.disabled = true));
        if (generateBtn) generateBtn.disabled = true;

        const prompt = `Create a high-resolution ${currentType} chart for PowerPoint visualizing: ${topic}.

Requirements:
1. Output a structured JSON code block with the exact chart specification:
\`\`\`json
{
  "chartType": "${currentType}",
  "title": "${topic}",
  "data": [
    { "label": "<Category>", "value": <NumericValue> }
  ]
}
\`\`\`
2. CRITICAL CONSTRAINT FOR TITLE:
   - The chart title MUST be short and punchy (maximum 2 to 5 words).
   - NEVER create long titles, never include parenthetical subtitles in the title, and never exceed 5 words!
3. Include 2 concise, executive takeaway bullet points analyzing the data under the chart.
4. CRITICAL SCOPE CONTRACT: Output content for EXACTLY ONE single slide. DO NOT regenerate, repeat, or expand upon previous presentation slides from earlier in this conversation under any circumstances.`;

        const displayBubble = `📈 [Create Chart] Generating ${currentType} chart for "${topic}"...`;
        await executeGeminiWorkflow(prompt, displayBubble);
      });
    });

    const promptInput = document.getElementById("promptText");
    if (promptInput) {
      promptInput.value = `Create a ${currentType} chart of `;
      promptInput.focus();
      if (promptInput.setSelectionRange) {
        promptInput.setSelectionRange(promptInput.value.length, promptInput.value.length);
      }
    }
  });
}

async function runDocIntelligencePrompt(instruction) {
  clearChatAndInput();
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
  clearChatAndInput();
  const loadingText = document.getElementById("loading");
  if (loadingText) {
    loadingText.innerText = "⚡ Reading slide context...";
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
  let slideLabel = "presentation deck";
  let displayBubble = "";

  if (actionType === "summarize" || actionType === "takeaways") {
    if (typeof window !== "undefined") window.__isSummarizeSlidesAction = true;
    if (slides && slides.length >= 2) {
      slideLabel = `${slides.length} highlighted slides`;
      slideContext = slides.map((s) => `[Highlighted Slide ${s.slideNumber}]:\n${s.text || "(No readable text)"}`).join("\n\n---\n\n");
      displayBubble = `📊 [Summarize Slides] Generating executive summary presentation (up to 5 slides) from ${slides.length} highlighted slides...`;
    } else {
      appendBubble(
        "ℹ️ **Summarize Slides**: Analyzing all slides across the entire presentation to generate a comprehensive deck summary (up to 5 slides)...",
        "assistant"
      );
      const fullDocText = await hostAdapter.getFullDocumentText();
      slideContext = fullDocText || (slides && slides.length === 1 ? `[Slide ${slides[0].slideNumber}]:\n${slides[0].text}` : "");
      slideLabel = "entire presentation deck";
      displayBubble = `📊 [Summarize Slides] Summarizing entire presentation deck (up to 5 slides)...`;
    }
  } else {
    if (typeof window !== "undefined") window.__isSummarizeSlidesAction = false;
    if (slides && slides.length > 0) {
      slideLabel = slides.length === 1 ? `Slide ${slides[0].slideNumber || 1}` : `${slides.length} highlighted slides`;
      slideContext = slides.map((s) => `[Highlighted Slide ${s.slideNumber || 1}]:\n${s.text || "(No readable text)"}`).join("\n\n---\n\n");
    } else {
      const fullDocText = await hostAdapter.getFullDocumentText();
      if (fullDocText && fullDocText.length > 10) {
        slideContext = fullDocText;
        slideLabel = "presentation deck";
      }
    }
  }

  let taskInstruction = "";

  switch (actionType) {
    case "risks":
      taskInstruction = `Based on the context from the ${slideLabel} provided below, extract and analyze all Key Strategic & Operational Risks, Blockers, Gaps, and Dependencies. Create content for a new PowerPoint slide titled "⚠️ Key Risks & Mitigations" with high-impact bullet points, severity assessments, and concrete mitigation actions.`;
      displayBubble = `⚠️ [Key Risks] Generating new risk analysis slide from ${slideLabel}...`;
      break;
    case "summarize":
    case "takeaways":
      taskInstruction = `Based strictly on the content from the ${slideLabel} provided below, create a comprehensive executive summary presentation (strictly between 3 and 5 slides maximum).
CRITICAL CLOSED-BOOK GROUNDING CONTRACT:
1. STRICT BOUNDARY: You are operating in 100% CLOSED-BOOK MODE. You must synthesize information ONLY AND EXCLUSIVELY from the text and data explicitly present in the selected slides below.
2. ZERO EXTERNAL KNOWLEDGE / ZERO HALLUCINATIONS: Do NOT introduce outside industry context, external market facts, assumptions, or topics that are not explicitly stated in the selected slides below. If a point or metric is not directly written in the provided slide text, omit it completely.
3. MULTI-SLIDE STRUCTURE & FORMAT (Strictly between 3 to 5 slides maximum):
   - You must generate AT MOST 5 slides total (strictly 3 to 5 slides). NEVER generate 6 or more slides.
   - Separate distinct topics, tables, and visual charts into their own slides (e.g. ## Slide 1: [Executive Overview / Main Metrics Table], ## Slide 2: [Category Breakdown / Visual Chart / Details Table], etc.).
   - If there are multiple tables or data sets, place each table on its own appropriate slide.
   - If a visual chart represents data, output a structured JSON code block with the exact data metrics (chartType: "doughnut" or "bar", title: "...", data: [...]) so our client presentation engine can render a crisp chart.
   - Break down the key takeaways into a dedicated single slide titled "## 📊 Executive Summary: Key Takeaways" with at most 3 to 4 concise executive bullet points (max 15 words per bullet). CRITICAL: Generate strictly at most 4 bullet points; do NOT exceed 4 bullet points or create continuation slides.
   - Format each slide with a clear markdown header (## Slide 1: [Title], ## Slide 2: [Title], etc.) so each section generates its own slide.
4. DO NOT output conversational preamble.`;
      if (!displayBubble) {
        displayBubble = `📊 [Summarize Slides] Generating executive summary presentation (up to 5 slides) from ${slideLabel}...`;
      }
      break;
    case "action_items":
      taskInstruction = `Based on the context from the ${slideLabel} provided below, extract all Action Items, Deliverables, Next Steps, and Ownership. Create content for a new PowerPoint slide titled "✅ Action Items & Next Steps" with actionable task bullets, suggested owners, and priority levels.`;
      displayBubble = `✅ [Action Items] Generating action items slide from ${slideLabel}...`;
      break;
  }

  let fullPrompt = "";
  if (slideContext && slideContext.length > 5) {
    fullPrompt = `CRITICAL SLIDE CONTEXT:\nThe following is the actual content extracted from the ${slideLabel}. You MUST base your summary strictly and exclusively on this content:\n"""\n${slideContext.substring(0, 50000)}\n"""\n\nTask: ${taskInstruction}`;
  } else {
    fullPrompt = taskInstruction;
  }

  // Apply PowerPoint prompt enhancer rules
  const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
  fullPrompt = enhancePromptForPowerPoint(fullPrompt);

  await executeGeminiWorkflow(fullPrompt, displayBubble);
}

async function callGeminiProxy(customPrompt = null, ignoreSelection = false, customDisplayBubble = null) {
  const promptInput = document.getElementById("promptText");
  const userText = customPrompt || (promptInput ? promptInput.value.trim() : "");

  let selectedText = "";
  if (!ignoreSelection && hostAdapter) {
    selectedText = await hostAdapter.getSelectedText();
    if (!isSubstantiveText(selectedText)) {
      selectedText = "";
    }
  }

  if (!userText && !selectedText) {
    appendBubble(`Please type a prompt or select content in ${hostAdapter ? hostAdapter.name : 'Office'} first.`, "system");
    return;
  }

  // Intercept generic Image or Chart requests in PowerPoint to prompt user interactively
  if (userText && hostAdapter && hostAdapter.name === "PowerPoint") {
    const trimmedLower = userText.trim().toLowerCase();
    const isGenericImage = /^(?:please\s+)?(?:generate|create|make|insert|add)\s+(?:an?\s+)?image\s*\.?$/i.test(trimmedLower);
    const isGenericChart = /^(?:please\s+)?(?:generate|create|make|insert|add)\s+(?:an?\s+)?(?:chart|graph|plot)\s*\.?$/i.test(trimmedLower);

    if (isGenericImage) {
      if (promptInput) promptInput.value = "";
      await handleCreateImageClick();
      return;
    }

    if (isGenericChart) {
      if (promptInput) promptInput.value = "";
      await handleCreateChartClick();
      return;
    }
  }

  let fullPrompt = "";
  let displayUserBubble = "";

  if (selectedText && userText) {
    fullPrompt = `Selected Document Context:\n"${selectedText}"\n\nUser Instruction: ${userText}`;
    displayUserBubble = customDisplayBubble || `📌 Context: "${selectedText.substring(0, 70)}${selectedText.length > 70 ? '...' : ''}"\n\n${userText}`;
  } else if (selectedText && !userText) {
    fullPrompt = `Please analyze, summarize, or explain the following selected text:\n"${selectedText}"`;
    displayUserBubble = `📌 Selected Text:\n"${selectedText.substring(0, 90)}${selectedText.length > 90 ? '...' : ''}"`;
  } else {
    fullPrompt = userText;
    displayUserBubble = customDisplayBubble || userText;
  }

  if (promptInput) promptInput.value = "";
  
  if (typeof window !== "undefined") {
    window.__lastUserPrompt = userText || fullPrompt || "";
    window.__isSummarizeSlidesAction = /^\s*(?:please\s+)?summarize\s+(?:the\s+)?(?:slides?|deck|presentation)\b/i.test(userText);
  }

  if (hostAdapter && hostAdapter.name === "PowerPoint") {
    const { enhancePromptForPowerPoint } = await import('../adapters/ppt/promptEnhancer.js');
    fullPrompt = enhancePromptForPowerPoint(fullPrompt);
  }

  await executeGeminiWorkflow(fullPrompt, displayUserBubble);
}

// Feature: Transform Doc to Deck (Upload document -> PowerPoint presentation)
function initDocToDeckFeature() {
  const fileInput = document.getElementById("docToDeckFileInput") || document.getElementById("docToDeckInput");
  if (!fileInput) return;

  fileInput.addEventListener("change", async (event) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    if (typeof window !== "undefined") {
      window.__isSummarizeSlidesAction = false;
    }

    clearChatAndInput();
    const loadingText = document.getElementById("loading");
    if (loadingText) {
      loadingText.innerText = `⚡ Document(s) uploaded. Generating presentation summary deck...`;
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

      appendBubble(
        `📄 **Doc to Deck**: Uploaded \`${fileNames.join(', ')}\`.\n\nA comprehensive presentation summary is being created from your document...`,
        "assistant"
      );

      const displayUserBubble = `📄 [Doc to Deck] Uploaded: ${fileNames.join(', ')}`;
      const prompt = `Extract the key takeaways from the attached document(s) and create a structured executive presentation slide deck with a total of up to 6 slides (1 Title/Intro Slide + 5 Content Slides).

PowerPoint Slide Deck Requirements:
1. SLIDE DECK STRUCTURE (MAXIMUM 6 SLIDES TOTAL):
   - ## Slide 1: [Relevant Emoji] [Presentation Title (3-4 words max)]
     ### [Compelling Subtitle / Executive Orientation]
     Slide 1 MUST be a dedicated TITLE & EXECUTIVE SUMMARY SLIDE containing ONLY:
     * Presentation Title: A crisp, impactful title capturing the overarching subject.
     * Subtitle: An executive subtitle setting the strategic context.
     * Executive Summary: Exactly ONE small, high-impact executive summary paragraph (2 to 4 sentences max) explaining what the document is all about and its key strategic value.
     * STRICT RULES FOR SLIDE 1:
       - ONLY include Title, Subtitle, and the ONE small Executive Summary paragraph.
       - NEVER include a "Deck Scope & Outline" section!
       - NEVER list upcoming slides (e.g. "Slide 2: ...", "Slide 3: ...")!
       - NEVER include granular operational bullets, department trivia, metrics, or tables on Slide 1!
   - ## Slide 2 to Slide 6 (5 Content Slides):
     5 focused executive content slides breaking down the core insights, findings, data, and recommendations from the document.
2. Provide EXACTLY ONE presentation deck (Slide 1 Title & Summary Slide + 5 Content Slides, maximum 6 slides total). DO NOT output multiple alternative options.
3. DO NOT output conversational preamble or filler (e.g. "Here is...", "Sure!"). Output the slide deck content directly.
4. For each content slide (Slides 2-6), structure with rich executive visual hierarchy:
   - "## Slide <N>: <Emoji> <Punchy Slide Title (3-4 words max)>"
   - "### <Crisp Subtitle / Strategic Takeaway>"
   - Format slide content using clean, professional formats:
     * Qualitative / Strategic Slides (Vision, Market Drivers, Operational Pillars, Roadmap): 3 to 4 high-impact bullet points with bold lead-ins (• **Strategic Pillar**: Detailed description).
     * Quantitative / Comparative Slides (Financials, KPIs, Decarbonization, Operational Metrics): When presenting dense metrics, performance targets, or before/after comparisons, use a clean Markdown table (| Dimension / Metric | Baseline / Prior | Target / Strategic Impact |). Follow the table with 1 to 2 concise takeaway bullets. (Limit tables to at most 1 or 2 slides in the deck).
     * Optional Visual Breakdown Slide (Only if document has market share or volume distribution): A clean high-resolution chart JSON block (\`\`\`json { "chartType": "doughnut", "title": "<Short Title (2-5 words max)>", "data": [...] } \`\`\`) accompanied by 2 narrative bullets.
5. CRITICAL CHART TITLE RULE: Any chart title MUST be short and punchy (2 to 5 words max).
6. Do NOT output internal design metadata like "Visual Concept:", "Color:", or font sizes.
7. Do NOT output pseudocode visual labels like "Metric Grid" or "Comparison Card" in plain text. Format content cleanly as narrative bullets, structured tables, or chart JSON.
7. Separate each slide cleanly with "---".`;

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

async function executeGeminiWorkflow(fullPrompt, displayUserBubble, attachments = null, options = {}) {
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

    // Detect if this is an explicit image creation intent
    const isImageIntent = Boolean(
      (options && options.isolateSession) ||
      (displayUserBubble && /Generate image|Create image|Insert image/i.test(displayUserBubble)) ||
      (fullPrompt && /\b(create|generate|make|draw|show|produce|render|provide|insert|add)\s+(an?\s+)?(image|picture|photo|illustration|graphic|visual|artwork)\b/i.test(fullPrompt)) ||
      (fullPrompt && /\b(image|picture|photo|illustration|graphic|visual)\s+(of|for|showing|depicting)\b/i.test(fullPrompt))
    );

    if (typeof window !== "undefined") {
      window.__lastUserPrompt = fullPrompt || displayUserBubble || "";
      if (isImageIntent) {
        window.__isSummarizeSlidesAction = false;
      }
    }

    // If image generation, isolate from prior conversation session to prevent JSON formatting and previous visual contagion
    const sessionIdToUse = isImageIntent ? null : currentSessionId;
    const historyToUse = isImageIntent ? [] : chatHistoryState;

    let data = await askGeminiEnterprise(fullPrompt, historyToUse, sessionIdToUse, true, attachments);

    // Fallback Safety Net: Check if model emitted a raw visual_request JSON block instead of calling Imagen
    let aiResponse = data.result || "No content returned.";
    if (aiResponse.includes('"visual_request"')) {
      try {
        const jsonMatch = aiResponse.match(/\{[\s\S]*?"visual_request"[\s\S]*?\}/);
        if (jsonMatch) {
          const parsed = JSON.parse(jsonMatch[0]);
          const extractedPrompt = parsed?.visual_request?.prompt;
          if (extractedPrompt && typeof extractedPrompt === "string") {
            console.log("[ImageFallback] Detected visual_request schema. Automatically invoking image generator with extracted prompt:", extractedPrompt.slice(0, 80));
            if (loadingText) {
              loadingText.innerText = "🎨 Generating image from visual specification...";
              loadingText.style.display = "block";
            }
            const fallbackPrompt = `Create a high-quality, professional image visual illustration of: ${extractedPrompt}`;
            const fallbackData = await askGeminiEnterprise(fallbackPrompt, [], null, true);
            if (fallbackData && (fallbackData.result || (fallbackData.images && fallbackData.images.length > 0))) {
              data = fallbackData;
              aiResponse = fallbackData.result || aiResponse;
            }
          }
        }
      } catch (fbErr) {
        console.warn("[ImageFallback] Error handling visual_request fallback:", fbErr);
      }
    }

    if (data.sessionId) {
      currentSessionId = data.sessionId;
    }
    if (Array.isArray(data.history)) {
      chatHistoryState = data.history;
    }

    const isSummarizeFlow = Boolean(
      (typeof window !== "undefined" && window.__isSummarizeSlidesAction) ||
      (displayUserBubble && displayUserBubble.includes("[Summarize Slides]"))
    );
    appendAssistantBubble(aiResponse, data, displayUserBubble || fullPrompt, { isSummarizeSlides: isSummarizeFlow });

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

function appendInteractiveBubble(htmlContent, onMountCallback = null) {
  const historyDiv = document.getElementById("chatHistory");
  if (!historyDiv) return null;
  const bubble = document.createElement("div");
  bubble.className = "chat-bubble assistant";
  bubble.innerHTML = htmlContent;
  historyDiv.appendChild(bubble);
  if (typeof onMountCallback === "function") {
    onMountCallback(bubble);
  }
  historyDiv.scrollTop = historyDiv.scrollHeight;
  return bubble;
}

/**
 * Opens an expanded high-resolution review modal for any generated visual image.
 * Provides explicit actions to review and decide to insert directly onto the current slide
 * or as a brand-new presentation slide.
 */
export function openImageZoomModal(imgSrc, altText = "Generated Visual") {
  if (!imgSrc) return;

  let modal = document.getElementById("imageZoomModal");
  if (!modal) {
    modal = document.createElement("div");
    modal.id = "imageZoomModal";
    modal.className = "image-zoom-overlay";
    document.body.appendChild(modal);
  }

  const isPPT = hostAdapter && hostAdapter.name === "PowerPoint";

  modal.innerHTML = `
    <div class="image-zoom-header">
      <div style="display:flex; align-items:center; gap:8px;">
        <span>🔍 Image Review</span>
        <div class="zoom-level-controls" style="display:inline-flex; gap:4px; background:rgba(255,255,255,0.15); border-radius:4px; padding:2px;">
          <button type="button" class="zoom-ctrl-btn" id="zoomCtrlFit" title="Fit to window" style="background:none; border:none; color:#cbd5e1; font-size:11px; padding:2px 7px; border-radius:3px; cursor:pointer;">1x Fit</button>
          <button type="button" class="zoom-ctrl-btn active" id="zoomCtrl2x" title="2x Enlarge" style="background:#0078d4; border:none; color:#ffffff; font-weight:bold; font-size:11px; padding:2px 7px; border-radius:3px; cursor:pointer;">2x Zoom</button>
          <button type="button" class="zoom-ctrl-btn" id="zoomCtrl3x" title="3x Extra Large" style="background:none; border:none; color:#cbd5e1; font-size:11px; padding:2px 7px; border-radius:3px; cursor:pointer;">3x Max</button>
        </div>
      </div>
      <button class="image-zoom-close" id="closeZoomModalBtn" title="Close preview (Esc)">✕</button>
    </div>
    <div class="image-zoom-body zoom-2x" id="zoomModalBody">
      <img src="${imgSrc}" alt="${altText}" />
    </div>
    <div class="image-zoom-footer">
      ${isPPT ? `
        <button class="image-zoom-btn" id="zoomInsertCurrentBtn">📌 Insert on Current Slide</button>
        <button class="image-zoom-btn secondary" id="zoomInsertNewBtn">➕ Insert as New Slide</button>
      ` : `
        <button class="image-zoom-btn" id="zoomInsertDocBtn">📌 Insert Image</button>
      `}
      <button class="image-zoom-btn secondary" id="zoomCancelBtn">Close</button>
    </div>
  `;

  modal.style.display = "flex";

  const bodyEl = document.getElementById("zoomModalBody");
  const btnFit = document.getElementById("zoomCtrlFit");
  const btn2x = document.getElementById("zoomCtrl2x");
  const btn3x = document.getElementById("zoomCtrl3x");

  const setZoom = (level) => {
    if (!bodyEl) return;
    bodyEl.classList.remove("zoom-fit", "zoom-2x", "zoom-3x");
    bodyEl.classList.add(`zoom-${level}`);
    [btnFit, btn2x, btn3x].forEach(b => {
      if (b) {
        b.style.background = "none";
        b.style.color = "#cbd5e1";
        b.style.fontWeight = "normal";
      }
    });
    const activeBtn = level === "fit" ? btnFit : (level === "2x" ? btn2x : btn3x);
    if (activeBtn) {
      activeBtn.style.background = "#0078d4";
      activeBtn.style.color = "#ffffff";
      activeBtn.style.fontWeight = "bold";
    }
  };

  if (btnFit) btnFit.onclick = (e) => { e.stopPropagation(); setZoom("fit"); };
  if (btn2x) btn2x.onclick = (e) => { e.stopPropagation(); setZoom("2x"); };
  if (btn3x) btn3x.onclick = (e) => { e.stopPropagation(); setZoom("3x"); };

  // Smooth drag-to-pan in enlarged review mode
  let isDown = false;
  let startX = 0;
  let startY = 0;
  let scrollLeft = 0;
  let scrollTop = 0;
  if (bodyEl) {
    bodyEl.onmousedown = (e) => {
      if (e.target.tagName === "BUTTON") return;
      isDown = true;
      startX = e.pageX - bodyEl.offsetLeft;
      startY = e.pageY - bodyEl.offsetTop;
      scrollLeft = bodyEl.scrollLeft;
      scrollTop = bodyEl.scrollTop;
    };
    bodyEl.onmouseleave = () => { isDown = false; };
    bodyEl.onmouseup = () => { isDown = false; };
    bodyEl.onmousemove = (e) => {
      if (!isDown) return;
      e.preventDefault();
      const x = e.pageX - bodyEl.offsetLeft;
      const y = e.pageY - bodyEl.offsetTop;
      const walkX = (x - startX) * 1.5;
      const walkY = (y - startY) * 1.5;
      bodyEl.scrollLeft = scrollLeft - walkX;
      bodyEl.scrollTop = scrollTop - walkY;
    };
  }

  const closeModal = () => {
    modal.style.display = "none";
  };

  const closeBtn = document.getElementById("closeZoomModalBtn");
  if (closeBtn) closeBtn.onclick = closeModal;

  const cancelBtn = document.getElementById("zoomCancelBtn");
  if (cancelBtn) cancelBtn.onclick = closeModal;

  modal.onclick = (e) => {
    if (e.target === modal) {
      closeModal();
    }
  };

  const handleKeydown = (e) => {
    if (e.key === "Escape") {
      closeModal();
      window.removeEventListener("keydown", handleKeydown);
    }
  };
  window.addEventListener("keydown", handleKeydown);

  if (isPPT) {
    const insertCurrentBtn = document.getElementById("zoomInsertCurrentBtn");
    if (insertCurrentBtn) {
      insertCurrentBtn.onclick = async () => {
        insertCurrentBtn.innerText = "⏳ Inserting...";
        insertCurrentBtn.disabled = true;
        try {
          await hostAdapter.insertContent(`<img src="${imgSrc}" />`, `![Visual](${imgSrc})`, {
            mode: "insert_current",
            imageOnly: true
          });
          insertCurrentBtn.innerText = "✅ Inserted!";
          insertCurrentBtn.classList.add("success");
          setTimeout(() => { closeModal(); }, 1000);
        } catch (err) {
          console.error("Zoom insert error:", err);
          insertCurrentBtn.innerText = "❌ Failed to insert";
          setTimeout(() => {
            insertCurrentBtn.innerText = "📌 Insert on Current Slide";
            insertCurrentBtn.disabled = false;
          }, 2000);
        }
      };
    }

    const insertNewBtn = document.getElementById("zoomInsertNewBtn");
    if (insertNewBtn) {
      insertNewBtn.onclick = async () => {
        insertNewBtn.innerText = "⏳ Creating slide...";
        insertNewBtn.disabled = true;
        try {
          await hostAdapter.insertContent(`<img src="${imgSrc}" />`, `![Visual](${imgSrc})`, {
            mode: "insert",
            imageOnly: true
          });
          insertNewBtn.innerText = "✅ Created!";
          insertNewBtn.classList.add("success");
          setTimeout(() => { closeModal(); }, 1000);
        } catch (err) {
          console.error("Zoom insert as new slide error:", err);
          insertNewBtn.innerText = "❌ Failed";
          setTimeout(() => {
            insertNewBtn.innerText = "➕ Insert as New Slide";
            insertNewBtn.disabled = false;
          }, 2000);
        }
      };
    }
  } else {
    const insertDocBtn = document.getElementById("zoomInsertDocBtn");
    if (insertDocBtn) {
      insertDocBtn.onclick = async () => {
        insertDocBtn.innerText = "⏳ Inserting...";
        insertDocBtn.disabled = true;
        try {
          await performDocumentInsertion(`<img src="${imgSrc}" />`, `![Generated Image](${imgSrc})`, "insert_cursor");
          insertDocBtn.innerText = "✅ Inserted!";
          setTimeout(() => { closeModal(); }, 1000);
        } catch (err) {
          console.error("Zoom insert doc error:", err);
          insertDocBtn.innerText = "❌ Failed";
          setTimeout(() => {
            insertDocBtn.innerText = "📌 Insert Image";
            insertDocBtn.disabled = false;
          }, 2000);
        }
      };
    }
  }
}

function appendAssistantBubble(text, apiData = null, originalPrompt = "", bubbleOptions = {}) {
  const historyDiv = document.getElementById("chatHistory");
  if (!historyDiv) return;
  const bubble = document.createElement("div");
  bubble.className = "chat-bubble assistant";

  const isSummarizeAction = Boolean(
    bubbleOptions?.isSummarizeSlides ||
    (typeof window !== "undefined" && window.__isSummarizeSlidesAction) ||
    (originalPrompt && (
      originalPrompt.toLowerCase().includes("[summarize slides]") ||
      (isPPT && /\bsummariz(?:e|ing)\s+(?:the\s+|all\s+|these\s+)?(?:slides?|deck|presentation)\b/i.test(originalPrompt))
    ))
  );
  bubble.dataset.isSummarizeSlides = isSummarizeAction ? "true" : "false";

  // Check if response contains structured chart JSON
  const hasStructuredChart = /```(?:json|chart|pie|bar|column|line|doughnut|donut)?\s*\{[\s\S]*?(?:chartType|chart_type|pie|bar|column|line|doughnut)[\s\S]*?```/i.test(text) ||
                             /\{\s*"(?:chartType|chart_type)"\s*:\s*"(?:pie|bar|line|doughnut|donut|column)"/i.test(text);

  // Check if user requested distinct non-chart image content (e.g. photo, illustration, logo)
  const promptLower = (originalPrompt || "").toLowerCase();
  const hasDistinctNonChartImageIntent = /(?:image|picture|photo|photograph|portrait|illustration|drawing|cat|dog|logo|camera|scenery)/i.test(promptLower);

  let fullText = text;
  // If apiData has images:
  // RULE: If structured chart is present and user did NOT ask for a distinct non-chart image,
  // do NOT append duplicate lower-res tool chart attachments from StreamAssist.
  // BUT if there is distinct non-chart image content, DO append it.
  if (apiData && Array.isArray(apiData.images) && apiData.images.length > 0) {
    if (!hasStructuredChart || hasDistinctNonChartImageIntent) {
      for (const imgUrl of apiData.images) {
        if (imgUrl && !fullText.includes(imgUrl)) {
          fullText += `\n\n![Generated Image](${imgUrl})\n\n`;
        }
      }
    }
  }

  // Parse markdown into executive HTML
  const formattedHtml = parseMarkdown(fullText, { hasDistinctNonChartImageIntent });

  const textDiv = document.createElement("div");
  textDiv.innerHTML = formattedHtml;
  bubble.appendChild(textDiv);

  // DOM SWEEP: Ensure EVERY <img> element in the assistant response has zoom controls and click-to-zoom
  const renderedImages = textDiv.querySelectorAll("img");
  renderedImages.forEach((img) => {
    img.classList.add("office-preview-img");
    img.style.cursor = "pointer";
    img.title = "Click to zoom / review image";

    // Direct click on the image itself ALWAYS opens zoom modal
    img.onclick = (e) => {
      e.stopPropagation();
      openImageZoomModal(img.src, img.alt);
    };

    // If img parent doesn't have an overlay zoom button:
    const parentContainer = img.parentElement;
    if (parentContainer && !parentContainer.querySelector(".img-zoom-btn")) {
      if (getComputedStyle(parentContainer).position === "static") {
        parentContainer.style.position = "relative";
      }
      const zoomOverlay = document.createElement("button");
      zoomOverlay.type = "button";
      zoomOverlay.className = "img-zoom-btn";
      zoomOverlay.title = "Zoom and review image";
      zoomOverlay.innerHTML = "🔍 Zoom";
      zoomOverlay.onclick = (e) => {
        e.stopPropagation();
        openImageZoomModal(img.src, img.alt);
      };
      parentContainer.appendChild(zoomOverlay);
    }

    // Set data attributes and wire direct click handlers on any zoom buttons inside the card
    const card = img.closest(".office-visual-image-card, .rendered-chart-container, .chat-bubble");
    if (card) {
      card.querySelectorAll(".img-zoom-btn, .img-action-btn-zoom").forEach(btn => {
        btn.setAttribute("data-img-src", img.src);
        btn.setAttribute("data-img-alt", img.alt || "Generated Visual");
        btn.onclick = (e) => {
          e.stopPropagation();
          openImageZoomModal(img.src, img.alt);
        };
      });
    }

    // If there is no bottom "Zoom & Review" action button in the card or immediate container:
    if (card && !card.querySelector(".img-action-btn-zoom")) {
      const zoomReviewBar = document.createElement("div");
      zoomReviewBar.style.cssText = "margin:8px 0; display:flex; justify-content:center; gap:6px;";
      const zoomReviewBtn = document.createElement("button");
      zoomReviewBtn.type = "button";
      zoomReviewBtn.className = "img-action-btn-zoom";
      zoomReviewBtn.setAttribute("data-img-src", img.src);
      zoomReviewBtn.setAttribute("data-img-alt", img.alt || "Generated Visual");
      zoomReviewBtn.innerHTML = "🔍 Zoom & Review";
      zoomReviewBtn.onclick = (e) => {
        e.stopPropagation();
        openImageZoomModal(img.src, img.alt);
      };
      zoomReviewBar.appendChild(zoomReviewBtn);

      if (img.parentElement && img.parentElement.parentNode) {
        img.parentElement.parentNode.insertBefore(zoomReviewBar, img.parentElement.nextSibling);
      } else {
        card.appendChild(zoomReviewBar);
      }
    }
  });

  // Feature 2 & Multi-turn: Action Toolbar & Refinement Chips
  const actionsContainer = document.createElement("div");
  actionsContainer.className = "response-actions-container";

  // Primary Actions: Replace / Insert / Copy
  const primaryActions = document.createElement("div");
  primaryActions.className = "primary-actions";

  const hostName = hostAdapter ? hostAdapter.name : "Word";
  const isPPT = hostName === "PowerPoint";
  const isExcel = hostName === "Excel";

  // Check if this response was generated from an image prompt or contains rendered visuals
  const isImageRequest = hasDistinctNonChartImageIntent ||
                         /(?:generate|create|make|draw|show|render|provide|insert|add)\s+(?:an?\s+)?(?:image|picture|photo|visual|illustration|graphic|infographic)/i.test(promptLower) ||
                         /(?:image|picture|photo|visual|illustration)\s+(?:of|for|showing|depicting)/i.test(promptLower);

  const shouldInsertImageOnly = renderedImages.length > 0 && (isImageRequest || (!fullText.includes("##") && textDiv.innerText.trim().length < 60));
  const isBubbleSummarize = bubble.dataset.isSummarizeSlides === "true";

  // Helper: Retrieve the primary non-chart visual image from renderedImages
  const getTargetVisualImage = () => {
    if (renderedImages.length === 0) return null;
    const nonChartImg = Array.from(renderedImages).reverse().find(img => {
      const isInsideChart = img.closest(".rendered-chart-container, [data-chart-title]");
      const isChartSrc = (img.src || "").toLowerCase().includes("chart");
      const isChartAlt = (img.alt || "").toLowerCase().includes("chart");
      return !isInsideChart && !isChartSrc && !isChartAlt;
    });
    return nonChartImg || renderedImages[renderedImages.length - 1] || renderedImages[0];
  };

  // 1. In-Place Replace Button (Word & Excel only; completely omitted in PowerPoint)
  let replaceBtn = null;
  if (!isPPT) {
    replaceBtn = document.createElement("button");
    replaceBtn.className = "action-btn replace";
    replaceBtn.innerHTML = shouldInsertImageOnly ? `🔄 Replace Image` : (isExcel ? `🔄 Replace in Sheet` : `🔄 Replace in Doc`);
    replaceBtn.title = shouldInsertImageOnly ? "Replace active document with image" : (isExcel ? "Replace active sheet content" : "Replace active draft or selection in Word");
    replaceBtn.onclick = async () => {
      if (shouldInsertImageOnly) {
        const img = getTargetVisualImage();
        if (img && img.src) {
          await performDocumentInsertion(`<img src="${img.src}" />`, `![Image](${img.src})`, "replace_draft", { imageOnly: true, isSummarizeSlides: false });
        }
      } else {
        await performDocumentInsertion(textDiv.innerHTML, fullText, "replace_draft", { isSummarizeSlides: isBubbleSummarize });
      }
    };
  }

  // 2. Insert on Current Slide Button (PowerPoint only)
  let insertCurrentBtn = null;
  if (isPPT) {
    insertCurrentBtn = document.createElement("button");
    insertCurrentBtn.className = "action-btn insert-current";
    insertCurrentBtn.innerHTML = shouldInsertImageOnly ? `📌 Insert Image on Slide` : `📌 Insert on Current Slide`;
    insertCurrentBtn.title = shouldInsertImageOnly ? "Insert image directly onto current slide" : "Insert generated content or image directly onto the currently active slide";
    insertCurrentBtn.onclick = async () => {
      if (shouldInsertImageOnly) {
        const img = getTargetVisualImage();
        if (img && img.src) {
          await performDocumentInsertion(`<img src="${img.src}" />`, `![Image](${img.src})`, "insert_current_slide", { imageOnly: true, isSummarizeSlides: false });
        }
      } else {
        await performDocumentInsertion(textDiv.innerHTML, fullText, "insert_current_slide", { isSummarizeSlides: isBubbleSummarize });
      }
    };
  }

  // 3. Insert as New Slide Button
  const insertBtn = document.createElement("button");
  insertBtn.className = "action-btn insert";
  insertBtn.innerHTML = shouldInsertImageOnly ? (isPPT ? `➕ Insert Image as New Slide` : `➕ Insert Image`) : (isPPT ? `➕ Insert as New Slide` : (isExcel ? `➕ Insert into Sheet` : `➕ Insert at Cursor`));
  insertBtn.title = shouldInsertImageOnly ? "Insert image into presentation" : (isPPT ? "Create new presentation slides at the end of the deck" : "Insert at current cursor location");
  insertBtn.onclick = async () => {
    if (shouldInsertImageOnly) {
      const img = getTargetVisualImage();
      if (img && img.src) {
        await performDocumentInsertion(`<img src="${img.src}" />`, `![Image](${img.src})`, "insert_cursor", { imageOnly: true, isSummarizeSlides: false });
      }
    } else {
      await performDocumentInsertion(textDiv.innerHTML, fullText, "insert_cursor", { isSummarizeSlides: isBubbleSummarize });
    }
  };

  if (replaceBtn) primaryActions.appendChild(replaceBtn);
  if (insertCurrentBtn) primaryActions.appendChild(insertCurrentBtn);
  primaryActions.appendChild(insertBtn);
  actionsContainer.appendChild(primaryActions);

  // Refinement Chips: Quick 1-Click Multi-Turn Prompts
  const refinementChips = document.createElement("div");
  refinementChips.className = "refinement-chips refinement-chips-container";

  const chips = [
    { label: "✍️ Professional Tone", prompt: "Please rewrite the above in an executive, formal, and highly professional corporate tone." },
    { label: "✂️ Make More Concise", prompt: "Please shorten and condense the above output into a concise version, keeping only the most essential executive points." }
  ];

  chips.forEach(chip => {
    const chipBtn = document.createElement("button");
    chipBtn.type = "button";
    chipBtn.className = "refinement-chip";
    chipBtn.innerText = chip.label;
    chipBtn.title = `Follow-up: "${chip.prompt}"`;
    chipBtn.onclick = async (e) => {
      e.preventDefault();
      e.stopPropagation();
      const runButton = document.getElementById("run");
      if (runButton && runButton.disabled) return;

      const promptInput = document.getElementById("promptText");
      if (promptInput) {
        promptInput.value = "";
      }
      await callGeminiProxy(chip.prompt, true, chip.label);
    };
    refinementChips.appendChild(chipBtn);
  });

  actionsContainer.appendChild(refinementChips);
  bubble.appendChild(actionsContainer);

  historyDiv.appendChild(bubble);
  historyDiv.scrollTop = historyDiv.scrollHeight;

  // POWERPOINT INSTANT SLIDE DECK ENHANCEMENT:
  // Immediately parse slides and display deck outline preview cards if multiple slides are detected
  if (isPPT) {
    try {
      import('../adapters/ppt/slidePreviewUI.js').then(({ enhanceBubbleWithSlideDeck }) => {
        enhanceBubbleWithSlideDeck(bubble, textDiv.innerHTML, fullText, hostAdapter);
      }).catch(err => {
        console.warn("Direct slide deck preview enhancement error:", err);
      });
    } catch (e) {
      console.warn("Direct slide preview invocation failed:", e);
    }
  }

  // Ensure action buttons remain fully visible when images finish decoding
  const bubbleImages = bubble.querySelectorAll("img");
  bubbleImages.forEach(img => {
    if (img.complete) {
      historyDiv.scrollTop = historyDiv.scrollHeight;
    } else {
      img.addEventListener("load", () => {
        historyDiv.scrollTop = historyDiv.scrollHeight;
      });
    }
  });
}

async function performDocumentInsertion(htmlContent, rawText, mode = "smart", options = {}) {
  const runButton = document.getElementById("run");
  const loadingText = document.getElementById("loading");

  if (runButton) runButton.disabled = true;
  if (loadingText) {
    let msg = "⚡ Updating document...";
    if (hostAdapter?.name === 'PowerPoint') {
      if (options?.imageOnly) {
        msg = mode === 'insert_current_slide' ? "⚡ Inserting image onto slide..." : "⚡ Inserting image as slide...";
      } else if (mode === 'replace_draft') {
        msg = "⚡ Replacing slide content...";
      } else if (mode === 'insert_current_slide') {
        msg = "⚡ Inserting onto current slide...";
      } else {
        msg = "⚡ Creating PowerPoint slides...";
      }
    }
    loadingText.innerText = msg;
    loadingText.style.display = "block";
  }

  try {
    const isPPT = hostAdapter?.name === "PowerPoint";
    if (isPPT) {
      await hostAdapter.insertContent(htmlContent, rawText, {
        ...options,
        mode: mode === "replace_draft" ? "replace" : (mode === "insert_current_slide" ? "insert_current" : "insert")
      });
    } else {
      await hostAdapter.insertContent(htmlContent, mode, options);
    }
    const debugStatus = document.getElementById("debugStatus");
    if (debugStatus) {
      const host = hostAdapter?.name || 'Office';
      let actionLabel = mode === 'replace_draft' ? 'Replaced' : (mode === 'insert_current_slide' ? 'Inserted on Slide' : 'Inserted');
      debugStatus.innerText = `Updated in ${host} (${actionLabel})`;
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
        const v110 = Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
        const v19 = Office.context.requirements.isSetSupported("PowerPointApi", "1.9");
        const v18 = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
        const v17 = Office.context.requirements.isSetSupported("PowerPointApi", "1.7");
        const v16 = Office.context.requirements.isSetSupported("PowerPointApi", "1.6");
        const v15 = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
        const v14 = Office.context.requirements.isSetSupported("PowerPointApi", "1.4");
        const v13 = Office.context.requirements.isSetSupported("PowerPointApi", "1.3");
        const v12 = Office.context.requirements.isSetSupported("PowerPointApi", "1.2");
        const v11 = Office.context.requirements.isSetSupported("PowerPointApi", "1.1");
        const highest = v110 ? "1.10" : v19 ? "1.9" : v18 ? "1.8" : v17 ? "1.7" : v16 ? "1.6" : v15 ? "1.5" : v14 ? "1.4" : v13 ? "1.3" : v12 ? "1.2" : v11 ? "1.1" : "Base";
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

