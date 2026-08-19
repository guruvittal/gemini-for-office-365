/**
 * PowerPoint Diagnostics & Live Test Console
 * 
 * Provides an open, real-time diagnostic suite inside the Taskpane:
 * 1. Inspects Office.js PowerPoint API requirements and platform details.
 * 2. Streams real-time slide creation and shape drawing logs to the UI.
 * 3. Provides standalone test buttons for slide creation and textbox verification.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

let diagInjected = false;

export function logToPPTConsole(msg, isError = false) {
  const logBox = document.getElementById("pptDebugLog");
  if (!logBox) return;
  const ts = new Date().toLocaleTimeString();
  const prefix = isError ? "❌" : "ℹ️";
  logBox.textContent += `\n[${ts}] ${prefix} ${msg}`;
  logBox.scrollTop = logBox.scrollHeight;
}

export function initPowerPointDiagnostics() {
  if (diagInjected || typeof document === "undefined") return;
  diagInjected = true;

  const chatHistory = document.getElementById("chatHistory");
  if (!chatHistory) return;

  const panel = document.createElement("div");
  panel.id = "pptDebugPanel";
  panel.style.cssText = `
    margin: 8px 10px 12px 10px;
    border: 1px solid #0078d4;
    border-radius: 6px;
    background: #fdfdfd;
    overflow: hidden;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    font-size: 11.5px;
    box-shadow: 0 2px 6px rgba(0, 120, 212, 0.12);
  `;

  panel.innerHTML = `
    <div id="pptDebugToggle" style="
      background: #0078d4;
      padding: 8px 10px;
      font-weight: 600;
      color: #ffffff;
      cursor: pointer;
      display: flex;
      justify-content: space-between;
      align-items: center;
    ">
      <span>🛠️ <strong>PowerPoint Diagnostic & Live Log Console</strong></span>
      <span id="pptDebugChevron" style="font-size: 11px; color: #ffffff;">▲ Open</span>
    </div>
    <div id="pptDebugContent" style="display: block; padding: 10px; background: #faf9f8;">
      <div style="display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 8px;">
        <button id="btnDiagCheckApi" style="padding: 5px 9px; font-size: 11px; font-weight: 600; background: #0078d4; color: #fff; border: none; border-radius: 4px; cursor: pointer;">🔍 Check API Support</button>
        <button id="btnDiagTestSlide" style="padding: 5px 9px; font-size: 11px; font-weight: 600; background: #107c41; color: #fff; border: none; border-radius: 4px; cursor: pointer;">🧪 Test: Add 1 Slide</button>
        <button id="btnDiagTestText" style="padding: 5px 9px; font-size: 11px; font-weight: 600; background: #5c2d91; color: #fff; border: none; border-radius: 4px; cursor: pointer;">📝 Test: Slide + Textbox</button>
        <button id="btnDiagClearLog" style="padding: 5px 9px; font-size: 11px; background: #edebe9; color: #323130; border: 1px solid #c8c6c4; border-radius: 4px; cursor: pointer;">🧹 Clear</button>
      </div>
      <div id="pptDebugLog" style="
        background: #1e1e1e;
        color: #00ff66;
        font-family: Monaco, Menlo, Consolas, monospace;
        font-size: 11px;
        padding: 8px;
        border-radius: 4px;
        height: 150px;
        overflow-y: auto;
        white-space: pre-wrap;
        line-height: 1.4;
      ">PowerPoint Diagnostic Console Initialized.\nReady to monitor slide generation.</div>
    </div>
  `;

  // Insert above chat history
  const parent = chatHistory.parentElement || document.body;
  parent.insertBefore(panel, chatHistory);

  // Toggle behavior
  const toggle = document.getElementById("pptDebugToggle");
  const content = document.getElementById("pptDebugContent");
  const chevron = document.getElementById("pptDebugChevron");

  toggle.addEventListener("click", () => {
    const isClosed = content.style.display === "none";
    content.style.display = isClosed ? "block" : "none";
    chevron.textContent = isClosed ? "▲ Open" : "▼ Open";
  });

  // Button 1: Check API Support
  document.getElementById("btnDiagCheckApi").addEventListener("click", async () => {
    logToPPTConsole("Checking PowerPoint Office.js environment...");
    try {
      const hasOffice = typeof Office !== "undefined";
      const hasPPT = typeof PowerPoint !== "undefined";
      logToPPTConsole(`Office.js Loaded: ${hasOffice}`);
      logToPPTConsole(`PowerPoint Namespace: ${hasPPT}`);

      if (hasOffice && Office.context) {
        logToPPTConsole(`Host: ${Office.context.host || "Unknown"}`);
        logToPPTConsole(`Platform: ${Office.context.platform || "Unknown"}`);

        if (Office.context.requirements) {
          const v11 = Office.context.requirements.isSetSupported("PowerPointApi", "1.1");
          const v12 = Office.context.requirements.isSetSupported("PowerPointApi", "1.2");
          const v13 = Office.context.requirements.isSetSupported("PowerPointApi", "1.3");
          const v14 = Office.context.requirements.isSetSupported("PowerPointApi", "1.4");
          const v15 = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
          logToPPTConsole(`PowerPointApi Support: 1.1=${v11}, 1.2=${v12}, 1.3=${v13}, 1.4=${v14}, 1.5=${v15}`);
        }
      }
    } catch (e) {
      logToPPTConsole(`API Check Error: ${e.message}`, true);
    }
  });

  // Button 2: Test Add 1 Slide
  document.getElementById("btnDiagTestSlide").addEventListener("click", async () => {
    logToPPTConsole("Calling PowerPoint.run(presentation.slides.add())...");
    try {
      await PowerPoint.run(async (context) => {
        context.presentation.slides.add();
        logToPPTConsole("Slide add queued. Calling context.sync()...");
        await context.sync();
      });
      logToPPTConsole("✅ Successfully added 1 blank slide via PowerPoint.run!");
    } catch (e) {
      logToPPTConsole(`Slide Add Error: ${e.message} (code: ${e.code || "N/A"})`, true);
    }
  });

  // Button 3: Test Add Slide with Textbox using getCount()
  document.getElementById("btnDiagTestText").addEventListener("click", async () => {
    logToPPTConsole("Testing Add Slide + Title Textbox + Body Textbox via getCount()...");
    try {
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        
        logToPPTConsole("1. Calling slides.add()...");
        slides.add();
        await context.sync();

        logToPPTConsole("2. Calling slides.getCount()...");
        const countResult = slides.getCount();
        await context.sync();

        const slideCount = countResult.value;
        logToPPTConsole(`3. Slide count: ${slideCount}. Fetching slide at index ${slideCount - 1}...`);
        const slide = slides.getItemAt(slideCount - 1);

        logToPPTConsole("4. Adding title textbox...");
        slide.shapes.addTextBox("🧪 Diagnostic Test Title", {
          left: 50,
          top: 40,
          width: 650,
          height: 55
        });

        logToPPTConsole("5. Adding body textbox...");
        slide.shapes.addTextBox("• Bullet item 1: Official Microsoft Office.js getCount() pattern\n• Bullet item 2: Direct geometry textbox creation verified", {
          left: 50,
          top: 110,
          width: 650,
          height: 300
        });

        logToPPTConsole("6. Finalizing sync...");
        await context.sync();
      });
      logToPPTConsole("✅ Successfully created Slide with Title & Body Textbox!");
    } catch (e) {
      logToPPTConsole(`Slide Text Error: ${e.message} (code: ${e.code || "N/A"})`, true);
    }
  });

  // Button 4: Clear Logs
  document.getElementById("btnDiagClearLog").addEventListener("click", () => {
    const box = document.getElementById("pptDebugLog");
    if (box) box.textContent = "Logs cleared.";
  });
}
