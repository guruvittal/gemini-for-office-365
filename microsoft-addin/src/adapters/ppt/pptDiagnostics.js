/**
 * PowerPoint Diagnostics & Live Logger
 * 
 * Routes PowerPoint slide generation logs and test calls directly
 * into the unified Troubleshooting & Diagnostics panel.
 */

export function logToPPTConsole(msg, isError = false) {
  const logBox = document.getElementById("diagLogBox");
  if (!logBox) {
    console.log(`[PPT] ${msg}`);
    return;
  }
  const ts = new Date().toLocaleTimeString();
  const prefix = isError ? "❌ [PPT]" : "ℹ️ [PPT]";
  logBox.textContent += `\n[${ts}] ${prefix} ${msg}`;
  logBox.scrollTop = logBox.scrollHeight;
}

export function initPowerPointDiagnostics() {
  // Integrated into unified taskpane troubleshooting panel
}
