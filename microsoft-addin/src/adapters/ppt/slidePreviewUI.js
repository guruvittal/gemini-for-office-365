/**
 * Slide Preview UI & Interactive Deck Handler for PowerPoint
 * 
 * Automatically attaches to the Add-in Taskpane in PowerPoint, detects multi-slide
 * content in Gemini responses, and renders an executive Slide Deck Outline Preview.
 * Enhances the native "+ Insert into Slides" toolbar button to generate the complete deck.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { parseSlides } from './slideParser.js';

let stylesInjected = false;

/**
 * Injects scoped styling for the PowerPoint Slide Deck Preview into the DOM.
 */
export function injectPowerPointStyles() {
  if (stylesInjected || typeof document === "undefined") return;
  stylesInjected = true;

  const styleEl = document.createElement("style");
  styleEl.id = "ppt-preview-styles";
  styleEl.textContent = `
    .ppt-deck-pill {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: #eff6fc;
      border: 1px solid #c7e0f4;
      border-left: 3px solid #0078d4;
      border-radius: 4px;
      padding: 3px 8px;
      font-size: 11px;
      color: #004e8c;
      font-weight: 600;
      margin-bottom: 8px;
    }
    .ppt-outline-details {
      margin-top: 8px;
      margin-bottom: 6px;
      font-size: 11px;
    }
    .ppt-outline-summary {
      cursor: pointer;
      color: #0078d4;
      font-weight: 500;
      user-select: none;
    }
    .ppt-outline-summary:hover {
      text-decoration: underline;
    }
  `;
  document.head.appendChild(styleEl);
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * Attaches the Slide Deck Badge & Outline to an assistant response bubble.
 * Crucially: NEVER hides the rendered response text.
 */
export function enhanceBubbleWithSlideDeck(bubbleEl, htmlContent, rawText, adapter) {
  if (!bubbleEl || bubbleEl.querySelector(".ppt-deck-pill")) return;

  const slides = parseSlides(htmlContent, rawText);
  if (!slides || slides.length === 0) return;

  injectPowerPointStyles();

  // 1. Add sleek badge at the top of the response bubble
  const pill = document.createElement("div");
  pill.className = "ppt-deck-pill";
  pill.innerHTML = `📊 <span>${slides.length > 1 ? slides.length + ' Slides' : 'Slide'} Ready to Insert</span>`;
  bubbleEl.insertBefore(pill, bubbleEl.firstChild);

  // 2. Update existing action buttons in the bubble
  const actionsContainer = bubbleEl.querySelector(".response-actions-container");
  if (actionsContainer) {
    const insertBtn = actionsContainer.querySelector(".insert-btn") || actionsContainer.querySelector(".action-btn.insert");
    if (insertBtn) {
      insertBtn.innerHTML = slides.length > 1 ? `➕ Insert ${slides.length} Slides` : `➕ Insert into Slide`;
      insertBtn.style.backgroundColor = "#107c41";
    }
    const replaceBtn = actionsContainer.querySelector(".replace-btn") || actionsContainer.querySelector(".action-btn.replace");
    if (replaceBtn) {
      replaceBtn.innerHTML = slides.length > 1 ? `🔄 Replace with ${slides.length} Slides` : `🔄 Replace Current Slide`;
    }

    // 3. If multi-slide deck, add an optional collapsed outline preview (does not take visual space by default)
    if (slides.length > 1) {
      const outlineDetails = document.createElement("details");
      outlineDetails.className = "ppt-outline-details";
      const summary = document.createElement("summary");
      summary.className = "ppt-outline-summary";
      summary.innerText = `📑 Preview Deck Outline (${slides.length} slides)`;
      outlineDetails.appendChild(summary);

      const outlineList = document.createElement("div");
      outlineList.style.cssText = "padding: 6px 8px; margin-top: 4px; background: #faf9f8; border: 1px solid #edebe9; border-radius: 4px; font-size: 11px; line-height: 1.4;";
      slides.forEach((s, idx) => {
        const item = document.createElement("div");
        item.style.marginBottom = "3px";
        item.innerHTML = `<b>Slide ${idx + 1}:</b> ${escapeHtml(s.title || "Untitled")}`;
        outlineList.appendChild(item);
      });
      outlineDetails.appendChild(outlineList);
      bubbleEl.insertBefore(outlineDetails, actionsContainer);
    }
  }
}

/**
 * Initializes a MutationObserver on #chatHistory to automatically enhance
 * all incoming assistant responses with the Slide Deck Preview.
 */
export function initSlidePreviewObserver(adapter) {
  if (typeof document === "undefined") return;

  const chatHistory = document.getElementById("chatHistory");
  if (!chatHistory) return;

  injectPowerPointStyles();

  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === 1 && node.classList.contains("assistant")) {
          setTimeout(() => {
            const rawText = node.innerText || "";
            const htmlContent = node.innerHTML || "";
            enhanceBubbleWithSlideDeck(node, htmlContent, rawText, adapter);
          }, 50);
        }
      });
    });
  });

  observer.observe(chatHistory, { childList: true, subtree: true });
}

