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
    .ppt-deck-preview-container {
      margin: 10px 0 8px 0;
      background: #ffffff;
      border: 1px solid #c7e0f4;
      border-left: 4px solid #0078d4;
      border-radius: 6px;
      padding: 10px;
      box-shadow: 0 2px 8px rgba(0, 120, 212, 0.08);
    }
    .ppt-deck-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
      padding-bottom: 6px;
      border-bottom: 1px solid #eff6fc;
    }
    .ppt-deck-title {
      font-size: 12px;
      font-weight: 700;
      color: #004e8c;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .ppt-deck-badge {
      font-size: 10.5px;
      background: #deecf9;
      color: #0078d4;
      padding: 2px 8px;
      border-radius: 10px;
      font-weight: 600;
    }
    .ppt-slides-list {
      display: flex;
      flex-direction: column;
      gap: 6px;
      max-height: 220px;
      overflow-y: auto;
      margin-bottom: 6px;
      padding-right: 2px;
    }
    .ppt-slide-card {
      background: #faf9f8;
      border: 1px solid #e1dfdd;
      border-radius: 4px;
      padding: 6px 8px;
      font-size: 11.5px;
      transition: all 0.15s ease;
    }
    .ppt-slide-card:hover {
      border-color: #0078d4;
      background: #f3f9fd;
    }
    .ppt-slide-card-header {
      font-weight: 600;
      color: #106ebe;
      display: flex;
      align-items: baseline;
      gap: 6px;
      margin-bottom: 3px;
    }
    .ppt-slide-num {
      font-size: 10px;
      background: #0078d4;
      color: #ffffff;
      padding: 1px 5px;
      border-radius: 3px;
      font-weight: 700;
    }
    .ppt-slide-preview-body {
      color: #605e5c;
      font-size: 11px;
      line-height: 1.4;
      white-space: pre-line;
      padding-top: 4px;
    }
    .ppt-deck-footer-hint {
      font-size: 11px;
      color: #005a9e;
      margin-top: 6px;
      text-align: center;
      font-weight: 500;
      background: #f3f9fd;
      padding: 4px 6px;
      border-radius: 4px;
    }
  `;
  document.head.appendChild(styleEl);
}

/**
 * Attaches the Slide Deck Outline Preview to an assistant response bubble.
 */
export function enhanceBubbleWithSlideDeck(bubbleEl, htmlContent, rawText, adapter) {
  if (!bubbleEl || bubbleEl.querySelector(".ppt-deck-preview-container")) return;

  const slides = parseSlides(htmlContent, rawText);
  if (slides.length < 2) return;

  injectPowerPointStyles();

  const container = document.createElement("div");
  container.className = "ppt-deck-preview-container";

  // Header
  const header = document.createElement("div");
  header.className = "ppt-deck-header";
  header.innerHTML = `
    <div class="ppt-deck-title">📊 <span>Presentation Deck Ready</span></div>
    <div class="ppt-deck-badge">${slides.length} Slides</div>
  `;
  container.appendChild(header);

  // Slides List
  const slidesList = document.createElement("div");
  slidesList.className = "ppt-slides-list";

  slides.forEach((slide, idx) => {
    const card = document.createElement("div");
    card.className = "ppt-slide-card";

    const previewBody = slide.body
      ? slide.body.trim()
      : "Full slide content & visual layout prepared.";

    card.innerHTML = `
      <div class="ppt-slide-card-header">
        <span class="ppt-slide-num">Slide ${idx + 1}</span>
        <span>${escapeHtml(slide.title)}</span>
      </div>
      <div class="ppt-slide-preview-body">${escapeHtml(previewBody)}</div>
    `;
    slidesList.appendChild(card);
  });
  container.appendChild(slidesList);

  // Footer instruction
  const footerHint = document.createElement("div");
  footerHint.className = "ppt-deck-footer-hint";
  footerHint.innerHTML = `👉 Click <b>"+ Insert into Slides"</b> below to create all ${slides.length} slides.`;
  container.appendChild(footerHint);

  // Update existing action buttons in the bubble
  const actionsContainer = bubbleEl.querySelector(".response-actions-container");
  if (actionsContainer) {
    const insertBtn = actionsContainer.querySelector(".insert-btn");
    if (insertBtn) {
      insertBtn.innerHTML = `+ Insert ${slides.length} Slides`;
    }
    const replaceBtn = actionsContainer.querySelector(".replace-btn");
    if (replaceBtn) {
      replaceBtn.innerHTML = `Replace with ${slides.length} Slides`;
    }
    bubbleEl.insertBefore(container, actionsContainer);
  } else {
    bubbleEl.appendChild(container);
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

function escapeHtml(str) {
  if (!str) return "";
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
