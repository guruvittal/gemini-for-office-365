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
  if (!slides || slides.length === 0) return;

  injectPowerPointStyles();

  const container = document.createElement("div");
  container.className = "ppt-deck-preview-container";

  // Header
  const header = document.createElement("div");
  header.className = "ppt-deck-header";
  header.innerHTML = `
    <div class="ppt-deck-title">📊 <span>Presentation ${slides.length > 1 ? "Deck " : ""}Ready</span></div>
    <div class="ppt-deck-badge">${slides.length} Slide${slides.length > 1 ? "s" : ""}</div>
  `;
  container.appendChild(header);

  // Slides List
  const slidesList = document.createElement("div");
  slidesList.className = "ppt-slides-list";

  slides.forEach((slide, idx) => {
    const card = document.createElement("div");
    card.className = "ppt-slide-card";

    let previewContent = "";
    if (slide.tableData && slide.tableData.rows && slide.tableData.rows.length > 0) {
      const headersHtml = (slide.tableData.headers || []).map(h => `<th style="padding: 4px 6px; background: #0078d4; color: #ffffff; font-weight: 600; text-align: left; font-size: 10.5px; border: 1px solid #c8c6c4;">${escapeHtml(h)}</th>`).join("");
      const rowsHtml = slide.tableData.rows.map((r, rIdx) => {
        const bg = rIdx % 2 === 1 ? '#f3f2f1' : '#ffffff';
        const cells = r.map((c, cIdx) => `<td style="padding: 4px 6px; font-size: 10.5px; border: 1px solid #edebe9; ${cIdx === 0 ? 'font-weight: 600;' : ''}">${escapeHtml(c)}</td>`).join("");
        return `<tr style="background: ${bg};">${cells}</tr>`;
      }).join("");

      previewContent = `
        <div style="overflow-x: auto; margin: 4px 0;">
          <table style="width: 100%; border-collapse: collapse; border: 1px solid #c8c6c4; font-size: 10.5px; line-height: 1.3;">
            ${headersHtml ? `<thead><tr>${headersHtml}</tr></thead>` : ''}
            <tbody>${rowsHtml}</tbody>
          </table>
        </div>
      `;
    } else {
      const previewBody = slide.body
        ? slide.body.trim()
        : "Full slide content & visual layout prepared.";
      previewContent = `<div style="white-space: pre-wrap;">${escapeHtml(previewBody)}</div>`;
    }

    card.innerHTML = `
      <div class="ppt-slide-card-header">
        <span class="ppt-slide-num">Slide ${idx + 1}</span>
        <span>${escapeHtml(slide.title)}</span>
      </div>
      <div class="ppt-slide-preview-body">${previewContent}</div>
    `;
    slidesList.appendChild(card);
  });
  container.appendChild(slidesList);

  // Footer instruction
  const footerHint = document.createElement("div");
  footerHint.className = "ppt-deck-footer-hint";
  footerHint.innerHTML = `👉 Click <b>"${slides.length > 1 ? '+ Insert ' + slides.length + ' Slides' : '+ Insert into Slide'}"</b> below to create the ${slides.length > 1 ? 'slides' : 'slide'}.`;
  container.appendChild(footerHint);

  // Optional collapsible raw text outline
  const rawDetails = document.createElement("details");
  rawDetails.className = "ppt-raw-details";
  rawDetails.style.cssText = "margin-top: 6px; font-size: 10.5px; color: #605e5c;";
  rawDetails.innerHTML = `
    <summary style="cursor: pointer; color: #0078d4; user-select: none; font-size: 10.5px; font-weight: 500;">📄 View Raw Markdown Text</summary>
    <div style="padding: 6px 8px; background: #faf9f8; border: 1px solid #edebe9; border-radius: 4px; margin-top: 4px; font-size: 11px; line-height: 1.4; max-height: 150px; overflow-y: auto;">
      ${htmlContent}
    </div>
  `;
  container.appendChild(rawDetails);

  // Update existing action buttons in the bubble & hide raw duplicate text
  const actionsContainer = bubbleEl.querySelector(".response-actions-container");
  if (actionsContainer) {
    const insertBtn = actionsContainer.querySelector(".insert-btn") || actionsContainer.querySelector(".action-btn.insert");
    if (insertBtn) {
      insertBtn.innerHTML = slides.length > 1 ? `➕ Insert ${slides.length} Slides` : `➕ Insert into Slide`;
    }
    const replaceBtn = actionsContainer.querySelector(".replace-btn") || actionsContainer.querySelector(".action-btn.replace");
    if (replaceBtn) {
      replaceBtn.innerHTML = slides.length > 1 ? `🔄 Replace with ${slides.length} Slides` : `🔄 Replace Current Slide`;
    }

    // Hide all raw markdown text siblings to eliminate duplicate visual text
    Array.from(bubbleEl.children).forEach(child => {
      if (child !== actionsContainer && child !== container) {
        child.style.display = "none";
      }
    });

    bubbleEl.insertBefore(container, actionsContainer);
  } else {
    Array.from(bubbleEl.children).forEach(child => {
      if (child !== container) child.style.display = "none";
    });
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
