/**
 * Host Adapter Implementation for Microsoft PowerPoint (PowerPoint.run + Office Common API)
 * 
 * Provides complete multi-slide presentation generation, intelligent outline/table
 * unpacking, visual slide deck preview, executive PowerPoint styling, and in-taskpane diagnostics.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { parseSlides, extractCleanBulletPoints } from './ppt/slideParser.js';
import { buildPresentation, compressImageForPowerPoint, insertOnCurrentSlide } from './ppt/slideBuilder.js';
import { initSlidePreviewObserver, injectPowerPointStyles } from './ppt/slidePreviewUI.js';
import { initPromptEnhancer, enhancePromptForPowerPoint } from './ppt/promptEnhancer.js';
import { initPowerPointDiagnostics } from './ppt/pptDiagnostics.js';

export class PPTAdapter {
  constructor() {
    this.name = "PowerPoint";
    
    try {
      injectPowerPointStyles();
      initSlidePreviewObserver(this);
      initPromptEnhancer();
      initPowerPointDiagnostics();
    } catch (e) {
      console.warn("PPTAdapter initialization warning:", e);
    }
  }

  /**
   * Safely extracts all text and table contents from a PowerPoint slide object.
   * Completely resilient against non-text shapes, groups, and API version differences.
   */
  async extractSlideText(slide, slideNum, context) {
    const slideLines = [];
    try {
      const shapes = slide.shapes;
      shapes.load("items/id,items/type,items/name");
      await context.sync();

      if (!shapes.items || shapes.items.length === 0) {
        return { slideNumber: slideNum, id: slide.id || `slide-${slideNum}`, text: "" };
      }

      // Track text ranges and tables safely based on verified shape type
      const textItemTrackers = [];
      const tableItemTrackers = [];

      for (const shape of shapes.items) {
        try {
          const type = shape.type;
          // Check for Table shape (starting in PowerPointApi 1.8)
          if (type === "Table" || (typeof PowerPoint !== "undefined" && PowerPoint.ShapeType && type === PowerPoint.ShapeType.table)) {
            if (typeof shape.getTable === "function") {
              const table = shape.getTable();
              table.load("values");
              tableItemTrackers.push(table);
            }
          } else if (
            type === "TextBox" || 
            type === "GeometricShape" || 
            type === "Placeholder" || 
            type === "Callout" ||
            !type // Fallback if type property was not populated
          ) {
            // Text-bearing shapes
            if (shape.textFrame) {
              const tr = shape.textFrame.textRange;
              tr.load("text");
              textItemTrackers.push(tr);
            }
          }
        } catch (shapeErr) {
          console.warn(`[PPTAdapter] Shape pre-check skipped on slide ${slideNum}:`, shapeErr);
        }
      }

      // Try batch sync for efficiency
      let batchSuccess = false;
      if (textItemTrackers.length > 0 || tableItemTrackers.length > 0) {
        try {
          await context.sync();
          batchSuccess = true;
        } catch (batchErr) {
          console.warn(`[PPTAdapter] Batch shape sync failed on slide ${slideNum}, falling back to per-shape sync:`, batchErr);
        }
      }

      if (batchSuccess) {
        for (const tr of textItemTrackers) {
          try {
            const val = (tr.text || "").trim();
            if (val) slideLines.push(val);
          } catch (_) {}
        }
        for (const tb of tableItemTrackers) {
          try {
            if (tb.values && Array.isArray(tb.values)) {
              for (const row of tb.values) {
                if (Array.isArray(row)) {
                  const rowStr = row.map(c => String(c ?? "").trim()).join(" | ");
                  if (rowStr.replace(/[|\s]/g, "")) {
                    slideLines.push(`| ${rowStr} |`);
                  }
                }
              }
            }
          } catch (_) {}
        }
      } else {
        // Fallback: Per-shape isolated extraction so a problematic shape never blocks other shapes
        for (const shape of shapes.items) {
          try {
            if (shape.type === "Table" && typeof shape.getTable === "function") {
              const table = shape.getTable();
              table.load("values");
              await context.sync();
              if (table.values && Array.isArray(table.values)) {
                for (const row of table.values) {
                  if (Array.isArray(row)) {
                    const rowStr = row.map(c => String(c ?? "").trim()).join(" | ");
                    if (rowStr.replace(/[|\s]/g, "")) slideLines.push(`| ${rowStr} |`);
                  }
                }
              }
            } else if (shape.textFrame) {
              const tr = shape.textFrame.textRange;
              tr.load("text");
              await context.sync();
              const val = (tr.text || "").trim();
              if (val) slideLines.push(val);
            }
          } catch (innerErr) {
            // Silently swallow per-shape errors so other shapes succeed
          }
        }
      }
    } catch (slideErr) {
      console.warn(`[PPTAdapter] Error extracting text from slide ${slideNum}:`, slideErr);
    }

    const slideText = slideLines.join("\n").trim();
    console.log(`📊 [PPTAdapter] Extracted ${slideLines.length} text/table blocks from Slide ${slideNum} (${slideText.length} chars)`);
    return {
      slideNumber: slideNum,
      id: slide.id || `slide-${slideNum}`,
      text: slideText
    };
  }

  // On-demand selection extraction (strictly user-triggered on button click, NO background event listeners)
  async getSelectedSlidesText() {
    const selectedSlidesData = [];
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          // Map all slide IDs in presentation to determine real 1-based slide numbers
          const allSlides = context.presentation.slides;
          allSlides.load("items/id");

          let selectedSlides = null;
          if (context.presentation.getSelectedSlides) {
            selectedSlides = context.presentation.getSelectedSlides();
            selectedSlides.load("items/id");
          }
          await context.sync();

          const slideIndexMap = {};
          if (allSlides.items) {
            allSlides.items.forEach((s, idx) => {
              if (s.id) slideIndexMap[s.id] = idx + 1;
            });
          }

          if (selectedSlides && selectedSlides.items && selectedSlides.items.length > 0) {
            for (let i = 0; i < selectedSlides.items.length; i++) {
              try {
                const slide = selectedSlides.items[i];
                const slideNum = (slide.id && slideIndexMap[slide.id]) ? slideIndexMap[slide.id] : (i + 1);
                const slideData = await this.extractSlideText(slide, slideNum, context);
                selectedSlidesData.push(slideData);
              } catch (slideLoopErr) {
                console.warn(`[PPTAdapter] Error in selected slide loop item ${i}:`, slideLoopErr);
              }
            }
          }
        });
      }
    } catch (err) {
      console.warn("PowerPoint getSelectedSlidesText warning:", err);
    }
    return selectedSlidesData;
  }

  // Read currently highlighted shape or text frame on the active slide on demand
  async getSelectedShapeText() {
    let selectedText = "";
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          if (context.presentation.getSelectedShapes) {
            const selection = context.presentation.getSelectedShapes();
            selection.load("items");
            await context.sync();

            if (selection.items && selection.items.length > 0) {
              const textTrackers = [];
              for (const shape of selection.items) {
                try {
                  if (shape.textFrame) {
                    const tr = shape.textFrame.textRange;
                    tr.load("text");
                    textTrackers.push(tr);
                  }
                } catch (_) {}
              }
              if (textTrackers.length > 0) {
                await context.sync();
                const texts = textTrackers
                  .map(tr => (tr.text || "").trim())
                  .filter(Boolean);
                selectedText = texts.join("\n\n");
              }
            }
          }
        });
      }
    } catch (err) {
      console.warn("PowerPoint getSelectedShapeText error:", err);
    }
    return selectedText;
  }

  // Read currently highlighted text or selected text shape on demand
  async getSelectedText() {
    // 1. Try Office Common getSelectedDataAsync first (fastest for user-highlighted text in any text box)
    try {
      if (typeof Office !== 'undefined' && Office.context?.document?.getSelectedDataAsync) {
        const textFromCommonApi = await new Promise((resolve) => {
          Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            (result) => {
              if (result && result.status === Office.AsyncResultStatus.Succeeded && typeof result.value === "string") {
                resolve(result.value.trim());
              } else {
                resolve("");
              }
            }
          );
        });
        if (textFromCommonApi && textFromCommonApi.length > 0) {
          return textFromCommonApi;
        }
      }
    } catch (_) {}

    // 2. Try selected shape(s) text (when user clicked or selected a text box / shape)
    const shapeText = await this.getSelectedShapeText();
    if (shapeText && shapeText.trim().length > 0) {
      return shapeText.trim();
    }

    // Return empty string if no specific text or text shape is selected
    // (Slide text extraction is handled on demand by getSelectedSlidesText)
    return "";
  }

  // Read full presentation text across all slides and shapes safely using index-based traversal
  async getFullDocumentText() {
    let fullText = "";
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items/id");
          const countResult = slides.getCount();
          await context.sync();

          const total = countResult.value || (slides.items ? slides.items.length : 0);
          const slideTexts = [];

          for (let i = 0; i < total; i++) {
            try {
              const slide = slides.getItemAt(i);
              const slideData = await this.extractSlideText(slide, i + 1, context);
              if (slideData.text && slideData.text.length > 0) {
                slideTexts.push(`[Slide ${i + 1}]:\n${slideData.text}`);
              }
            } catch (slideErr) {
              console.warn(`[PPTAdapter] Error in getFullDocumentText slide ${i + 1}:`, slideErr);
            }
          }

          fullText = slideTexts.join("\n\n---\n\n");
        });
      }
    } catch (e) {
      console.warn("PPT full text read error:", e);
    }
    return fullText.trim();
  }

  // Parse HTML or Markdown content into executive slide structures
  async parseSlidesFromHtml(htmlContent, rawText = "") {
    return parseSlides(htmlContent, rawText);
  }

  // Insert AI content as executive PowerPoint slides with exact positioning & visuals
  async insertContent(htmlContent, rawText = "", options = {}) {
    const debugStatus = document.getElementById("debugStatus");
    const loadingText = document.getElementById("loading");
    const isReplace = options.mode === "replace" || options.mode === "replace_draft";

    try {
      window.__isGeneratingSlides = true;
      if (typeof PowerPoint === 'undefined') {
        throw new Error("PowerPoint Office.js environment is not available.");
      }

      // 1. If replacing and an active shape/text is selected, perform in-place text replacement in the shape
      if (isReplace) {
        let shapeReplaced = false;
        try {
          await PowerPoint.run(async (context) => {
            if (context.presentation.getSelectedShapes) {
              const selection = context.presentation.getSelectedShapes();
              selection.load("items/textFrame");
              await context.sync();
              if (selection.items && selection.items.length > 0) {
                const shape = selection.items[0];
                if (shape.textFrame) {
                  const cleanBullets = extractCleanBulletPoints(htmlContent, rawText);
                  shape.textFrame.textRange.text = cleanBullets;
                  try {
                    shape.textFrame.textRange.font.size = 14;
                  } catch (_) {}
                  await context.sync();

                  // Bold lead-in phrases before colons
                  try {
                    const lines = cleanBullets.split("\n");
                    let charOffset = 0;
                    for (const l of lines) {
                      const colonIdx = l.indexOf(":");
                      const dashIdx = l.indexOf("—");
                      const sepIdx = colonIdx > 0 ? colonIdx : (dashIdx > 0 ? dashIdx : -1);
                      if (sepIdx > 0 && sepIdx < 45 && typeof shape.textFrame.textRange.getSubstring === "function") {
                        try {
                          const leadIn = shape.textFrame.textRange.getSubstring(charOffset, sepIdx + 1);
                          leadIn.font.bold = true;
                        } catch (_) {}
                      }
                      charOffset += l.length + 1;
                    }
                  } catch (_) {}

                  shapeReplaced = true;
                }
              }
            }
          });
        } catch (shapeErr) {
          console.warn("Shape replacement check:", shapeErr);
        }

        if (shapeReplaced) {
          if (debugStatus) debugStatus.innerText = "✅ Replaced text in slide!";
          if (loadingText) loadingText.style.display = "none";
          return [{ title: "Updated Shape", body: rawText }];
        }
      }

      // 1b. If inserting on current slide
      const isInsertCurrent = options.mode === "insert_current" || options.mode === "insert_current_slide";
      if (isInsertCurrent) {
        if (debugStatus) debugStatus.innerText = "Inserting content onto current slide...";
        if (loadingText) {
          loadingText.innerText = "⚡ Inserting onto current slide...";
          loadingText.style.display = "block";
        }
        const slideStructures = await this.parseSlidesFromHtml(htmlContent, rawText);
        if (!slideStructures || slideStructures.length === 0) {
          throw new Error("Slide parser returned 0 slide structures.");
        }
        await insertOnCurrentSlide(slideStructures, options);
        if (debugStatus) debugStatus.innerText = "✅ Inserted on current slide!";
        if (loadingText) loadingText.style.display = "none";
        return slideStructures;
      }

      // 2. Otherwise parse slide structures and build / replace slide(s)
      if (debugStatus) debugStatus.innerText = "Parsing presentation structure...";
      const slideStructures = await this.parseSlidesFromHtml(htmlContent, rawText);

      if (!slideStructures || slideStructures.length === 0) {
        throw new Error("Slide parser returned 0 slide structures.");
      }

      const onProgress = (prog) => {
        const msg = `⚡ ${isReplace ? 'Replacing' : 'Creating'} slide ${prog.current}/${prog.total}: "${prog.title}"...`;
        if (debugStatus) debugStatus.innerText = msg;
        if (loadingText) {
          loadingText.innerText = msg;
          loadingText.style.display = "block";
        }
        if (typeof options.onProgress === "function") {
          options.onProgress(prog);
        }
      };

      await buildPresentation(slideStructures, options, onProgress);

      if (debugStatus) {
        debugStatus.innerText = `✅ ${isReplace ? 'Replaced' : 'Created'} ${slideStructures.length} slide(s) in PowerPoint!`;
      }
      if (loadingText) {
        loadingText.style.display = "none";
      }

      return slideStructures;
    } catch (err) {
      console.error("PPTAdapter Exception:", err);
      const errDetail = err.message || JSON.stringify(err);
      if (debugStatus) {
        debugStatus.innerHTML = `<span style="color:#d13438; font-weight:bold;">🔴 PPT Error: ${errDetail}</span>`;
      }
      if (loadingText) {
        loadingText.innerText = `🔴 Error: ${errDetail}`;
      }
      throw err;
    } finally {
      window.__isGeneratingSlides = false;
    }
  }

  // Scan in-slide / in-shape @gemini commands for PowerPoint
  async checkInDocumentCommands(forceRun = false, callbacks = {}) {
    if (callbacks.onStatus) callbacks.onStatus("PowerPoint Adapter Ready (@gemini in shape)");
  }
}
