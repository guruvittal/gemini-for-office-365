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

  // Safely extract all text and table contents from a single PowerPoint slide by its ID in an isolated session
  async _extractSlideTextById(slideId) {
    let resultText = "";
    try {
      await PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItem(slideId);
        const shapes = slide.shapes;
        shapes.load("items/name, items/type");
        await context.sync();

        if (!shapes.items || shapes.items.length === 0) return;

        const textTrackers = [];
        const tableTrackers = [];

        for (const shape of shapes.items) {
          const typeStr = (shape.type || "").toString().toLowerCase();

          // 1. Table shape extraction
          if (typeStr.includes("table") && typeof shape.getTable === "function") {
            try {
              const table = shape.getTable();
              table.load("values");
              tableTrackers.push(table);
            } catch (_) {}
          } else {
            // 2. Text shape extraction (TextBox, GeometricShape, Callout, etc.)
            try {
              if (shape.textFrame) {
                const tr = shape.textFrame.textRange;
                tr.load("text");
                textTrackers.push(tr);
              }
            } catch (_) {}
          }
        }

        await context.sync();

        const textLines = [];
        for (const tr of textTrackers) {
          try {
            if (tr.text && tr.text.trim()) {
              textLines.push(tr.text.trim());
            }
          } catch (_) {}
        }

        for (const tbl of tableTrackers) {
          try {
            if (tbl.values && Array.isArray(tbl.values)) {
              const tableRows = tbl.values
                .map(row => (Array.isArray(row) ? row.join(" | ") : String(row)))
                .filter(line => line.trim().length > 0);
              if (tableRows.length > 0) {
                textLines.push(tableRows.join("\n"));
              }
            }
          } catch (_) {}
        }

        resultText = textLines.join("\n\n");
      });
    } catch (err) {
      console.warn(`[PPTAdapter] Error extracting text from slide ${slideId}:`, err);
    }
    return resultText;
  }

  // All automatic highlighting and selection functionality is completely disabled
  // per user request to isolate and prevent PowerPoint Online selection jitters, slide jumps, and DLP popups.
  async getSelectedSlidesText() {
    return [];
  }

  // Read currently highlighted shape or text frame on the active slide (disabled)
  async getSelectedShapeText() {
    return "";
  }

  // Read currently highlighted slide(s) text or active shape selection (disabled)
  async getSelectedText() {
    return "";
  }

  // Read full presentation text across all slides and shapes
  async getFullDocumentText() {
    const fullTextParts = [];
    try {
      if (typeof PowerPoint !== 'undefined') {
        const slideIds = [];
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items/id");
          await context.sync();

          if (slides.items && slides.items.length > 0) {
            for (const s of slides.items) {
              if (s.id) slideIds.push(s.id);
            }
          }
        });

        for (let i = 0; i < slideIds.length; i++) {
          const sId = slideIds[i];
          const slideText = await this._extractSlideTextById(sId);
          if (slideText && slideText.trim()) {
            fullTextParts.push(`--- Slide ${i + 1} ---\n${slideText.trim()}`);
          }
        }
      }
    } catch (e) {
      console.warn("PPT full text read error:", e);
    }
    return fullTextParts.join("\n\n");
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
    const isInsertCurrent = options.mode === "insert_current" || options.mode === "insert_current_slide";

    try {
      window.__isGeneratingSlides = true;
      if (typeof PowerPoint === 'undefined') {
        throw new Error("PowerPoint Office.js environment is not available.");
      }

      // If user selected "Insert on Current Slide", insert onto active slide without wiping existing shapes
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
                    const paras = shape.textFrame.textRange.paragraphs;
                    paras.load("items/text");
                    await context.sync();
                    if (paras.items) {
                      for (const p of paras.items) {
                        const pText = p.text || "";
                        const colonIdx = pText.indexOf(":");
                        const dashIdx = pText.indexOf("—");
                        const sepIdx = colonIdx > 0 ? colonIdx : (dashIdx > 0 ? dashIdx : -1);
                        if (sepIdx > 0 && sepIdx < 45 && typeof p.getSubstring === "function") {
                          try {
                            const leadIn = p.getSubstring(0, sepIdx + 1);
                            leadIn.font.bold = true;
                          } catch (_) {}
                        }
                      }
                      await context.sync();
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
