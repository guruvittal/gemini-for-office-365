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

  // Read full presentation text across all slides and shapes safely using index-based traversal
  async getFullDocumentText() {
    let fullText = "";
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          const countResult = slides.getCount();
          await context.sync();

          const total = countResult.value || 0;
          const slideTexts = [];

          for (let i = 0; i < total; i++) {
            const slide = slides.getItemAt(i);
            const shapes = slide.shapes;
            shapes.load("items");
            await context.sync();

            const shapeTexts = [];
            if (shapes.items) {
              for (const shape of shapes.items) {
                if (shape.textFrame) {
                  const tr = shape.textFrame.textRange;
                  tr.load("text");
                  shapeTexts.push(tr);
                }
              }
            }

            if (shapeTexts.length > 0) {
              await context.sync();
              const lines = shapeTexts.map(t => (t.text || "").trim()).filter(Boolean);
              if (lines.length > 0) {
                slideTexts.push(`--- Slide ${i + 1} ---\n${lines.join("\n")}`);
              }
            }
          }

          fullText = slideTexts.join("\n\n");
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
