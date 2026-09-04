/**
 * Host Adapter Implementation for Microsoft PowerPoint (PowerPoint.run + Office Common API)
 * 
 * Provides complete multi-slide presentation generation, intelligent outline/table
 * unpacking, visual slide deck preview, executive PowerPoint styling, and in-taskpane diagnostics.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { parseSlides } from './ppt/slideParser.js';
import { buildPresentation, compressImageForPowerPoint } from './ppt/slideBuilder.js';
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

  // Read text and metadata from currently highlighted/selected slide(s) in PowerPoint
  async getSelectedSlidesText() {
    const selectedSlidesData = [];
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          // 1. Check for selected slides via PowerPointApi 1.5+
          if (context.presentation.getSelectedSlides) {
            const selectedSlides = context.presentation.getSelectedSlides();
            selectedSlides.load("items");
            await context.sync();

            if (selectedSlides.items && selectedSlides.items.length > 0) {
              for (let i = 0; i < selectedSlides.items.length; i++) {
                const slide = selectedSlides.items[i];
                const shapes = slide.shapes;
                shapes.load("items");
                await context.sync();

                let slideText = "";
                for (const shape of shapes.items) {
                  if (shape.textFrame) {
                    const tr = shape.textFrame.textRange;
                    tr.load("text");
                    await context.sync();
                    if (tr.text && tr.text.trim()) {
                      slideText += tr.text.trim() + "\n";
                    }
                  }
                }
                if (slideText.trim()) {
                  selectedSlidesData.push({
                    slideNumber: i + 1,
                    id: slide.id || `slide-${i + 1}`,
                    text: slideText.trim()
                  });
                }
              }
            }
          }

          // 2. If no slide selection found, check if any shapes on the active slide are selected
          if (selectedSlidesData.length === 0 && context.presentation.getSelectedShapes) {
            const selectedShapes = context.presentation.getSelectedShapes();
            selectedShapes.load("items");
            await context.sync();

            if (selectedShapes.items && selectedShapes.items.length > 0) {
              let shapeText = "";
              for (const shape of selectedShapes.items) {
                if (shape.textFrame) {
                  const tr = shape.textFrame.textRange;
                  tr.load("text");
                  await context.sync();
                  if (tr.text && tr.text.trim()) {
                    shapeText += tr.text.trim() + "\n";
                  }
                }
              }
              if (shapeText.trim()) {
                selectedSlidesData.push({
                  slideNumber: 1,
                  id: "active-slide",
                  text: shapeText.trim()
                });
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

  // Read currently highlighted text frame or shape text in PowerPoint
  async getSelectedText() {
    let selectedText = "";
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          const selection = context.presentation.getSelectedShapes();
          selection.load("items");
          await context.sync();

          if (selection.items && selection.items.length > 0) {
            const shape = selection.items[0];
            if (shape.textFrame) {
              const textRange = shape.textFrame.textRange;
              textRange.load("text");
              await context.sync();
              selectedText = textRange.text ? textRange.text.trim() : "";
            }
          }
        });
      }
    } catch (err) {
      console.warn("PowerPoint selection read error:", err);
    }
    return selectedText;
  }

  // Read full presentation text across all slides and shapes
  async getFullDocumentText() {
    let fullText = "";
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          for (const s of slides.items) {
            const shapes = s.shapes;
            shapes.load("items");
            await context.sync();
            for (const shape of shapes.items) {
              if (shape.textFrame) {
                const tr = shape.textFrame.textRange;
                tr.load("text");
                await context.sync();
                if (tr.text) fullText += tr.text + "\n";
              }
            }
          }
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

    try {
      if (typeof PowerPoint === 'undefined') {
        throw new Error("PowerPoint Office.js environment is not available.");
      }

      if (debugStatus) debugStatus.innerText = "Parsing presentation structure...";
      const slideStructures = await this.parseSlidesFromHtml(htmlContent, rawText);

      if (!slideStructures || slideStructures.length === 0) {
        throw new Error("Slide parser returned 0 slide structures.");
      }

      const onProgress = (prog) => {
        const msg = `⚡ Creating slide ${prog.current}/${prog.total}: "${prog.title}"...`;
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
        debugStatus.innerText = `✅ Created ${slideStructures.length} slides in PowerPoint!`;
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
    }
  }

  // Scan in-slide / in-shape @gemini commands for PowerPoint
  async checkInDocumentCommands(forceRun = false, callbacks = {}) {
    if (callbacks.onStatus) callbacks.onStatus("PowerPoint Adapter Ready (@gemini in shape)");
  }
}
