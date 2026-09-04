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

  // Safely extract all text from all shapes on a PowerPoint slide
  async _extractSlideText(context, slide) {
    const textLines = [];
    try {
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      if (!shapes.items || shapes.items.length === 0) return "";

      // Stage 1: Load shape metadata and textFrame.hasText safely
      const shapeTrackers = [];
      for (const shape of shapes.items) {
        try {
          shape.load("name, type");
          const tf = shape.textFrame;
          if (tf) {
            tf.load("hasText");
          }
          shapeTrackers.push({ shape, tf });
        } catch (e) {
          console.warn("PPT shape meta load warning:", e);
        }
      }
      await context.sync();

      // Stage 2: Load textRange.text for shapes where hasText is true
      const textShapes = [];
      for (const tracker of shapeTrackers) {
        try {
          if (tracker.tf && tracker.tf.hasText) {
            tracker.tf.textRange.load("text");
            textShapes.push(tracker);
          }
        } catch (e) {
          console.warn("PPT textRange load warning:", e);
        }
      }

      if (textShapes.length > 0) {
        await context.sync();
        for (const tracker of textShapes) {
          try {
            if (tracker.tf && tracker.tf.textRange && tracker.tf.textRange.text) {
              const txt = tracker.tf.textRange.text.trim();
              if (txt) textLines.push(txt);
            }
          } catch (e) {
            console.warn("PPT text read error:", e);
          }
        }
      }
    } catch (slideErr) {
      console.warn("PPT slide extract error:", slideErr);
    }
    return textLines.join("\n\n");
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
                const slideText = await this._extractSlideText(context, slide);
                if (slideText && slideText.trim()) {
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
              const shapeTrackers = [];
              for (const shape of selectedShapes.items) {
                const tf = shape.textFrame;
                if (tf) {
                  tf.load("hasText");
                  shapeTrackers.push({ shape, tf });
                }
              }
              await context.sync();

              const textShapes = [];
              for (const tracker of shapeTrackers) {
                if (tracker.tf && tracker.tf.hasText) {
                  tracker.tf.textRange.load("text");
                  textShapes.push(tracker);
                }
              }

              if (textShapes.length > 0) {
                await context.sync();
                const shapeTexts = [];
                for (const tracker of textShapes) {
                  if (tracker.tf && tracker.tf.textRange && tracker.tf.textRange.text) {
                    const txt = tracker.tf.textRange.text.trim();
                    if (txt) shapeTexts.push(txt);
                  }
                }
                if (shapeTexts.length > 0) {
                  selectedSlidesData.push({
                    slideNumber: 1,
                    id: "active-selection",
                    text: shapeTexts.join("\n\n")
                  });
                }
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
            const shapeTrackers = [];
            for (const shape of selection.items) {
              const tf = shape.textFrame;
              if (tf) {
                tf.load("hasText");
                shapeTrackers.push({ shape, tf });
              }
            }
            await context.sync();

            const textShapes = [];
            for (const tracker of shapeTrackers) {
              if (tracker.tf && tracker.tf.hasText) {
                tracker.tf.textRange.load("text");
                textShapes.push(tracker);
              }
            }

            if (textShapes.length > 0) {
              await context.sync();
              const texts = [];
              for (const tracker of textShapes) {
                if (tracker.tf && tracker.tf.textRange && tracker.tf.textRange.text) {
                  const t = tracker.tf.textRange.text.trim();
                  if (t) texts.push(t);
                }
              }
              selectedText = texts.join("\n\n");
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
    const fullTextParts = [];
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();

          if (slides.items && slides.items.length > 0) {
            for (let i = 0; i < slides.items.length; i++) {
              const slide = slides.items[i];
              const slideText = await this._extractSlideText(context, slide);
              if (slideText && slideText.trim()) {
                fullTextParts.push(`--- Slide ${i + 1} ---\n${slideText.trim()}`);
              }
            }
          }
        });
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
