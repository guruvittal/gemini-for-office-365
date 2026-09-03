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

  // Build an executive PowerPoint presentation (.pptx) as a Base64 string using PptxGenJS
  async generatePptxBase64(slideStructures, rawText = "") {
    const PptxConstructor = (typeof window !== 'undefined' && window.PptxGenJS) ? window.PptxGenJS : null;
    if (!PptxConstructor) {
      throw new Error("PptxGenJS browser bundle is not loaded.");
    }

    const pres = new PptxConstructor();
    pres.layout = "LAYOUT_16x9";
    pres.author = "Gemini Enterprise";
    pres.title = "Executive Presentation";

    for (let idx = 0; idx < slideStructures.length; idx++) {
      const slideData = slideStructures[idx];
      const slide = pres.addSlide();
      const isTitle = idx === 0 && slideStructures.length > 1;
      slide.background = { color: isTitle ? "F8F9FA" : "FFFFFF" };

      const hasImages = slideData.compressedImages && slideData.compressedImages.length > 0;
      const imagesToInsert = hasImages ? slideData.compressedImages : (slideData.base64Images || []);

      if (isTitle) {
        // Executive Title Slide
        slide.addShape(pres.ShapeType.roundRect, {
          x: 0.8,
          y: 1.2,
          w: 11.7,
          h: 4.8,
          fill: { color: "FFFFFF" },
          line: { color: "E1DFDD", width: 1 }
        });

        slide.addShape(pres.ShapeType.rect, {
          x: 0.8,
          y: 1.2,
          w: 11.7,
          h: 0.12,
          fill: { color: "0078D4" }
        });

        slide.addText(slideData.title || "Executive Briefing", {
          x: 1.2,
          y: 1.8,
          w: 10.9,
          h: 1.4,
          fontSize: 30,
          bold: true,
          color: "0078D4",
          align: "left"
        });

        const subTitleText = slideData.body ? slideData.body.replace(/^[•\s*-]+/gm, "").trim() : "Strategic Overview & Executive Summary";
        slide.addText(subTitleText, {
          x: 1.2,
          y: 3.3,
          w: 10.9,
          h: 1.2,
          fontSize: 16,
          color: "605E5C",
          align: "left"
        });

        slide.addText("Generated by Gemini Enterprise • Powered by Google Cloud", {
          x: 1.2,
          y: 5.2,
          w: 10.9,
          h: 0.4,
          fontSize: 11,
          color: "8A8886",
          align: "left"
        });

      } else {
        // Executive Content Slide
        slide.addText(slideData.title || `Slide ${idx + 1}`, {
          x: 0.8,
          y: 0.4,
          w: 11.7,
          h: 0.7,
          fontSize: 22,
          bold: true,
          color: "0078D4"
        });

        slide.addShape(pres.ShapeType.rect, {
          x: 0.8,
          y: 1.15,
          w: 11.7,
          h: 0.04,
          fill: { color: "0078D4" }
        });

        const bodyLines = (slideData.body || "").split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        const isTable = bodyLines.some(l => l.includes(":") || l.includes("—") || l.startsWith("|"));
        const contentWidth = imagesToInsert.length > 0 ? 6.5 : 11.7;

        if (isTable && bodyLines.length >= 2) {
          const tableRows = [];
          for (const line of bodyLines) {
            const clean = line.replace(/^[•*\-\s|]+/, "").replace(/[|\s]+$/, "");
            const parts = clean.split(/[:—|]/).map(p => p.trim()).filter(Boolean);
            if (parts.length >= 2) {
              tableRows.push([
                { text: parts[0], options: { fill: { color: tableRows.length % 2 === 0 ? "F8F9FA" : "FFFFFF" }, color: "004578", bold: true, fontSize: 12, valign: "middle" } },
                { text: parts.slice(1).join(" — "), options: { fill: { color: tableRows.length % 2 === 0 ? "F8F9FA" : "FFFFFF" }, color: "323130", fontSize: 12, valign: "middle" } }
              ]);
            } else if (parts.length === 1) {
              tableRows.push([
                { text: "•", options: { fill: { color: "FFFFFF" }, color: "0078D4", bold: true, fontSize: 12 } },
                { text: parts[0], options: { fill: { color: "FFFFFF" }, color: "323130", fontSize: 12 } }
              ]);
            }
          }

          if (tableRows.length > 0) {
            slide.addTable(tableRows, {
              x: 0.8,
              y: 1.4,
              w: contentWidth,
              colW: imagesToInsert.length > 0 ? [2.0, 4.5] : [3.0, 8.7],
              border: { type: "solid", pt: 1, color: "E1DFDD" },
              margin: [5, 8, 5, 8]
            });
          }
        } else {
          const bulletItems = bodyLines.map(line => {
            const clean = line.replace(/^[-*•]\s*/, "").replace(/\*\*(.*?)\*\*/g, "$1");
            return {
              text: clean,
              options: {
                fontSize: 13,
                color: "323130",
                bullet: true,
                lineSpacing: 22
              }
            };
          });

          if (bulletItems.length > 0) {
            slide.addText(bulletItems, {
              x: 0.8,
              y: 1.4,
              w: contentWidth,
              h: 4.8,
              valign: "top"
            });
          }
        }

        if (imagesToInsert.length > 0) {
          const rawImg = imagesToInsert[0];
          const cleanBase64 = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
          try {
            slide.addImage({
              data: `image/png;base64,${cleanBase64}`,
              x: 7.6,
              y: 1.4,
              w: 4.9,
              h: 4.8
            });
          } catch (imgErr) {
            console.warn("PptxGenJS image insert warning:", imgErr);
          }
        }

        slide.addText(`Slide ${idx + 1} • Gemini Enterprise`, {
          x: 0.8,
          y: 6.7,
          w: 11.7,
          h: 0.3,
          fontSize: 10,
          color: "8A8886",
          align: "right"
        });
      }
    }

    const base64Data = await pres.write({ outputType: "base64" });
    return base64Data;
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
