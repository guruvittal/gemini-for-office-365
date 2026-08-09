/**
 * Host Adapter Implementation for Microsoft PowerPoint (PowerPoint.run + Office Common API)
 * 
 * @author Sathya AG, Principal Architect, Google
 */

// Helper to compress high-res Vertex AI base64 PNGs into lightweight (~120KB) PNGs for instantaneous PowerPoint Office.js syncing
function compressImageForPowerPoint(base64Str, maxWidth = 900, maxHeight = 600) {
  return new Promise((resolve) => {
    try {
      const cleanRaw = base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
      const img = new Image();
      img.onload = () => {
        let w = img.width || 800;
        let h = img.height || 600;
        if (w > maxWidth || h > maxHeight) {
          const ratio = Math.min(maxWidth / w, maxHeight / h);
          w = Math.round(w * ratio);
          h = Math.round(h * ratio);
        }
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.clearRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        const dataUrl = canvas.toDataURL("image/png");
        const compressedBase64 = dataUrl.replace(/^data:image\/[^;]+;base64,/, "");
        resolve(compressedBase64);
      };
      img.onerror = () => {
        resolve(cleanRaw);
      };
      img.src = base64Str.startsWith("data:") ? base64Str : `data:image/png;base64,${cleanRaw}`;
    } catch (e) {
      resolve(base64Str.replace(/^data:image\/[^;]+;base64,/, "").replace(/[\r\n\s]+/g, "").trim());
    }
  });
}

// Host Adapter implementation for Microsoft PowerPoint (PowerPoint.run & Office Common API)
export class PPTAdapter {
  constructor() {
    this.name = "PowerPoint";
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

  // Read full presentation text
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

  // Insert AI content as executive PowerPoint slides with exact positioning & visuals
  async insertContent(htmlContent, rawText = "") {
    const debugStatus = document.getElementById("debugStatus");

    try {
      if (typeof PowerPoint === 'undefined') {
        throw new Error("PowerPoint Office.js environment is not available.");
      }

      if (debugStatus) debugStatus.innerText = "Parsing slides & preparing high-res graphics...";
      const slideStructures = await this.parseSlidesFromHtml(htmlContent, rawText);

      if (slideStructures.length === 0) {
        throw new Error("Slide parser returned 0 slide structures.");
      }

      // Pre-compress all images outside of PowerPoint.run to prevent Office.js context invalidation
      for (const slide of slideStructures) {
        slide.compressedImages = [];
        if (slide.base64Images && slide.base64Images.length > 0) {
          for (const rawImg of slide.base64Images) {
            try {
              const comp = await compressImageForPowerPoint(rawImg);
              if (comp && comp.length > 50) {
                slide.compressedImages.push(comp);
              }
            } catch (cErr) {
              console.warn("Image pre-compression notice:", cErr);
            }
          }
        }
      }

      for (let i = 0; i < slideStructures.length; i++) {
        const slideData = slideStructures[i];
        if (debugStatus) {
          debugStatus.innerText = `Creating slide ${i + 1}/${slideStructures.length}: "${slideData.title}"...`;
        }

        await PowerPoint.run(async (context) => {
          const presentation = context.presentation;
          presentation.slides.load("items");
          await context.sync();

          let newSlide;
          // Re-use initial blank slide #1 if opening a fresh deck, otherwise add a new slide
          if (i === 0 && presentation.slides.items.length === 1) {
            newSlide = presentation.slides.items[0];
          } else {
            presentation.slides.add();
            await context.sync();
            presentation.slides.load("items");
            await context.sync();
            newSlide = presentation.slides.items[presentation.slides.items.length - 1];
          }

          if (!newSlide) {
            throw new Error(`Could not access slide at index ${i}`);
          }

          // 2. Clear default placeholder shapes
          try {
            const existingShapes = newSlide.shapes;
            existingShapes.load("items");
            await context.sync();
            for (let s = existingShapes.items.length - 1; s >= 0; s--) {
              try {
                existingShapes.items[s].delete();
              } catch (delErr) {
                // ignore
              }
            }
            await context.sync();
          } catch (cleanErr) {
            console.warn("Placeholder cleanup notice:", cleanErr);
          }

          // 3. Add Top Blue Accent Line
          try {
            const accent = newSlide.shapes.addGeometricShape("Rectangle");
            accent.left = 40;
            accent.top = 30;
            accent.width = 880;
            accent.height = 4;
            accent.fill.setSolidColor("#0078d4");
            await context.sync();
          } catch (accErr) {
            console.warn("Accent line warning:", accErr);
          }

          // 4. Add Positioned Slide Title
          const hasImage = slideData.compressedImages && slideData.compressedImages.length > 0;
          const hasBody = slideData.body && slideData.body.trim().length > 0;
          const isTitleSlide = !hasBody && !hasImage;

          try {
            const titleText = slideData.title ? slideData.title : `Slide ${i + 1}`;
            const titleBox = newSlide.shapes.addTextBox(titleText);
            if (isTitleSlide) {
              titleBox.left = 80;
              titleBox.top = 180;
              titleBox.width = 800;
              titleBox.height = 140;
              titleBox.textFrame.textRange.font.size = 32;
            } else {
              titleBox.left = 40;
              titleBox.top = 45;
              titleBox.width = 880;
              titleBox.height = 55;
              titleBox.textFrame.textRange.font.size = 24;
            }
            titleBox.textFrame.textRange.font.bold = true;
            titleBox.textFrame.textRange.font.color = "#0078d4";
            titleBox.textFrame.textRange.font.name = "Segoe UI";
            await context.sync();
          } catch (tErr) {
            console.warn("Title shape positioning error:", tErr);
          }

          // 5. Add Native Visual Chart Image (Inside PowerPoint.run)
          const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
            ? slideData.compressedImages
            : (slideData.base64Images || []);

          const imgOpts = !hasBody
            ? { left: 140, top: 110, width: 680, height: 380 }
            : { left: 480, top: 110, width: 440, height: 340 };

          if (imagesToInsert.length > 0) {
            for (const rawOrCompressed of imagesToInsert) {
              const cleanBase64 = rawOrCompressed
                .replace(/^data:image\/[^;]+;base64,/i, "")
                .replace(/[\r\n\s]+/g, "")
                .trim();

              try {
                newSlide.shapes.addImage(cleanBase64, {
                  left: imgOpts.left,
                  top: imgOpts.top,
                  width: imgOpts.width,
                  height: imgOpts.height
                });
                await context.sync();
                console.log(`shapes.addImage succeeded on slide ${i + 1}`);
              } catch (iErr) {
                console.warn(`shapes.addImage notice on slide ${i + 1}:`, iErr);
              }
            }
          }

          // 6. Add Positioned Slide Body / Bullet Points (Left Column if image exists, or Full Width)
          if (hasBody) {
            const bodyWidth = hasImage ? 430 : 880;
            try {
              // Parse lines into structured paragraphs with rich rules
              const rawLines = slideData.body.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
              const cleanParagraphs = [];
              const paragraphMeta = [];

              for (const line of rawLines) {
                const isBullet = line.startsWith("•") || line.startsWith("-") || line.startsWith("*");
                const isSubHeader = line.endsWith(":") && !isBullet && line.length < 60;
                let cleanText = line.replace(/^[-*•]\s*/, "");
                cleanText = cleanText.replace(/\*\*(.*?)\*\*/g, "$1"); // Clean markdown bold markers

                if (isBullet) {
                  cleanParagraphs.push(`•  ${cleanText}`);
                } else {
                  cleanParagraphs.push(cleanText);
                }

                paragraphMeta.push({ isBullet, isSubHeader, rawLine: line });
              }

              const bodyBox = newSlide.shapes.addTextBox(cleanParagraphs.join("\n\n"));
              bodyBox.left = 40;
              bodyBox.top = 110;
              bodyBox.width = bodyWidth;
              bodyBox.height = 380;
              bodyBox.textFrame.textRange.font.size = 13;
              bodyBox.textFrame.textRange.font.name = "Segoe UI";
              bodyBox.textFrame.textRange.font.color = "#201f1e";
              bodyBox.textFrame.wordWrap = true;
              await context.sync();

              // Apply rich formatting to individual paragraphs
              try {
                bodyBox.textFrame.textRange.paragraphs.load("items");
                await context.sync();

                const paraItems = bodyBox.textFrame.textRange.paragraphs.items;
                for (let p = 0; p < paraItems.length; p++) {
                  const meta = paragraphMeta[p];
                  if (!meta) continue;

                  if (meta.isSubHeader) {
                    paraItems[p].font.bold = true;
                    paraItems[p].font.color = "#004e8c";
                    paraItems[p].font.size = 14;
                  } else if (meta.isBullet) {
                    paraItems[p].font.color = "#323130";
                    paraItems[p].font.size = 13;
                  }
                }
                await context.sync();
              } catch (pStyleErr) {
                console.warn("Paragraph rich styling warning:", pStyleErr);
              }
            } catch (bErr) {
              console.warn("Body shape positioning error:", bErr);
            }
          }

          // 7. Select newly created slide for Common API bridge
          try {
            newSlide.load("id");
            await context.sync();
            if (newSlide.id && context.presentation.setSelectedSlides) {
              context.presentation.setSelectedSlides([newSlide.id]);
              await context.sync();
            }
          } catch (selErr) {
            console.warn("Slide selection notice:", selErr);
          }
        });

        // 8. Universal Common API Image Injection (Runs outside PowerPoint.run on the selected slide)
        const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
          ? slideData.compressedImages
          : (slideData.base64Images || []);

        if (imagesToInsert.length > 0 && typeof Office !== "undefined" && Office.context && Office.context.document && Office.context.document.setSelectedDataAsync) {
          const hasBody = slideData.body && slideData.body.trim().length > 0;
          const imgOpts = !hasBody
            ? { left: 140, top: 110, width: 680, height: 380 }
            : { left: 480, top: 110, width: 440, height: 340 };

          for (const rawImg of imagesToInsert) {
            const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
            await new Promise((resolve) => {
              Office.context.document.setSelectedDataAsync(
                clean,
                {
                  coercionType: Office.CoercionType.Image,
                  imageLeft: imgOpts.left,
                  imageTop: imgOpts.top,
                  imageWidth: imgOpts.width,
                  imageHeight: imgOpts.height
                },
                (asyncResult) => {
                  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log(`Common API setSelectedDataAsync inserted image on slide ${i + 1}`);
                  } else {
                    console.warn(`Common API setSelectedDataAsync notice on slide ${i + 1}:`, asyncResult.error);
                  }
                  resolve();
                }
              );
            });
          }
        }
      }

      if (debugStatus) {
        debugStatus.innerText = `✅ Successfully created ${slideStructures.length} executive slides in PowerPoint!`;
      }

    } catch (err) {
      console.error("PPTAdapter Exception:", err);
      if (debugStatus) {
        debugStatus.innerHTML = `<span style="color:red; font-weight:bold;">PPT Error: ${err.message}</span>`;
      }
      throw err;
    }
  }

  // Parse HTML content into executive slide structures
  async parseSlidesFromHtml(htmlContent, rawText = "") {
    if (!htmlContent && !rawText) return [];

    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = htmlContent || rawText;

    // Remove citation notes / footers from slide body text
    const noteCallouts = tempDiv.querySelectorAll("blockquote, .note, [style*='background-color:#f0f6ff']");
    noteCallouts.forEach(n => n.remove());

    // Tier 1: DOM Header Elements (H1, H2, H3)
    const headerEls = Array.from(tempDiv.querySelectorAll("h1, h2, h3"));
    if (headerEls.length >= 2) {
      const slides = [];
      for (let i = 0; i < headerEls.length; i++) {
        const h = headerEls[i];
        const rawTitle = h.innerText ? h.innerText.trim() : `Slide ${i + 1}`;
        const title = rawTitle.replace(/^Slide\s*\d+[:\-]?\s*/i, "").trim();

        let bodyLines = [];
        let sectionImgs = [];
        let curr = h.nextElementSibling;

        while (curr && !["H1", "H2", "H3"].includes(curr.tagName)) {
          // Check for images inside this section
          const currImgs = Array.from(curr.querySelectorAll("img")).map(img => {
            return img.src || img.getAttribute("src") || "";
          }).filter(s => s && s.length > 50);

          if (curr.tagName === "IMG") {
            const imgSrc = curr.src || curr.getAttribute("src") || "";
            if (imgSrc.length > 50) currImgs.push(imgSrc);
          }

          if (currImgs.length > 0) sectionImgs.push(...currImgs);

          // Skip verified source footers
          if (curr.innerText && curr.innerText.includes("Verified Sources")) {
            curr = curr.nextElementSibling;
            continue;
          }

          // Format HTML Tables cleanly as executive bullet lines
          if (curr.tagName === "TABLE") {
            const rows = Array.from(curr.querySelectorAll("tr"));
            rows.forEach((tr, rIdx) => {
              const cells = Array.from(tr.querySelectorAll("th, td"))
                .map(c => (c.innerText || c.textContent || "").trim())
                .filter(Boolean);
              if (cells.length >= 2) {
                if (rIdx === 0 && tr.querySelector("th")) {
                  // Header row
                  bodyLines.push(`• [${cells.join(" | ")}]`);
                } else {
                  bodyLines.push(`• ${cells[0]}: ${cells.slice(1).join(" — ")}`);
                }
              }
            });
          } else if (curr.tagName === "UL" || curr.tagName === "OL") {
            Array.from(curr.querySelectorAll("li")).forEach(li => {
              const txt = (li.innerText || li.textContent || "").trim();
              if (txt) bodyLines.push(`• ${txt.replace(/^[-•*]\s*/, "")}`);
            });
          } else {
            const txt = (curr.innerText || curr.textContent || "").trim();
            if (txt) {
              if (txt.startsWith("-") || txt.startsWith("•") || txt.startsWith("*")) {
                bodyLines.push(`• ${txt.replace(/^[-•*]\s*/, "")}`);
              } else {
                bodyLines.push(txt);
              }
            }
          }
          curr = curr.nextElementSibling;
        }

        slides.push({
          title: title,
          body: bodyLines.join("\n\n"),
          base64Images: sectionImgs
        });
      }
      if (slides.length >= 2) return slides;
    }

    // Tier 2: Text-Based Header Splitting
    const textContent = tempDiv.innerText || tempDiv.textContent || rawText;
    const blocks = textContent
      .split(/(?:^|\n)(?=(?:#+|Slide\s*\d+|\d+\.\s+[A-Z]|\*\*\s*Slide))/gi)
      .map(b => b.trim())
      .filter(b => b.length > 0);

    if (blocks.length >= 2) {
      const slides = [];
      blocks.forEach((block, idx) => {
        const lines = block.split("\n").map(l => l.trim()).filter(l => l.length > 0);
        if (lines.length > 0) {
          const rawTitle = lines[0]
            .replace(/^[#*\s:]+/, "")
            .replace(/^\d+\.\s*/, "")
            .replace(/^Slide\s*\d+[:\-]?\s*/i, "")
            .trim();
          
          const title = rawTitle ? rawTitle : `Slide ${idx + 1}`;

          const bodyLines = lines.slice(1).map(l => {
            const clean = l.replace(/^[-•*]\s*/, "").trim();
            return clean ? `• ${clean}` : "";
          }).filter(l => l.length > 0);

          const slideImgs = allBase64Images[idx] ? [allBase64Images[idx]] : [];

          slides.push({
            title: title,
            body: bodyLines.length > 0 ? bodyLines.join("\n") : lines.slice(1).join("\n"),
            base64Images: slideImgs
          });
        }
      });
      if (slides.length >= 2) return slides;
    }

    // Tier 3: Single Slide Fallback
    const lines = textContent.split("\n").map(l => l.trim()).filter(l => l.length > 0);
    const title = lines[0] ? lines[0].replace(/^[#*\s:]+/, "") : "Executive Presentation";
    const body = lines.slice(1).join("\n");
    return [{ title: title, body: body || textContent, base64Images: allBase64Images }];
  }

  // Scan in-slide / in-shape @gemini commands for PowerPoint
  async checkInDocumentCommands(forceRun = false, callbacks = {}) {
    if (callbacks.onStatus) callbacks.onStatus("PowerPoint Adapter Ready (@gemini in shape)");
  }
}
