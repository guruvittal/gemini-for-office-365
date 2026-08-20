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
        // --- Executive Title Slide ---
        // White rounded card container
        slide.addShape(pres.ShapeType.roundRect, {
          x: 0.8,
          y: 1.2,
          w: 11.7,
          h: 4.8,
          fill: { color: "FFFFFF" },
          line: { color: "E1DFDD", width: 1 }
        });

        // Top Accent Stripe (Google Blue)
        slide.addShape(pres.ShapeType.rect, {
          x: 0.8,
          y: 1.2,
          w: 11.7,
          h: 0.12,
          fill: { color: "0078D4" }
        });

        // Main Title
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

        // Subtitle / Overview
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

        // Footer Badge
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
        // --- Executive Content Slide ---
        // Header Title
        slide.addText(slideData.title || `Slide ${idx + 1}`, {
          x: 0.8,
          y: 0.4,
          w: 11.7,
          h: 0.7,
          fontSize: 22,
          bold: true,
          color: "0078D4"
        });

        // Top Accent Divider
        slide.addShape(pres.ShapeType.rect, {
          x: 0.8,
          y: 1.15,
          w: 11.7,
          h: 0.04,
          fill: { color: "0078D4" }
        });

        // Check if body is formatted as a table / key-value list
        const bodyLines = (slideData.body || "").split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        const isTable = bodyLines.some(l => l.includes(":") || l.includes("—") || l.startsWith("|"));

        const contentWidth = imagesToInsert.length > 0 ? 6.5 : 11.7;

        if (isTable && bodyLines.length >= 2) {
          // Render as styled executive table
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
          // Render as clean bullet points
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

        // Add Chart / Visual Image if present
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

        // Slide Footer
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
  async insertContent(htmlContent, rawText = "") {
    const debugStatus = document.getElementById("debugStatus");

    try {
      if (typeof PowerPoint === 'undefined' || !PowerPoint.run) {
        throw new Error("PowerPoint Office.js environment is not available.");
      }

      if (debugStatus) debugStatus.innerText = "Parsing presentation structure...";
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

      // --- PRIMARY PATH: Atomic insertSlidesFromBase64 via PptxGenJS (PowerPointApi 1.2+) ---
      try {
        if (debugStatus) debugStatus.innerText = "Building executive deck package...";
        const base64Pptx = await this.generatePptxBase64(slideStructures, rawText);

        if (debugStatus) debugStatus.innerText = `Injecting ${slideStructures.length} slides into PowerPoint...`;

        await PowerPoint.run(async (context) => {
          context.presentation.insertSlidesFromBase64(base64Pptx, {
            formatting: "UseDestinationTheme"
          });
          await context.sync();
        });

        if (debugStatus) {
          debugStatus.innerText = `✅ Successfully created ${slideStructures.length} executive slides in PowerPoint!`;
        }
        return;

      } catch (pptxErr) {
        console.warn("insertSlidesFromBase64 failed, engaging 2-phase fallback:", pptxErr);
      }

      // --- SECONDARY FALLBACK: 2-Phase getCount() + getItemAt() (PowerPointApi 1.4+) ---
      let createdSlideCount = 0;
      for (let i = 0; i < slideStructures.length; i++) {
        const slideData = slideStructures[i];
        const hasImage = slideData.compressedImages && slideData.compressedImages.length > 0;
        const hasBody = slideData.body && slideData.body.trim().length > 0;
        const isTitleSlide = (i === 0 && !hasImage) || (!hasBody && !hasImage);
        const titleText = slideData.title || `Slide ${i + 1}`;

        let cleanBodyText = "";
        if (hasBody) {
          const rawLines = slideData.body.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
          cleanBodyText = rawLines.map(line => {
            const isBullet = line.startsWith("•") || line.startsWith("-") || line.startsWith("*");
            const cleanText = line.replace(/^[-*•]\s*/, "").replace(/\*\*(.*?)\*\*/g, "$1");
            return isBullet ? `•  ${cleanText}` : cleanText;
          }).join("\n\n");
        }

        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          const countResult = slides.getCount();
          slides.add();
          await context.sync();

          // Safely acquire the newly created slide proxy
          const newSlide = slides.getItemAt(countResult.value);

          // 1. Title
          newSlide.shapes.addTextBox(titleText, {
            left: isTitleSlide ? 60 : 40,
            top: isTitleSlide ? 120 : 35,
            width: isTitleSlide ? 600 : 640,
            height: isTitleSlide ? 90 : 55
          });

          // 2. Body
          if (cleanBodyText) {
            newSlide.shapes.addTextBox(cleanBodyText, {
              left: isTitleSlide ? 60 : 40,
              top: isTitleSlide ? 220 : 100,
              width: hasImage ? 330 : (isTitleSlide ? 600 : 640),
              height: 270
            });
          }

          // 3. Images
          const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
            ? slideData.compressedImages
            : (slideData.base64Images || []);

          if (imagesToInsert.length > 0) {
            const imgOpts = !hasBody
              ? { left: 100, top: 100, width: 520, height: 270 }
              : { left: 390, top: 100, width: 310, height: 260 };

            for (const rawOrCompressed of imagesToInsert) {
              const cleanBase64 = rawOrCompressed.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
              try {
                newSlide.shapes.addImage(cleanBase64, {
                  left: imgOpts.left,
                  top: imgOpts.top,
                  width: imgOpts.width,
                  height: imgOpts.height
                });
              } catch (_) {}
            }
          }

          await context.sync();
          createdSlideCount++;
        });
      }

      if (debugStatus) {
        debugStatus.innerText = `✅ Successfully created ${createdSlideCount} executive slides in PowerPoint!`;
      }

    } catch (err) {
      console.error("PPTAdapter Exception:", err);
      if (debugStatus) {
        debugStatus.innerHTML = `<span style="color:red; font-weight:bold;">PPT Error: ${err.message}</span>`;
      }
      throw err;
    }
  }

  // Parse HTML content and raw text into executive slide structures
  async parseSlidesFromHtml(htmlContent, rawText = "") {
    if (!htmlContent && !rawText) return [];

    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = htmlContent || "";

    // Extract all base64 images from HTML upfront
    const allBase64Images = Array.from(tempDiv.querySelectorAll("img")).map(img => {
      return img.src || img.getAttribute("src") || "";
    }).filter(s => s && s.length > 50);

    // Remove citation notes / footers
    const noteCallouts = tempDiv.querySelectorAll("blockquote, .note, [style*='background-color:#f0f6ff']");
    noteCallouts.forEach(n => n.remove());

    function cleanTitle(raw, fallback = "Executive Summary") {
      if (!raw) return fallback;
      let t = raw.trim();
      t = t.replace(/^[\p{Emoji}\u2600-\u27bf\s]+/gu, "");
      if (/^Slide\s*#?\d+\s*[:\-]?\s*.+/i.test(t)) {
        t = t.replace(/^Slide\s*#?\d+\s*[:\-]?\s*/i, "");
      }
      t = t.replace(/\*\*/g, "").replace(/^[#*\s:•\-\d.]+/g, "").replace(/\n.*/gs, "").trim();
      if (t.length > 70) {
        t = t.substring(0, 65).replace(/\s+\S*$/, "") + "...";
      }
      return t || fallback;
    }

    function cleanParagraphsToBullets(text) {
      if (!text) return "";
      return text.split(/\r?\n/)
        .map(l => l.trim())
        .filter(Boolean)
        .filter(l => !l.toLowerCase().includes("cannot generate a physical") && !l.toLowerCase().includes("since i cannot generate"))
        .map(l => {
          if (l.startsWith("•") || l.startsWith("-") || l.startsWith("*")) {
            return `• ${l.replace(/^[-*•]\s*/, "").replace(/\*\*/g, "")}`;
          }
          return l.replace(/\*\*/g, "");
        })
        .join("\n\n");
    }

    function parseMarkdownTableToBullets(text) {
      const lines = text.split(/\r?\n/);
      const bulletLines = [];
      let isFirstRow = true;

      for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed.startsWith("|") && trimmed.endsWith("|")) {
          if (trimmed.includes("---")) continue;
          const cells = trimmed.split("|").map(c => c.trim()).filter(Boolean);
          if (cells.length >= 2) {
            const firstLow = cells[0].toLowerCase();
            if (isFirstRow && (firstLow.includes("element") || firstLow.includes("focus") || firstLow.includes("initiative") || firstLow.includes("area") || firstLow.includes("topic") || firstLow.includes("column") || firstLow.includes("metric") || firstLow.includes("key"))) {
              isFirstRow = false;
              continue;
            }
            isFirstRow = false;
            bulletLines.push(`• ${cells[0]}: ${cells.slice(1).join(" — ")}`);
          } else if (cells.length === 1) {
            bulletLines.push(`• ${cells[0]}`);
          }
        } else if (trimmed.length > 0) {
          if (trimmed.startsWith("•") || trimmed.startsWith("-") || trimmed.startsWith("*")) {
            bulletLines.push(`• ${trimmed.replace(/^[-*•]\s*/, "")}`);
          } else {
            bulletLines.push(trimmed);
          }
        }
      }
      return bulletLines.join("\n\n");
    }

    const slides = [];

    // --- Strategy 1: Universal Element Traversal (Headers, Paragraphs with Slide X, Tables) ---
    const allElements = Array.from(tempDiv.querySelectorAll("h1, h2, h3, h4, p, table, ul, ol"));
    let currentSlide = null;

    for (const el of allElements) {
      const text = (el.innerText || el.textContent || "").trim();
      const isSlideHeader = /Slide\s*#?\d+/i.test(text) || /^[\s\p{Emoji}\u2600-\u27bf]*Slide\s*\d+/iu.test(text);

      if (["H1", "H2", "H3", "H4"].includes(el.tagName) || isSlideHeader) {
        if (currentSlide && (currentSlide.body_lines.length > 0 || currentSlide.title)) {
          slides.push({
            title: cleanTitle(currentSlide.title),
            body: cleanParagraphsToBullets(currentSlide.body_lines.join("\n")),
            base64Images: allBase64Images[slides.length] ? [allBase64Images[slides.length]] : []
          });
        }
        currentSlide = { title: text, body_lines: [] };
        continue;
      }

      if (!currentSlide) {
        // Collect pre-slide overview if substantive
        if (text && !text.toLowerCase().includes("cannot generate") && !text.toLowerCase().includes("since i cannot")) {
          currentSlide = { title: "Executive Overview", body_lines: [text] };
        }
        continue;
      }

      if (el.tagName === "TABLE") {
        const rows = Array.from(el.querySelectorAll("tr"));
        rows.forEach((tr, rIdx) => {
          const cells = Array.from(tr.querySelectorAll("th, td"))
            .map(c => (c.innerText || c.textContent || "").trim())
            .filter(Boolean);
          if (cells.length >= 2) {
            if (rIdx === 0 && tr.querySelector("th")) {
              return; // skip header row
            }
            currentSlide.body_lines.push(`• ${cells[0]}: ${cells.slice(1).join(" — ")}`);
          } else if (cells.length === 1) {
            currentSlide.body_lines.push(`• ${cells[0]}`);
          }
        });
      } else if (el.tagName === "UL" || el.tagName === "OL") {
        Array.from(el.querySelectorAll("li")).forEach(li => {
          const txt = (li.innerText || li.textContent || "").trim();
          if (txt) currentSlide.body_lines.push(`• ${txt.replace(/^[-•*]\s*/, "")}`);
        });
      } else if (el.tagName === "P") {
        if (text && !text.toLowerCase().includes("cannot generate a physical") && !text.toLowerCase().includes("since i cannot generate")) {
          currentSlide.body_lines.push(text);
        }
      }
    }

    function postProcessSlides(slideList) {
      return slideList.map((s, idx) => {
        let t = s.title;
        let b = s.body;
        // If title is generic like "Title" or "Title Slide" or "Slide 1" and body has a Title: row
        if (/^(?:title|title slide|slide\s*\d+)$/i.test(t.trim()) && b) {
          const m = b.match(/(?:•\s*Title:\s*|\*\*Title:\*\*\s*)([^\n]+)/i);
          if (m) {
            t = m[1].trim();
            b = b.replace(/•\s*Title:[^\n]+\n*/i, "").trim();
          }
        }
        return {
          ...s,
          title: t,
          body: b
        };
      });
    }

    if (currentSlide && (currentSlide.body_lines.length > 0 || currentSlide.title)) {
      slides.push({
        title: cleanTitle(currentSlide.title),
        body: cleanParagraphsToBullets(currentSlide.body_lines.join("\n")),
        base64Images: allBase64Images[slides.length] ? [allBase64Images[slides.length]] : []
      });
    }

    if (slides.length >= 2) return postProcessSlides(slides);

    // --- Strategy 2: Raw Text / Markdown Regex Splitting ---
    const textSource = rawText || tempDiv.innerText || tempDiv.textContent || "";
    const rawBlocks = textSource
      .split(/(?:^|\n)(?=(?:[\s\p{Emoji}\u2600-\u27bf]{0,6}Slide\s*#?\d+|#{1,4}\s+|(?:\d+\.\s+[A-Z])|\*\*[A-Z][^\n]{3,60}\*\*))/giu)
      .map(b => b.trim())
      .filter(b => b.length > 0 && !b.toLowerCase().includes("cannot generate a physical") && !b.toLowerCase().includes("since i cannot generate"));

    if (rawBlocks.length >= 2) {
      const parsedSlides = [];
      rawBlocks.forEach((block, idx) => {
        const lines = block.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        if (lines.length > 0) {
          const title = cleanTitle(lines[0], `Slide ${idx + 1}`);
          const bodyText = lines.slice(1).join("\n");
          const bullets = parseMarkdownTableToBullets(bodyText);
          parsedSlides.push({
            title: title,
            body: bullets || cleanParagraphsToBullets(lines.join("\n")),
            base64Images: allBase64Images[idx] ? [allBase64Images[idx]] : []
          });
        }
      });
      if (parsedSlides.length >= 2) return postProcessSlides(parsedSlides);
    }

    // --- Strategy 3: Multi-Paragraph Chunking Fallback ---
    const allLines = textSource.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (allLines.length > 4) {
      const chunked = [];
      const chunkSize = 4;
      for (let i = 0; i < allLines.length; i += chunkSize) {
        const chunk = allLines.slice(i, i + chunkSize);
        const firstLine = chunk[0];
        const isHeader = firstLine.length < 60 && !firstLine.startsWith("•");
        const title = isHeader ? cleanTitle(firstLine) : (i === 0 ? "Executive Summary" : `Key Insights (Part ${Math.floor(i / chunkSize) + 1})`);
        const bodyLines = isHeader ? chunk.slice(1) : chunk;
        chunked.push({
          title: title,
          body: cleanParagraphsToBullets(bodyLines.join("\n")),
          base64Images: allBase64Images[chunked.length] ? [allBase64Images[chunked.length]] : []
        });
      }
      if (chunked.length >= 2) return postProcessSlides(chunked);
    }

    // --- Strategy 4: Clean Single Slide Fallback ---
    const firstLine = allLines[0] || "Executive Presentation";
    const title = firstLine.length < 60 ? cleanTitle(firstLine) : "Executive Summary";
    const bodyContent = firstLine.length < 60 ? allLines.slice(1).join("\n") : allLines.join("\n");
    return postProcessSlides([{
      title: title,
      body: cleanParagraphsToBullets(bodyContent),
      base64Images: allBase64Images
    }]);
  }

  // Scan in-slide / in-shape @gemini commands for PowerPoint
  async checkInDocumentCommands(forceRun = false, callbacks = {}) {
    if (callbacks.onStatus) callbacks.onStatus("PowerPoint Adapter Ready (@gemini in shape)");
  }
}
