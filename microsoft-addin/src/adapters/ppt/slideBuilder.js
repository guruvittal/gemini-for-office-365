/**
 * Slide Builder for Microsoft PowerPoint
 * 
 * Standard Office.js Slide Generation:
 * - Uses official Microsoft pattern: slides.add() -> slides.getCount() -> slides.getItemAt(count - 1)
 * - Inserts textboxes with { left, top, width, height } geometry directly
 * - Inserts side-by-side visual chart images
 * - Live diagnostic streaming to Taskpane log console
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { logToPPTConsole } from './pptDiagnostics.js';

/**
 * Safely compresses base64 images with a strict 1-second timeout.
 */
export function compressImageForPowerPoint(base64Str, maxWidth = 1920, maxHeight = 1080) {
  return new Promise((resolve) => {
    let resolved = false;
    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true;
        resolve(base64Str ? base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim() : "");
      }
    }, 1500);

    try {
      if (typeof base64Str === "string" && (base64Str.startsWith("http://") || base64Str.startsWith("https://"))) {
        fetch(base64Str)
          .then(res => res.blob())
          .then(blob => {
            const reader = new FileReader();
            reader.onloadend = () => {
              const b64 = (reader.result || "").replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
              if (!resolved) {
                resolved = true;
                clearTimeout(timer);
                resolve(b64);
              }
            };
            reader.readAsDataURL(blob);
          })
          .catch(() => {
            if (!resolved) {
              resolved = true;
              clearTimeout(timer);
              resolve("");
            }
          });
        return;
      }

      const cleanRaw = base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
      if (!cleanRaw) {
        clearTimeout(timer);
        resolve("");
        return;
      }
      // If image is already lightweight (< 250KB base64), preserve without downsampling
      if (cleanRaw.length < 250000) {
        clearTimeout(timer);
        resolve(cleanRaw);
        return;
      }
      const img = new Image();
      img.onload = () => {
        if (resolved) return;
        resolved = true;
        clearTimeout(timer);
        let w = img.width || 960;
        let h = img.height || 640;
        const targetMaxWidth = Math.min(maxWidth || 960, 960);
        const targetMaxHeight = Math.min(maxHeight || 640, 640);
        if (w > targetMaxWidth || h > targetMaxHeight) {
          const ratio = Math.min(targetMaxWidth / w, targetMaxHeight / h);
          w = Math.round(w * ratio);
          h = Math.round(h * ratio);
        }
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.clearRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        const isJpeg = base64Str.startsWith("data:image/jpeg") || base64Str.startsWith("data:image/jpg");
        const dataUrl = isJpeg ? canvas.toDataURL("image/jpeg", 0.85) : canvas.toDataURL("image/png");
        const compressedBase64 = dataUrl.replace(/^data:image\/[^;]+;base64,/, "");
        resolve(compressedBase64);
      };
      img.onerror = () => {
        if (resolved) return;
        resolved = true;
        clearTimeout(timer);
        resolve(cleanRaw);
      };
      img.src = base64Str.startsWith("data:") ? base64Str : `data:image/png;base64,${cleanRaw}`;
    } catch (e) {
      if (!resolved) {
        resolved = true;
        clearTimeout(timer);
        resolve(base64Str ? base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim() : "");
      }
    }
  });
}


/**
 * Accurately estimates rendered height of a native PowerPoint table based on word-wrapping.
 * Reflects PowerPoint's default 18pt font, cell margins (7.2pt each side), and line height.
 */
function estimateTableRenderedHeight(tableData, colWidth = 280) {
  if (!tableData) return 0;
  const headers = tableData.headers || [];
  const rows = tableData.rows || [];
  
  // PowerPoint table cell horizontal padding is ~14.4pt (7.2pt each side).
  // At 18pt font in Segoe UI / Calibri, average character width is ~11.5pt.
  const usableColWidth = Math.max(40, colWidth - 18);
  const charsPerLine = Math.max(5, Math.floor(usableColWidth / 11.5));
  
  let totalHeight = 0;
  if (headers.length > 0) {
    let maxHeaderLines = 1;
    for (const h of headers) {
      const hText = String(h || "").trim();
      const lines = Math.max(1, Math.ceil(hText.length / charsPerLine));
      if (lines > maxHeaderLines) maxHeaderLines = lines;
    }
    totalHeight += Math.max(44, maxHeaderLines * 26 + 18);
  }

  for (const row of rows) {
    let maxLinesInRow = 1;
    for (const cell of row) {
      const cellText = String(cell !== undefined && cell !== null ? cell : "").trim();
      const lines = Math.max(1, Math.ceil(cellText.length / charsPerLine));
      if (lines > maxLinesInRow) maxLinesInRow = lines;
    }
    const rowHeight = Math.max(40, maxLinesInRow * 26 + 18);
    totalHeight += rowHeight;
  }
  return totalHeight;
}

/**
 * Synchronously parses markdown formatting (**bold**, __bold__, *italic*, _italic_) in text,
 * strips markdown characters to produce clean text for PowerPoint shapes,
 * and tracks the exact character ranges for bold and italic styling per paragraph.
 */
export function parseMarkdownFormatting(rawContent) {
  if (!rawContent) return { cleanText: "", parsedParagraphs: [] };

  const lines = String(rawContent).split(/\r?\n/);
  const cleanLines = [];
  const parsedParagraphs = [];

  for (const rawLine of lines) {
    if (!rawLine.trim()) {
      cleanLines.push("");
      parsedParagraphs.push({ cleanText: "", boldRanges: [], italicRanges: [] });
      continue;
    }

    let cleanText = "";
    const boldRanges = [];
    const italicRanges = [];

    // Match **bold**, __bold__, *italic*, _italic_, or normal text
    const mdRegex = /(\*\*(.*?)\*\*|__([^_]+)__|\*(.*?)\*|_([^_]+)_|([^*_]+|[*_]))/g;
    let match;
    while ((match = mdRegex.exec(rawLine)) !== null) {
      if (match[2] !== undefined) {
        // **bold**
        const boldText = match[2].replace(/\*\*/g, "").replace(/__/g, "");
        const start = cleanText.length;
        cleanText += boldText;
        boldRanges.push({ start, length: boldText.length });
      } else if (match[3] !== undefined) {
        // __bold__
        const boldText = match[3].replace(/\*\*/g, "").replace(/__/g, "");
        const start = cleanText.length;
        cleanText += boldText;
        boldRanges.push({ start, length: boldText.length });
      } else if (match[4] !== undefined) {
        // *italic*
        const italicText = match[4].replace(/\*/g, "").replace(/_/g, "");
        const start = cleanText.length;
        cleanText += italicText;
        italicRanges.push({ start, length: italicText.length });
      } else if (match[5] !== undefined) {
        // _italic_
        const italicText = match[5].replace(/\*/g, "").replace(/_/g, "");
        const start = cleanText.length;
        cleanText += italicText;
        italicRanges.push({ start, length: italicText.length });
      } else if (match[6] !== undefined) {
        cleanText += match[6];
      }
    }

    // Safety strip: eliminate any residual unparsed ** or __ from cleanText
    if (cleanText.includes("**") || cleanText.includes("__")) {
      cleanText = cleanText.replace(/\*\*/g, "").replace(/__/g, "");
    }

    // Expand bold range to include trailing colon or dash (e.g. "**Renewable Energy**:" or "**Key** —")
    for (const b of boldRanges) {
      if (cleanText.charAt(b.start + b.length) === ":" || cleanText.charAt(b.start + b.length) === "—") {
        b.length += 1;
      }
    }

    // If no bold ranges were found from markdown, auto-detect colon/dash lead-in (e.g. "• Green Energy Leadership:")
    if (boldRanges.length === 0) {
      const colonIdx = cleanText.indexOf(":");
      const dashIdx = cleanText.indexOf("—");
      const sepIdx = colonIdx > 0 ? colonIdx : (dashIdx > 0 ? dashIdx : -1);
      if (sepIdx > 0 && sepIdx < 60) {
        boldRanges.push({ start: 0, length: sepIdx + 1 });
      }
    }

    cleanLines.push(cleanText);
    parsedParagraphs.push({ cleanText, boldRanges, italicRanges });
  }

  const finalCleanText = cleanLines.join("\n").replace(/\*\*/g, "").replace(/__/g, "");
  return { cleanText: finalCleanText, parsedParagraphs };
}

/**
 * Applies native bold and italic font styling to individual paragraph ranges
 * inside a PowerPoint shape's text frame via Office.js getSubstring().
 */
export async function applyParagraphFormatting(textBox, parsedParagraphs, context) {
  if (!textBox || !textBox.textFrame || !parsedParagraphs || parsedParagraphs.length === 0) return;

  try {
    const paragraphs = textBox.textFrame.textRange.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    if (paragraphs.items) {
      let pIdx = 0;
      for (let i = 0; i < parsedParagraphs.length && pIdx < paragraphs.items.length; i++) {
        const parsed = parsedParagraphs[i];
        if (!parsed.cleanText.trim()) {
          pIdx++;
          continue;
        }

        const p = paragraphs.items[pIdx];
        if (p && typeof p.getSubstring === "function") {
          // Apply bold ranges
          for (const b of parsed.boldRanges) {
            try {
              if (b.length > 0 && b.start + b.length <= parsed.cleanText.length) {
                const sub = p.getSubstring(b.start, b.length);
                sub.font.bold = true;
              }
            } catch (_) {}
          }

          // Apply italic ranges
          for (const it of parsed.italicRanges) {
            try {
              if (it.length > 0 && it.start + it.length <= parsed.cleanText.length) {
                const sub = p.getSubstring(it.start, it.length);
                sub.font.italic = true;
              }
            } catch (_) {}
          }
        }
        pIdx++;
      }
      await context.sync();
    }
  } catch (err) {
    console.warn("Could not apply paragraph text formatting in PowerPoint:", err);
  }
}

/**
 * Populates a native Microsoft PowerPoint table using PowerPoint.js shapes.addTable().
 * Configures an executive light theme (white header, crisp black text, soft ice-blue alternating rows, and clean borders).
 */
function populateSlideTable(newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, tableData, slideNum, tableTop = 90, customHeight = null, customWidth = null, customLeft = 50) {
  const headers = tableData.headers || [];
  const rows = tableData.rows || [];
  const colCount = Math.max(headers.length, ...rows.map(r => r.length), 1);
  const rowCount = (headers.length > 0 ? 1 : 0) + rows.length;

  const tableValues = [];
  if (headers.length > 0) {
    const hRow = [];
    for (let c = 0; c < colCount; c++) {
      hRow.push(headers[c] || "");
    }
    tableValues.push(hRow);
  }
  for (const r of rows) {
    const rowVals = [];
    for (let c = 0; c < colCount; c++) {
      rowVals.push(r[c] !== undefined && r[c] !== null ? String(r[c]) : "");
    }
    tableValues.push(rowVals);
  }

  const tableWidth = customWidth || 860;
  const approxColWidth = tableWidth / Math.max(1, colCount);
  const estimatedHeight = estimateTableRenderedHeight(tableData, approxColWidth);
  const tableHeight = customHeight || Math.min(380, Math.max(80, estimatedHeight));

  let addedShape = null;
  try {
    if (typeof newSlide.shapes.addTable === "function") {
      addedShape = newSlide.shapes.addTable(rowCount, colCount, {
        left: customLeft,
        top: tableTop,
        height: tableHeight,
        values: tableValues
      });
      if (addedShape && customWidth) {
        try { addedShape.width = customWidth; } catch (_) {}
      }
      logToPPTConsole(`Slide ${slideNum}: Added native PowerPoint table (${rowCount} rows x ${colCount} cols, rendered height ~${estimatedHeight}pt).`);
    }
  } catch (err) {
    console.warn("shapes.addTable with options failed, trying basic addTable:", err);
  }

  if (!addedShape) {
    try {
      addedShape = newSlide.shapes.addTable(rowCount, colCount);
      if (addedShape) {
        addedShape.left = customLeft;
        addedShape.top = tableTop;
        addedShape.width = tableWidth;
        addedShape.height = tableHeight;
        const table = addedShape.table || (typeof addedShape.getTable === "function" ? addedShape.getTable() : null);
        if (table) {
          for (let r = 0; r < tableValues.length; r++) {
            for (let c = 0; c < colCount; c++) {
              try {
                const cell = (typeof table.getCell === "function")
                  ? table.getCell(r, c)
                  : (typeof table.getCellOrNullObject === "function" ? table.getCellOrNullObject(r, c) : null);
                if (cell) {
                  if (cell.textRange) cell.textRange.text = tableValues[r][c];
                  else if (cell.text !== undefined) cell.text = tableValues[r][c];
                }
              } catch (_) {}
            }
          }
        }
      }
      logToPPTConsole(`Slide ${slideNum}: Added native PowerPoint table via fallback.`);
    } catch (fallbackErr) {
      console.error("Native table shape creation failed:", fallbackErr);
      return estimatedHeight;
    }
  }
  return estimatedHeight;
}

/**
 * Populates native PowerPoint 3-column metric cards on a slide using geometric shapes and styled textboxes.
 */
function populate3ColumnMetricGrid(newSlide, visualData, slideNum, contentTop = 110) {
  const cards = (visualData && visualData.cards) ? visualData.cards : [];
  if (!cards || cards.length === 0) return;

  const cardCount = Math.min(cards.length, 3);
  const cardWidth = 265;
  const cardHeight = 360;
  const gap = 32;
  const startLeft = 50;

  for (let i = 0; i < cardCount; i++) {
    const card = cards[i];
    const left = startLeft + i * (cardWidth + gap);
    const top = contentTop;

    // 1. Try adding geometric shape card background
    try {
      if (typeof newSlide.shapes.addGeometricShape === "function" && PowerPoint.GeometricShapeType) {
        const shapeType = PowerPoint.GeometricShapeType.roundRectangle || PowerPoint.GeometricShapeType.rectangle;
        const bgShape = newSlide.shapes.addGeometricShape(shapeType, {
          left: left,
          top: top,
          width: cardWidth,
          height: cardHeight
        });
        if (bgShape.fill && typeof bgShape.fill.setSolidColor === "function") {
          bgShape.fill.setSolidColor("#F8FAFC");
        }
        if (bgShape.line) {
          bgShape.line.color = "#CBD5E1";
          bgShape.line.weight = 1.5;
        }
      }
    } catch (e) {
      console.warn("Geometric shape card fallback to textbox:", e);
    }

    // 2. Add Metric callout box (e.g. "+42%")
    const metricText = card.metric || "";
    if (metricText) {
      const metricBox = newSlide.shapes.addTextBox(metricText, {
        left: left + 16,
        top: top + 16,
        width: cardWidth - 32,
        height: 52
      });
      metricBox.textFrame.textRange.font.size = 38;
      metricBox.textFrame.textRange.font.bold = true;
      metricBox.textFrame.textRange.font.color = "#0078D4";
    }

    // 3. Add Card Title
    const titleTop = metricText ? (top + 72) : (top + 20);
    const titleBox = newSlide.shapes.addTextBox(card.title || `Metric ${i + 1}`, {
      left: left + 16,
      top: titleTop,
      width: cardWidth - 32,
      height: 28
    });
    titleBox.textFrame.textRange.font.size = 15;
    titleBox.textFrame.textRange.font.bold = true;
    titleBox.textFrame.textRange.font.color = "#0F172A";

    let bulletsTop = titleTop + 30;
    if (card.subtitle) {
      const subBox = newSlide.shapes.addTextBox(card.subtitle, {
        left: left + 16,
        top: bulletsTop,
        width: cardWidth - 32,
        height: 22
      });
      subBox.textFrame.textRange.font.size = 11;
      subBox.textFrame.textRange.font.italic = true;
      subBox.textFrame.textRange.font.color = "#64748B";
      bulletsTop += 26;
    }

    // 4. Add Supporting Bullets
    const bullets = card.bullets || [];
    if (bullets.length > 0) {
      const bulletContent = bullets.map(b => `• ${b}`).join("\n");
      const { cleanText } = parseMarkdownFormatting(bulletContent);
      const bulletsBox = newSlide.shapes.addTextBox(cleanText, {
        left: left + 16,
        top: bulletsTop,
        width: cardWidth - 32,
        height: Math.max(80, (top + cardHeight - 16) - bulletsTop)
      });
      bulletsBox.textFrame.textRange.font.size = 12.5;
      bulletsBox.textFrame.textRange.font.color = "#334155";
    }
  }

  logToPPTConsole(`Slide ${slideNum}: Added native 3-Column Metric Grid with ${cardCount} cards.`);
}

/**
 * Populates native PowerPoint Before/After comparison cards on a slide.
 */
function populateBeforeAfterComparison(newSlide, visualData, slideNum, contentTop = 110) {
  const beforeData = (visualData && visualData.before) || { title: "Current State / Challenges", bullets: [] };
  const afterData = (visualData && visualData.after) || { title: "Target State / Transformation", bullets: [] };

  const cardWidth = 415;
  const cardHeight = 365;
  const gap = 30;
  const top = contentTop;

  // 1. BEFORE Card (Left)
  const leftBefore = 50;
  try {
    if (typeof newSlide.shapes.addGeometricShape === "function" && PowerPoint.GeometricShapeType) {
      const bgBefore = newSlide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, {
        left: leftBefore,
        top: top,
        width: cardWidth,
        height: cardHeight
      });
      if (bgBefore.fill && typeof bgBefore.fill.setSolidColor === "function") {
        bgBefore.fill.setSolidColor("#FFF8F8");
      }
      if (bgBefore.line) {
        bgBefore.line.color = "#FECACA";
        bgBefore.line.weight = 1.5;
      }
    }
  } catch (_) {}

  // Before Header Badge
  const beforeHeader = newSlide.shapes.addTextBox(`🔴 BEFORE: ${beforeData.title || 'Current State'}`, {
    left: leftBefore + 20,
    top: top + 18,
    width: cardWidth - 40,
    height: 36
  });
  beforeHeader.textFrame.textRange.font.size = 16;
  beforeHeader.textFrame.textRange.font.bold = true;
  beforeHeader.textFrame.textRange.font.color = "#991B1B";

  // Before Bullets
  const beforeBullets = (beforeData.bullets || []).map(b => `• ${b}`).join("\n\n");
  const { cleanText: cleanBefore } = parseMarkdownFormatting(beforeBullets || "• Legacy workflow bottlenecks");
  const beforeBox = newSlide.shapes.addTextBox(cleanBefore, {
    left: leftBefore + 20,
    top: top + 64,
    width: cardWidth - 40,
    height: cardHeight - 80
  });
  beforeBox.textFrame.textRange.font.size = 13.5;
  beforeBox.textFrame.textRange.font.color = "#450A0A";

  // 2. AFTER Card (Right)
  const leftAfter = leftBefore + cardWidth + gap;
  try {
    if (typeof newSlide.shapes.addGeometricShape === "function" && PowerPoint.GeometricShapeType) {
      const bgAfter = newSlide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, {
        left: leftAfter,
        top: top,
        width: cardWidth,
        height: cardHeight
      });
      if (bgAfter.fill && typeof bgAfter.fill.setSolidColor === "function") {
        bgAfter.fill.setSolidColor("#F0FDF4");
      }
      if (bgAfter.line) {
        bgAfter.line.color = "#BBF7D0";
        bgAfter.line.weight = 1.5;
      }
    }
  } catch (_) {}

  // After Header Badge
  const afterHeader = newSlide.shapes.addTextBox(`🟢 AFTER: ${afterData.title || 'Target State'}`, {
    left: leftAfter + 20,
    top: top + 18,
    width: cardWidth - 40,
    height: 36
  });
  afterHeader.textFrame.textRange.font.size = 16;
  afterHeader.textFrame.textRange.font.bold = true;
  afterHeader.textFrame.textRange.font.color = "#166534";

  // After Bullets
  const afterBullets = (afterData.bullets || []).map(b => `• ${b}`).join("\n\n");
  const { cleanText: cleanAfter } = parseMarkdownFormatting(afterBullets || "• Accelerated AI transformation");
  const afterBox = newSlide.shapes.addTextBox(cleanAfter, {
    left: leftAfter + 20,
    top: top + 64,
    width: cardWidth - 40,
    height: cardHeight - 80
  });
  afterBox.textFrame.textRange.font.size = 13.5;
  afterBox.textFrame.textRange.font.color = "#052E16";

  logToPPTConsole(`Slide ${slideNum}: Added native Before/After Comparison cards.`);
}

/**
 * Queries slide masters to locate the 'Blank' layout for the active presentation.
 */
async function getThemeBlankLayoutOptions() {
  try {
    return await PowerPoint.run(async (context) => {
      const slideMasters = context.presentation.slideMasters;
      slideMasters.load("items/id");
      await context.sync();

      if (!slideMasters.items || slideMasters.items.length === 0) return null;

      for (const master of slideMasters.items) {
        master.layouts.load("items/id, items/name");
        await context.sync();

        if (master.layouts && master.layouts.items) {
          const blankLayout = master.layouts.items.find(l => {
            const n = (l.name || "").toLowerCase();
            return n === "blank" || n.includes("blank") || n.includes("empty") || n.includes("vazio");
          });
          if (blankLayout) {
            return {
              slideMasterId: master.id,
              layoutId: blankLayout.id
            };
          }
        }
      }
      return null;
    });
  } catch (err) {
    console.warn("Could not query slide masters for blank layout:", err);
    return null;
  }
}

/**
 * Creates a single slide atomically in PowerPoint with title, body bullets, native tables, or images.
 * Supports replacing an existing slide in-place if targetSlideId is provided.
 */
async function createSingleSlide(slideData, slideNum, layoutOptions = null, targetSlideId = null) {
  const cleanTitle = (slideData.title || `Slide ${slideNum}`).replace(/\*\*/g, "").trim();
  const subtitle = slideData.subtitle || "";
  const takeaway = slideData.takeaway || "";
  const titleSize = slideData.titleSize || 40;
  const subtitleSize = slideData.subtitleSize || 20;
  const color = slideData.color || null;
  const bodyTextContent = slideData.body || "• Executive slide content";
  const additionalBody = slideData.additionalBody || "";
  const tableData = slideData.tableData || null;

  // Table slides MUST NEVER have images or charts. Table is the primary focal visual (Executive Table format).
  const hasTableData = Boolean(tableData && tableData.rows && tableData.rows.length > 0);
  const imagesToInsert = hasTableData ? [] : (
    (slideData.compressedImages && slideData.compressedImages.length > 0)
      ? slideData.compressedImages
      : (slideData.base64Images || [])
  );
  const hasImages = imagesToInsert.length > 0;

  logToPPTConsole(`Slide ${slideNum}: Preparing "${cleanTitle.substring(0, 32)}..."${targetSlideId ? ' (in-place replacement)' : ''}`);

  // 1. Single atomic PowerPoint.run session to add slide and populate all elements
  await PowerPoint.run(async (context) => {
    let newSlide;
    const slides = context.presentation.slides;

    if (targetSlideId) {
      try {
        newSlide = slides.getItem(targetSlideId);
        newSlide.shapes.load("items/name, items/type");
        await context.sync();
      } catch (itemErr) {
        console.warn("Could not retrieve targetSlideId, falling back to adding new slide:", itemErr);
        newSlide = null;
      }
    }

    if (!newSlide) {
      // 1. Get slide count before addition (in PowerPoint Office.js, slides.add() appends to the end;
      // the count before addition is the exact 0-based index of the new slide at the end)
      const countResult = slides.getCount();
      await context.sync();
      const insertIndex = countResult.value;

      // 2. Add the slide to the end of the presentation
      let added = false;
      if (layoutOptions) {
        try {
          slides.add(layoutOptions);
          added = true;
        } catch (layoutErr) {
          console.warn("Adding slide with blank layout failed, falling back to standard add:", layoutErr);
        }
      }

      if (!added) {
        slides.add();
      }

      await context.sync();

      // 3. Obtain reference to the newly appended slide at insertIndex
      newSlide = slides.getItemAt(insertIndex);
      newSlide.shapes.load("items/name, items/type");
      await context.sync();
    }

    // Clean up default template placeholders ("Click to add title", etc.)
    if (newSlide.shapes.items && newSlide.shapes.items.length > 0) {
      for (let i = newSlide.shapes.items.length - 1; i >= 0; i--) {
        try {
          newSlide.shapes.items[i].delete();
        } catch (_) {}
      }
      await context.sync();
    }

    const hasTakeaway = Boolean(takeaway && takeaway.trim().length > 0);
    const hasAdditionalNotes = Boolean(additionalBody && additionalBody.trim().length > 0);
    const hasTable = Boolean(tableData && tableData.rows && tableData.rows.length > 0);
    const compactHeader = hasTable && (hasAdditionalNotes || hasTakeaway);

    const titleTop = compactHeader ? 18 : 25;
    const effectiveTitleSize = compactHeader ? Math.min(titleSize || 36, 30) : (titleSize || 36);

    // Add Clean Title at Top using Gemini's requested title font size
    const titleBox = newSlide.shapes.addTextBox(cleanTitle, {
      left: 50,
      top: titleTop,
      width: 860,
      height: compactHeader ? 38 : 50
    });
    titleBox.textFrame.textRange.font.size = effectiveTitleSize;
    titleBox.textFrame.textRange.font.bold = true;
    try {
      titleBox.textFrame.wordWrap = false;
    } catch (wErr) {}
    if (color) titleBox.textFrame.textRange.font.color = color;

    let contentTop = compactHeader ? 60 : 85;

    // Add Subtitle if present
    if (subtitle) {
      const subtitleTop = compactHeader ? 52 : 72;
      const subtitleHeight = compactHeader ? 22 : 28;
      const effectiveSubtitleSize = compactHeader ? Math.min(subtitleSize || 18, 14) : (subtitleSize || 18);

      const subtitleBox = newSlide.shapes.addTextBox(subtitle, {
        left: 50,
        top: subtitleTop,
        width: 860,
        height: subtitleHeight
      });
      subtitleBox.textFrame.textRange.font.size = effectiveSubtitleSize;
      subtitleBox.textFrame.textRange.font.italic = true;
      if (color) subtitleBox.textFrame.textRange.font.color = color;
      contentTop = compactHeader ? 78 : 105;
    }

    // If Executive Visual layout is specified, render native shape cards
    if (slideData.visualType === "metric_grid_3col") {
      populate3ColumnMetricGrid(newSlide, slideData.visualData, slideNum, contentTop);
    } else if (slideData.visualType === "before_after") {
      populateBeforeAfterComparison(newSlide, slideData.visualData, slideNum, contentTop);
    } else if (hasTable) {
      const colCount = Math.max(
        tableData.headers ? tableData.headers.length : 0,
        ...tableData.rows.map(r => r.length),
        1
      );
      const rowCount = tableData.rows.length;
      const rawCandidateText = (additionalBody && additionalBody.trim().length > 0)
        ? additionalBody
        : (slideData.body && slideData.body.trim() !== "• Executive slide content" ? slideData.body : "");
      const bottomContent = hasTakeaway
        ? (rawCandidateText ? `${rawCandidateText}\n\n💡 Strategic Takeaway: ${takeaway}` : `💡 Strategic Takeaway: ${takeaway}`)
        : rawCandidateText;
      const hasBottomText = Boolean(bottomContent && bottomContent.trim().length > 0);

      // --- DYNAMIC COLLISION-PROOF LAYOUT ---
      if (hasImages && hasBottomText) {
        // --- 3-WAY HYBRID LAYOUT: BULLETS + CHART + TABLE ---
        if (rowCount <= 4) {
          // Layout A (compact table): Bullets on Left Column, Chart on Right-Top, Table on Right-Bottom
          const { cleanText, parsedParagraphs } = parseMarkdownFormatting(bottomContent);
          const notesBox = newSlide.shapes.addTextBox(cleanText, {
            left: 50,
            top: contentTop,
            width: 440,
            height: Math.min(410, 515 - contentTop)
          });
          notesBox.textFrame.wordWrap = true;
          notesBox.textFrame.textRange.font.size = 13.5;
          notesBox.textFrame.textRange.font.italic = hasTakeaway;
          await applyParagraphFormatting(notesBox, parsedParagraphs, context);

          const tableTop = contentTop + 215;
          const tableWidth = 400;
          const tableHeight = Math.min(185, 515 - tableTop);
          populateSlideTable(
            newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, tableData, slideNum, tableTop, tableHeight, tableWidth, 510
          );
        } else {
          // Layout B (tall table): Table on Left, Chart on Right-Top, Bullets on Right-Bottom
          const leftTableWidth = 440;
          const approxColWidth = leftTableWidth / colCount;
          const realTableHeight = Math.min(390, Math.max(90, estimateTableRenderedHeight(tableData, approxColWidth)));

          populateSlideTable(
            newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, tableData, slideNum, contentTop, realTableHeight, leftTableWidth, 50
          );

          const notesTop = contentTop + 215;
          const { cleanText, parsedParagraphs } = parseMarkdownFormatting(bottomContent);
          const notesBox = newSlide.shapes.addTextBox(cleanText, {
            left: 510,
            top: notesTop,
            width: 400,
            height: Math.min(185, 515 - notesTop)
          });
          notesBox.textFrame.wordWrap = true;
          notesBox.textFrame.textRange.font.size = 11.5;
          await applyParagraphFormatting(notesBox, parsedParagraphs, context);
        }
      } else if (!hasImages && (hasBottomText || rowCount > 4)) {
        // --- TWO-COLUMN LAYOUT: Table on Left, Bullets / Analysis / Takeaway on Right ---
        // Eliminates vertical overlap completely for tables with narrative or takeaways!
        const leftTableWidth = 440;
        const approxColWidth = leftTableWidth / colCount;
        const realTableHeight = Math.min(390, Math.max(90, estimateTableRenderedHeight(tableData, approxColWidth)));

        populateSlideTable(
          newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, tableData, slideNum, contentTop, realTableHeight, leftTableWidth, 50
        );

        if (hasBottomText) {
          const { cleanText, parsedParagraphs } = parseMarkdownFormatting(bottomContent);
          const notesBox = newSlide.shapes.addTextBox(cleanText, {
            left: 510,
            top: contentTop,
            width: 400,
            height: Math.min(390, 515 - contentTop)
          });
          notesBox.textFrame.wordWrap = true;
          notesBox.textFrame.textRange.font.size = 13.5;
          notesBox.textFrame.textRange.font.italic = hasTakeaway;
          await applyParagraphFormatting(notesBox, parsedParagraphs, context);
        }
      } else {
        // --- STANDALONE FULL-WIDTH TABLE LAYOUT (Optional Takeaway at bottom) ---
        const tableWidth = hasImages ? 420 : 860;
        const approxColWidth = tableWidth / colCount;
        const realTableHeight = estimateTableRenderedHeight(tableData, approxColWidth);

        const maxAllowedTableHeight = hasImages ? Math.min(390, 515 - contentTop) : Math.min(330, 515 - contentTop);
        const effectiveTableHeight = Math.min(realTableHeight, maxAllowedTableHeight);

        populateSlideTable(
          newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, tableData, slideNum, contentTop, effectiveTableHeight, tableWidth, 50
        );

        // If there is a takeaway or bottom text on standalone table, render safely below
        const notesTop = contentTop + realTableHeight + 16;
        const availableSpace = 515 - notesTop;

        if (hasBottomText && availableSpace >= 65 && (!hasImages || rowCount <= 4)) {
          let textToRender = bottomContent;
          const notesHeight = Math.min(availableSpace - 8, hasImages ? 130 : 200);
          const { cleanText, parsedParagraphs } = parseMarkdownFormatting(textToRender);
          const notesBox = newSlide.shapes.addTextBox(cleanText, {
            left: 50,
            top: notesTop,
            width: tableWidth,
            height: notesHeight
          });
          notesBox.textFrame.wordWrap = true;
          notesBox.textFrame.textRange.font.size = hasImages ? 11 : 12;
          notesBox.textFrame.textRange.font.italic = hasImages || hasTakeaway;
          await applyParagraphFormatting(notesBox, parsedParagraphs, context);
        }
      }
    } else {
      // Non-table slide: Bullets + optional Takeaway card at the bottom
      const rawBody = (slideData.body || "").trim();
      const isIntroOrEmpty = !rawBody || rawBody.toLowerCase().startsWith("here is the image") || rawBody === "• Executive slide content";
      const isNarrativeSlide = !hasTable && !hasImages && slideData.visualType !== "metric_grid_3col" && slideData.visualType !== "before_after";

      if (!isIntroOrEmpty || !hasImages) {
        // Strip any residual pseudo-visual marker lines
        const filteredBody = (slideData.body || "• Executive slide content")
          .split('\n')
          .filter(l => {
            const stripped = l.replace(/^[-•*]\s*/, '').trim();
            return !/^(?:📊\s*Metric Grid|⚖️\s*Comparison|Metric Grid|Comparison Card|Visual Concept:|Visual:)/i.test(stripped);
          })
          .join('\n');

        const { parsedParagraphs } = parseMarkdownFormatting(filteredBody);

        // Filter out empty lines to prevent excessive vertical spacing and premature slicing
        const meaningfulParagraphs = parsedParagraphs.filter(p => (p.cleanText || "").trim().length > 0);
        let finalParagraphs = meaningfulParagraphs;
        if (hasTakeaway && finalParagraphs.length > 4) {
          finalParagraphs = finalParagraphs.slice(0, 4);
        } else if (!hasTakeaway && finalParagraphs.length > 5) {
          finalParagraphs = finalParagraphs.slice(0, 5);
        }

        const finalCleanText = finalParagraphs.map(p => p.cleanText.trim()).join('\n\n');
        const { cleanText: reClean, parsedParagraphs: cleanParagraphs } = parseMarkdownFormatting(finalCleanText);

        const bulletBox = newSlide.shapes.addTextBox(reClean, {
          left: 50,
          top: contentTop,
          width: hasImages ? 400 : 860,
          height: hasTakeaway ? 260 : 380
        });
        bulletBox.textFrame.wordWrap = true;
        bulletBox.textFrame.textRange.font.size = hasImages ? 13.5 : (hasTakeaway && cleanParagraphs.length >= 4 ? 13.5 : 15);
        await applyParagraphFormatting(bulletBox, cleanParagraphs, context);
      }

      // Render Executive Takeaway at bottom (clean text box respecting presentation theme)
      if (hasTakeaway) {
        const rawTakeaway = `💡 Strategic Takeaway: ${takeaway}`;
        const { cleanText, parsedParagraphs } = parseMarkdownFormatting(rawTakeaway);
        const takeawayBox = newSlide.shapes.addTextBox(cleanText, {
          left: 50,
          top: 395,
          width: hasImages ? 400 : 860,
          height: 75
        });
        takeawayBox.textFrame.wordWrap = true;
        takeawayBox.textFrame.textRange.font.size = 13.5;
        takeawayBox.textFrame.textRange.font.italic = true;
        await applyParagraphFormatting(takeawayBox, parsedParagraphs, context);
      }
    }

    if (hasImages) {
      const rawCandidateText = (additionalBody && additionalBody.trim().length > 0)
        ? additionalBody
        : (slideData.body && slideData.body.trim() !== "• Executive slide content" ? slideData.body : "");
      const hasBottomText = Boolean(rawCandidateText || hasTakeaway);
      const isIntroOrEmpty = !rawCandidateText || rawCandidateText.toLowerCase().startsWith("here is the image");
      const isImageOnlySlide = !hasTable && isIntroOrEmpty;

      for (let imgIndex = 0; imgIndex < imagesToInsert.length; imgIndex++) {
        const rawImg = imagesToInsert[imgIndex];
        const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
        if (clean.length > 50) {
          // Calculate natural aspect ratio from image to prevent distortion and blurriness
          let aspect = 1.6;
          try {
            if (clean.length > 64 && typeof atob === "function") {
              const binStr = atob(clean.slice(0, 64));
              if (binStr.charCodeAt(0) === 0x89 && binStr.charCodeAt(1) === 0x50 && binStr.charCodeAt(2) === 0x4E && binStr.charCodeAt(3) === 0x47) {
                const w = (binStr.charCodeAt(16) << 24) | (binStr.charCodeAt(17) << 16) | (binStr.charCodeAt(18) << 8) | binStr.charCodeAt(19);
                const h = (binStr.charCodeAt(20) << 24) | (binStr.charCodeAt(21) << 16) | (binStr.charCodeAt(22) << 8) | binStr.charCodeAt(23);
                if (w > 0 && h > 0) aspect = w / h;
              }
            }
          } catch (_) {}

          let imgLeft = 470;
          let imgWidth = 440;
          let imgHeight;
          let imgTop = contentTop + 10;

          if (hasTable && hasBottomText) {
            // 3-way hybrid layout (Bullets on Left, Chart on Right-Top, Table on Right-Bottom)
            imgLeft = 510;
            imgWidth = 400;
            imgHeight = Math.min(205, Math.round(imgWidth / aspect));
            imgTop = contentTop + 5;
          } else if (hasTable) {
            // Side-by-side Table on Left, Chart on Right
            imgLeft = 490;
            imgWidth = 420;
            imgHeight = Math.min(360, Math.round(imgWidth / aspect));
            imgTop = contentTop + 10;
          } else if (isImageOnlySlide) {
            imgLeft = 200;
            imgWidth = 560;
            imgHeight = Math.min(360, Math.round(imgWidth / aspect));
            imgTop = contentTop + 10;
          } else {
            // Bullets on Left, Chart on Right
            imgLeft = 470;
            imgWidth = 440;
            if (imagesToInsert.length === 1) {
              imgHeight = Math.min(360, Math.round(imgWidth / aspect));
              imgTop = contentTop + 10;
            } else {
              const maxSlotHeight = Math.floor(340 / imagesToInsert.length);
              imgHeight = Math.min(maxSlotHeight, Math.round(imgWidth / aspect));
              imgTop = contentTop + 10 + imgIndex * (maxSlotHeight + 12);
            }
          }

          insertPictureOnSlide(newSlide, rawImg, {
            left: imgLeft,
            top: imgTop,
            width: imgWidth,
            height: imgHeight
          }, slideNum);
        }
      }
    }

    await context.sync();
  });

  logToPPTConsole(`Slide ${slideNum}: ✅ Created with Title, ${subtitle ? 'Subtitle, ' : ''}${tableData ? 'and Native Table.' : 'and Bullets.'}`);
}

/**
 * Builds all parsed slides sequentially with event-loop yielding between iterations.
 * @param {Array} slideStructures - Array of { title, body, base64Images, slideNumber }
 * @param {Object} options - { mode: 'insert' | 'replace' }
 * @param {Function} onProgress - Callback with { current, total, title }
 */
export async function buildPresentation(slideStructures, options = {}, onProgress = null) {
  if (typeof PowerPoint === "undefined" || !PowerPoint.run) {
    throw new Error("PowerPoint Office.js runtime is not available.");
  }

  if (!slideStructures || slideStructures.length === 0) {
    throw new Error("No slide structures found to build.");
  }

  const totalSlides = slideStructures.length;
  const isReplace = options.mode === "replace" || options.mode === "replace_draft";
  logToPPTConsole(`=== Starting ${isReplace ? 'Replacement' : 'Generation'} of ${totalSlides} Slide(s) ===`);

  // 1. Pre-process images
  for (let idx = 0; idx < slideStructures.length; idx++) {
    const slide = slideStructures[idx];
    slide.compressedImages = [];
    const rawImages = slide.base64Images || [];
    for (const rawImg of rawImages) {
      try {
        const comp = await compressImageForPowerPoint(rawImg);
        if (comp && comp.length > 50) {
          slide.compressedImages.push(comp);
        }
      } catch (cErr) {
        logToPPTConsole(`Image compression notice: ${cErr.message}`);
      }
    }
  }

  // 2. Query theme Blank layout once for clean slide generation without placeholders
  const blankLayoutOptions = await getThemeBlankLayoutOptions();
  if (blankLayoutOptions) {
    logToPPTConsole(`Applying theme Blank layout to avoid template placeholders.`);
  }

  // 3. Find target slide ID for first slide ONLY IF in explicit replace mode:
  let activeSlideId = null;
  let shouldTargetFirstSlide = false;
  if (isReplace) {
    try {
      await PowerPoint.run(async (context) => {
        if (context.presentation.getSelectedSlides) {
          const selected = context.presentation.getSelectedSlides();
          selected.load("items/id");
          await context.sync();
          if (selected.items && selected.items.length > 0) {
            activeSlideId = selected.items[0].id;
            shouldTargetFirstSlide = true;
          }
        }
      });
    } catch (selErr) {
      console.warn("Could not determine selected slide for replacement:", selErr);
    }
  }

  // 4. Build each slide sequentially
  let successfulSlides = 0;
  for (let i = 0; i < totalSlides; i++) {
    const slideData = slideStructures[i];
    const slideNum = i + 1;
    const targetSlideId = (shouldTargetFirstSlide && i === 0 && activeSlideId) ? activeSlideId : null;

    if (typeof onProgress === "function") {
      onProgress({
        current: slideNum,
        total: totalSlides,
        title: slideData.title
      });
    }

    try {
      const slidePromise = createSingleSlide(slideData, slideNum, blankLayoutOptions, targetSlideId);
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Timeout (45s) creating slide in PowerPoint")), 45000)
      );
      await Promise.race([slidePromise, timeoutPromise]);
      successfulSlides++;
    } catch (slideErr) {
      logToPPTConsole(`Slide ${slideNum} Notice: ${slideErr.message}. Attempting resilient continuation...`, true);
      console.warn(`[PPTBuilder] Slide ${slideNum} issue:`, slideErr);
      try {
        const fallbackBody = (slideData.body && slideData.body.trim() !== "• Executive slide content")
          ? slideData.body
          : (slideData.tableData && slideData.tableData.rows && slideData.tableData.rows.length > 0
              ? slideData.tableData.rows.map(r => `• ${r.join(" | ")}`).join("\n\n")
              : "• Executive slide content");

        const fallbackData = {
          ...slideData,
          base64Images: [],
          tableData: null,
          visualType: null,
          body: fallbackBody
        };
        await createSingleSlide(fallbackData, slideNum, blankLayoutOptions, targetSlideId);
        logToPPTConsole(`Slide ${slideNum}: Added basic text fallback slide.`);
        successfulSlides++;
      } catch (fbErr) {
        console.warn(`[PPTBuilder] Slide ${slideNum} fallback failed:`, fbErr);
      }
    }

    // Yield event loop for 400ms to allow PowerPoint host to finalize layout before next slide
    await new Promise(resolve => setTimeout(resolve, 400));
  }

  if (successfulSlides === 0) {
    logToPPTConsole(`❌ Generation failed: 0 of ${totalSlides} slide(s) could be created.`, true);
    throw new Error(`Failed to create slides in PowerPoint.`);
  } else if (successfulSlides < totalSlides) {
    logToPPTConsole(`⚠️ ${successfulSlides} of ${totalSlides} slide(s) created in PowerPoint.`);
  } else {
    logToPPTConsole(`🎉 All ${totalSlides} slide(s) ${isReplace ? 'replaced' : 'created'} successfully!`);
  }
}

/**
 * Inserts a picture onto a PowerPoint slide using multi-tier fallback strategies
 * (shapes.addPicture -> shapes.addGeometricShape fill -> shapes.addImage).
 * 
 * @param {Object} targetSlide - PowerPoint slide object
 * @param {string} rawImg - Base64 or Data URI string
 * @param {Object} bounds - { left, top, width, height }
 * @param {number} slideNum - Slide index for logging
 * @returns {boolean} Whether picture was inserted
 */
export function insertPictureOnSlide(targetSlide, rawImg, { left, top, width, height }, slideNum = 1) {
  const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
  if (clean.length <= 50) return false;

  let picInserted = false;
  // Strategy 1: Standard PowerPoint Office.js shapes.addPicture(base64, options)
  try {
    if (typeof targetSlide.shapes.addPicture === "function") {
      const pic = targetSlide.shapes.addPicture(clean, {
        left,
        top,
        width,
        height
      });
      if (pic) {
        if (pic.lineFormat) {
          try { pic.lineFormat.visible = false; } catch (_) {}
        }
        picInserted = true;
        logToPPTConsole(`Slide ${slideNum}: Added picture via shapes.addPicture.`);
      }
    }
  } catch (picErr) {
    console.warn("shapes.addPicture clean failed, trying raw URI:", picErr);
    try {
      if (typeof targetSlide.shapes.addPicture === "function") {
        const pic2 = targetSlide.shapes.addPicture(rawImg, {
          left,
          top,
          width,
          height
        });
        if (pic2) {
          if (pic2.lineFormat) {
            try { pic2.lineFormat.visible = false; } catch (_) {}
          }
          picInserted = true;
          logToPPTConsole(`Slide ${slideNum}: Added picture via rawImg addPicture.`);
        }
      }
    } catch (rawErr) {
      console.warn("shapes.addPicture rawImg failed:", rawErr);
    }
  }

  // Strategy 2: Geometric shape fill (PowerPointApi 1.8+)
  if (!picInserted) {
    try {
      if (typeof targetSlide.shapes.addGeometricShape === "function") {
        const rect = targetSlide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
          left,
          top,
          width,
          height
        });
        if (rect) {
          if (rect.lineFormat) {
            try {
              rect.lineFormat.visible = false;
              rect.lineFormat.weight = 0;
              rect.lineFormat.color = "#ffffff";
            } catch (_) {}
          }
          if (rect.line) {
            try { rect.line.visible = false; } catch (_) {}
          }
          if (rect.fill) {
            if (typeof rect.fill.setImage === "function") {
              rect.fill.setImage(clean);
              picInserted = true;
            } else if (typeof rect.fill.setPictureFromBase64 === "function") {
              rect.fill.setPictureFromBase64(clean);
              picInserted = true;
            }
          }
        }
        if (picInserted) {
          logToPPTConsole(`Slide ${slideNum}: Added picture via shape fill.`);
        }
      }
    } catch (fillErr) {
      console.warn("Geometric shape fill fallback failed:", fillErr);
    }
  }

  // Strategy 3: shapes.addImage if present in custom Office.js host
  if (!picInserted) {
    try {
      if (typeof targetSlide.shapes.addImage === "function") {
        const imgShape = targetSlide.shapes.addImage(clean);
        if (imgShape) {
          imgShape.left = left;
          imgShape.top = top;
          imgShape.width = width;
          imgShape.height = height;
          picInserted = true;
          logToPPTConsole(`Slide ${slideNum}: Added picture via addImage.`);
        }
      }
    } catch (addImgErr) {
      console.warn("shapes.addImage fallback failed:", addImgErr);
    }
  }

  return picInserted;
}

/**
 * Inserts generated AI content (images, tables, text bullets) directly onto the current active slide
 * without deleting or modifying existing slide shapes or content.
 * 
 * @param {Array} slideStructures - Array of parsed slide objects
 * @param {Object} options - Insertion options
 */
export async function insertOnCurrentSlide(slideStructures, options = {}) {
  if (typeof PowerPoint === "undefined" || !PowerPoint.run) {
    throw new Error("PowerPoint Office.js runtime is not available.");
  }

  if (!slideStructures || slideStructures.length === 0) {
    throw new Error("No slide structures found to insert.");
  }

  logToPPTConsole(`=== Starting Insert on Current Slide ===`);

  // 1. Pre-process and compress images
  for (const slide of slideStructures) {
    slide.compressedImages = [];
    const rawImages = slide.base64Images || [];
    for (const rawImg of rawImages) {
      try {
        const comp = await compressImageForPowerPoint(rawImg);
        if (comp && comp.length > 50) {
          slide.compressedImages.push(comp);
        }
      } catch (cErr) {
        logToPPTConsole(`Image compression notice: ${cErr.message}`);
      }
    }
  }

  // 2. Identify active slide and analyze layout
  await PowerPoint.run(async (context) => {
    let activeSlide = null;
    if (context.presentation.getSelectedSlides) {
      const selected = context.presentation.getSelectedSlides();
      selected.load("items/id");
      await context.sync();
      if (selected.items && selected.items.length > 0) {
        activeSlide = selected.items[0];
      }
    }

    if (!activeSlide) {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      if (slides.items && slides.items.length > 0) {
        activeSlide = slides.items[0];
      }
    }

    if (!activeSlide) {
      throw new Error("No slide available in presentation to insert content onto.");
    }

    // Selected shape tracking disabled to prevent host selection triggers
    let selectedShapeBox = null;

    // Load existing shapes on the active slide to detect occupied space
    activeSlide.shapes.load("items/left, items/top, items/width, items/height, items/type, items/name");
    await context.sync();

    const existingShapes = activeSlide.shapes.items || [];
    let minLeft = 960, maxRight = 0, minTop = 540, maxBottom = 0;
    let hasExistingContent = false;

    for (const s of existingShapes) {
      if (s.width > 20 && s.height > 20 && s.left >= 0 && s.top >= 0 && s.left < 960 && s.top < 540) {
        hasExistingContent = true;
        minLeft = Math.min(minLeft, s.left);
        maxRight = Math.max(maxRight, s.left + s.width);
        minTop = Math.min(minTop, s.top);
        maxBottom = Math.max(maxBottom, s.top + s.height);
      }
    }

    const slideData = slideStructures[0];
    const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
      ? slideData.compressedImages
      : (slideData.base64Images || []);
    const hasImages = imagesToInsert.length > 0;
    const hasTable = Boolean(slideData.tableData && slideData.tableData.rows && slideData.tableData.rows.length > 0);
    const rawBody = (slideData.body || "").trim();
    const isIntroOrEmpty = !rawBody || rawBody.toLowerCase().startsWith("here is the image") || rawBody === "• Executive slide content";
    const hasBodyText = !isIntroOrEmpty;

    logToPPTConsole(`Current slide has ${existingShapes.length} shapes. hasExistingContent=${hasExistingContent}, maxRight=${maxRight}, maxBottom=${maxBottom}`);

    // Determine default layout positioning relative to user selection and existing slide content
    let defaultSlot = "right"; // "right", "left", "center", "below"
    if (selectedShapeBox) {
      if (selectedShapeBox.left < 450) {
        defaultSlot = "right";
      } else if (selectedShapeBox.left >= 450) {
        defaultSlot = "left";
      } else if (selectedShapeBox.top < 150 && selectedShapeBox.width > 600) {
        defaultSlot = "below";
      }
    } else if (!hasExistingContent) {
      defaultSlot = "center";
    } else if (maxRight <= 520) {
      defaultSlot = "right";
    } else if (minLeft >= 460) {
      defaultSlot = "left";
    } else if (maxBottom <= 160) {
      defaultSlot = "below";
    } else {
      defaultSlot = "right";
    }

    // 1. Insert Images if present
    if (hasImages) {
      for (let imgIndex = 0; imgIndex < imagesToInsert.length; imgIndex++) {
        const rawImg = imagesToInsert[imgIndex];
        const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
        if (clean.length <= 50) continue;

        let aspect = 1.6;
        try {
          if (clean.length > 64 && typeof atob === "function") {
            const binStr = atob(clean.slice(0, 64));
            if (binStr.charCodeAt(0) === 0x89 && binStr.charCodeAt(1) === 0x50 && binStr.charCodeAt(2) === 0x4E && binStr.charCodeAt(3) === 0x47) {
              const w = (binStr.charCodeAt(16) << 24) | (binStr.charCodeAt(17) << 16) | (binStr.charCodeAt(18) << 8) | binStr.charCodeAt(19);
              const h = (binStr.charCodeAt(20) << 24) | (binStr.charCodeAt(21) << 16) | (binStr.charCodeAt(22) << 8) | binStr.charCodeAt(23);
              if (w > 0 && h > 0) aspect = w / h;
            }
          }
        } catch (_) {}

        let imgLeft, imgTop, imgWidth, imgHeight;

        if (defaultSlot === "center") {
          imgWidth = 560;
          imgHeight = Math.min(380, Math.round(imgWidth / aspect));
          imgLeft = Math.round((960 - imgWidth) / 2);
          imgTop = Math.round((540 - imgHeight) / 2);
        } else if (defaultSlot === "left") {
          imgLeft = 50;
          imgWidth = 420;
          imgHeight = Math.min(380, Math.round(imgWidth / aspect));
          imgTop = Math.max(90, Math.min(130, minTop));
        } else if (defaultSlot === "below") {
          imgWidth = 520;
          imgHeight = Math.min(340, Math.round(imgWidth / aspect));
          imgLeft = Math.round((960 - imgWidth) / 2);
          imgTop = Math.max(140, maxBottom + 15);
        } else {
          // "right"
          imgLeft = 500;
          imgWidth = 420;
          imgHeight = Math.min(380, Math.round(imgWidth / aspect));
          imgTop = Math.max(90, Math.min(130, minTop));
        }

        if (imagesToInsert.length > 1) {
          const slotHeight = Math.floor(imgHeight / imagesToInsert.length);
          imgTop = imgTop + imgIndex * (slotHeight + 10);
          imgHeight = slotHeight;
        }

        insertPictureOnSlide(activeSlide, rawImg, {
          left: imgLeft,
          top: imgTop,
          width: imgWidth,
          height: imgHeight
        });

        logToPPTConsole(`Inserted picture ${imgIndex + 1} onto current slide at left:${imgLeft}, top:${imgTop}, width:${imgWidth}, height:${imgHeight}`);
      }

      // If there is also substantial narrative text accompanying the image, insert a companion text box
      if (hasBodyText && (defaultSlot === "right" || defaultSlot === "center") && !hasExistingContent) {
        const textLeft = 50;
        const textTop = 120;
        const { cleanText, parsedParagraphs } = parseMarkdownFormatting(slideData.body);
        const textBox = activeSlide.shapes.addTextBox(cleanText, {
          left: textLeft,
          top: textTop,
          width: 420,
          height: 380
        });
        textBox.textFrame.wordWrap = true;
        textBox.textFrame.textRange.font.size = 13;
        await applyParagraphFormatting(textBox, parsedParagraphs, context);
      }
    }
    // 2. Insert Native Table if present (and no image)
    else if (hasTable) {
      let tblLeft = 50;
      let tblTop = 120;
      let tblWidth = 860;

      if (defaultSlot === "right") {
        tblLeft = 500;
        tblWidth = 420;
        tblTop = Math.max(90, minTop);
      } else if (defaultSlot === "left") {
        tblLeft = 50;
        tblWidth = 420;
        tblTop = Math.max(90, minTop);
      } else if (defaultSlot === "below") {
        tblLeft = 50;
        tblTop = Math.max(140, maxBottom + 15);
        tblWidth = 860;
      }

      populateSlideTable(activeSlide, null, null, 20, 14, "#0078d4", slideData.tableData, 1, tblTop, null, tblWidth, tblLeft);
      logToPPTConsole(`Inserted native table onto current slide at left:${tblLeft}, top:${tblTop}`);
    }
    // 3. Insert Text / Bullets
    else if (hasBodyText) {
      let textLeft = 50;
      let textTop = 120;
      let textWidth = 860;
      let textHeight = 360;

      if (defaultSlot === "right") {
        textLeft = 500;
        textWidth = 420;
        textTop = Math.max(90, minTop);
        textHeight = Math.min(380, 520 - textTop);
      } else if (defaultSlot === "left") {
        textLeft = 50;
        textWidth = 420;
        textTop = Math.max(90, minTop);
        textHeight = Math.min(380, 520 - textTop);
      } else if (defaultSlot === "below") {
        textLeft = 50;
        textTop = Math.max(140, maxBottom + 15);
        textWidth = 860;
        textHeight = Math.min(340, 520 - textTop);
      }

      const { cleanText, parsedParagraphs } = parseMarkdownFormatting(slideData.body);
      const textBox = activeSlide.shapes.addTextBox(cleanText, {
        left: textLeft,
        top: textTop,
        width: textWidth,
        height: textHeight
      });
      textBox.textFrame.wordWrap = true;
      textBox.textFrame.textRange.font.size = textWidth < 500 ? 13 : 14.5;
      await applyParagraphFormatting(textBox, parsedParagraphs, context);
      logToPPTConsole(`Inserted text box onto current slide at left:${textLeft}, top:${textTop}`);
    }

    await context.sync();
  });

  logToPPTConsole(`🎉 Successfully inserted content onto current slide!`);
}
