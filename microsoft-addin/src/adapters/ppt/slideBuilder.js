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
export function compressImageForPowerPoint(base64Str, maxWidth = 800, maxHeight = 550) {
  return new Promise((resolve) => {
    let resolved = false;
    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true;
        resolve(base64Str ? base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim() : "");
      }
    }, 1000);

    try {
      const cleanRaw = base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
      if (!cleanRaw) {
        clearTimeout(timer);
        resolve("");
        return;
      }
      const img = new Image();
      img.onload = () => {
        if (resolved) return;
        resolved = true;
        clearTimeout(timer);
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
        resolve(base64Str.replace(/^data:image\/[^;]+;base64,/, "").replace(/[\r\n\s]+/g, "").trim());
      }
    }
  });
}


/**
 * Accurately estimates rendered height of a native PowerPoint table based on word-wrapping.
 */
function estimateTableRenderedHeight(tableData, colWidth = 280) {
  if (!tableData) return 0;
  const headers = tableData.headers || [];
  const rows = tableData.rows || [];
  let totalHeight = headers.length > 0 ? 36 : 0; // Header row height

  for (const row of rows) {
    // Determine maximum length among cell values in this row
    const maxChars = Math.max(...row.map(c => String(c !== undefined && c !== null ? c : "").trim().length), 0);
    // At ~colWidth, approx colWidth / 8.5 characters fit per line
    const charsPerLine = Math.max(12, Math.floor(colWidth / 8.5));
    const approxLines = Math.max(1, Math.ceil(maxChars / charsPerLine));

    // Cell padding (14pt) + line spacing (~16pt per line)
    const rowHeight = Math.max(28, approxLines * 16 + 14);
    totalHeight += rowHeight;
  }
  return totalHeight;
}

/**
 * Populates a native Microsoft PowerPoint table using PowerPoint.js shapes.addTable().
 */
function populateSlideTable(newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, tableData, slideNum, tableTop = 90, customHeight = null, customWidth = null) {
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
  const tableHeight = customHeight || Math.min(380, Math.max(90, estimatedHeight));

  let addedShape = null;
  try {
    if (typeof newSlide.shapes.addTable === "function") {
      addedShape = newSlide.shapes.addTable(rowCount, colCount, {
        left: 50,
        top: tableTop,
        width: tableWidth,
        height: tableHeight,
        values: tableValues
      });
      logToPPTConsole(`Slide ${slideNum}: Added native PowerPoint table (${rowCount} rows x ${colCount} cols, rendered height ~${estimatedHeight}pt).`);
    }
  } catch (err) {
    console.warn("shapes.addTable with options failed, trying basic addTable:", err);
  }

  if (!addedShape) {
    try {
      addedShape = newSlide.shapes.addTable(rowCount, colCount);
      const table = addedShape.getTable();
      for (let r = 0; r < tableValues.length; r++) {
        for (let c = 0; c < colCount; c++) {
          const cell = table.getCellOrNullObject(r, c);
          if (cell) cell.text = tableValues[r][c];
        }
      }
      logToPPTConsole(`Slide ${slideNum}: Added native PowerPoint table via getCell.`);
    } catch (fallbackErr) {
      console.error("Native table shape creation failed:", fallbackErr);
      return estimatedHeight;
    }
  }

  // Format table font size to 11-12pt so cells fit cleanly without extreme wrapping
  if (addedShape && typeof addedShape.getTable === "function") {
    try {
      const table = addedShape.getTable();
      for (let r = 0; r < tableValues.length; r++) {
        for (let c = 0; c < colCount; c++) {
          const cell = table.getCellOrNullObject(r, c);
          if (cell && cell.textFrame && cell.textFrame.textRange) {
            cell.textFrame.textRange.font.size = r === 0 ? 12 : 11;
            if (r === 0) cell.textFrame.textRange.font.bold = true;
          }
        }
      }
    } catch (_) {}
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
      const bulletsBox = newSlide.shapes.addTextBox(bulletContent, {
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
  const beforeBox = newSlide.shapes.addTextBox(beforeBullets || "• Legacy workflow bottlenecks", {
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
  const afterBox = newSlide.shapes.addTextBox(afterBullets || "• Accelerated AI transformation", {
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

  const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
    ? slideData.compressedImages
    : (slideData.base64Images || []);
  const hasImages = imagesToInsert.length > 0;

  logToPPTConsole(`Slide ${slideNum}: Preparing "${cleanTitle.substring(0, 32)}..."${targetSlideId ? ' (in-place replacement)' : ''}`);

  // 1. If not replacing an existing slide, add slide using Theme Blank layout if available
  if (!targetSlideId) {
    let added = false;
    if (layoutOptions) {
      try {
        await PowerPoint.run(async (context) => {
          context.presentation.slides.add(layoutOptions);
          await context.sync();
          added = true;
        });
      } catch (layoutErr) {
        console.warn("Adding slide with blank layout failed, falling back to standard add:", layoutErr);
      }
    }

    if (!added) {
      await PowerPoint.run(async (context) => {
        context.presentation.slides.add();
        await context.sync();
      });
    }
  }

  // 2. Eliminate template placeholders ("Click to add title", "Click to add subtitle") or previous shapes
  try {
    await PowerPoint.run(async (context) => {
      let targetSlide;
      if (targetSlideId) {
        targetSlide = context.presentation.slides.getItem(targetSlideId);
      } else {
        const slides = context.presentation.slides;
        const countResult = slides.getCount();
        await context.sync();
        targetSlide = slides.getItemAt(countResult.value - 1);
      }

      targetSlide.shapes.load("items/name, items/type");
      await context.sync();

      if (targetSlide.shapes.items && targetSlide.shapes.items.length > 0) {
        for (let i = targetSlide.shapes.items.length - 1; i >= 0; i--) {
          const s = targetSlide.shapes.items[i];
          try {
            s.delete();
          } catch (_) {}
        }
        await context.sync();
      }
    });
  } catch (cleanErr) {
    // Fallback: If shape deletion is blocked by PowerPoint host, neutralize by moving off-canvas and clearing text
    try {
      await PowerPoint.run(async (context) => {
        let targetSlide;
        if (targetSlideId) {
          targetSlide = context.presentation.slides.getItem(targetSlideId);
        } else {
          const slides = context.presentation.slides;
          const countResult = slides.getCount();
          await context.sync();
          targetSlide = slides.getItemAt(countResult.value - 1);
        }

        targetSlide.shapes.load("items/name, items/type");
        await context.sync();

        if (targetSlide.shapes.items) {
          for (const s of targetSlide.shapes.items) {
            try {
              s.textFrame.textRange.text = " ";
            } catch (_) {}
            try {
              s.left = -5000;
              s.top = -5000;
            } catch (_) {}
          }
          await context.sync();
        }
      });
    } catch (_) {}
  }

  // 3. Populate slide content in a fresh, uncorrupted PowerPoint.run
  await PowerPoint.run(async (context) => {
    let newSlide;
    if (targetSlideId) {
      newSlide = context.presentation.slides.getItem(targetSlideId);
    } else {
      const slides = context.presentation.slides;
      const countResult = slides.getCount();
      await context.sync();
      newSlide = slides.getItemAt(countResult.value - 1);
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
      const tableWidth = hasImages ? 400 : 860;
      const approxColWidth = tableWidth / colCount;
      const realTableHeight = estimateTableRenderedHeight(tableData, approxColWidth);

      populateSlideTable(
        newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, tableData, slideNum, contentTop, realTableHeight, tableWidth
      );

      // Render takeaway or additional notes below the table without overlapping
      const bottomContent = hasTakeaway ? `💡 Strategic Takeaway: ${takeaway}` : additionalBody;
      if (bottomContent && bottomContent.trim().length > 0) {
        const notesTop = Math.max(contentTop + realTableHeight + 16, 320);
        const notesHeight = Math.max(45, Math.min(140, 520 - notesTop));
        const notesBox = newSlide.shapes.addTextBox(bottomContent, {
          left: 50,
          top: notesTop,
          width: tableWidth,
          height: notesHeight
        });
        notesBox.textFrame.wordWrap = true;
        notesBox.textFrame.textRange.font.size = realTableHeight > 200 ? 11.5 : 12.5;
        notesBox.textFrame.textRange.font.italic = true;
        try {
          if (hasTakeaway) {
            notesBox.textFrame.textRange.getSubstring(0, 22).font.bold = true;
          }
        } catch (_) {}
      }
    } else {
      // Non-table slide: Bullets + optional Takeaway card at the bottom
      const bodyHeight = hasTakeaway ? 280 : 380;
      const bodyBox = newSlide.shapes.addTextBox(bodyTextContent, {
        left: 50,
        top: contentTop,
        width: hasImages ? 400 : 860,
        height: bodyHeight
      });
      bodyBox.textFrame.textRange.font.size = 16;

      // Format bullet points with bold lead-ins for key points before colons or dashes
      try {
        const paragraphs = bodyBox.textFrame.textRange.paragraphs;
        paragraphs.load("items/text");
        await context.sync();
        if (paragraphs.items) {
          for (const p of paragraphs.items) {
            const pText = p.text || "";
            const colonIdx = pText.indexOf(":");
            const dashIdx = pText.indexOf("—");
            const sepIdx = colonIdx > 0 ? colonIdx : (dashIdx > 0 ? dashIdx : -1);
            if (sepIdx > 0 && sepIdx < 50 && typeof p.getSubstring === "function") {
              try {
                const leadIn = p.getSubstring(0, sepIdx + 1);
                leadIn.font.bold = true;
              } catch (_) {}
            }
          }
        }
      } catch (boldErr) {
        console.warn("Lead-in bolding notice:", boldErr);
      }

      // Render dedicated Executive Takeaway Callout Box at bottom
      if (hasTakeaway) {
        const takeawayBox = newSlide.shapes.addTextBox(`💡 Strategic Takeaway: ${takeaway}`, {
          left: 50,
          top: 395,
          width: hasImages ? 400 : 860,
          height: 75
        });
        takeawayBox.textFrame.textRange.font.size = 14;
        takeawayBox.textFrame.textRange.font.italic = true;
        try {
          takeawayBox.textFrame.textRange.getSubstring(0, 22).font.bold = true;
        } catch (_) {}
      }
    }

    if (hasImages) {
      for (const rawImg of imagesToInsert) {
        const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
        if (clean.length > 50) {
          try {
            const imgShape = newSlide.shapes.addImage(clean);
            imgShape.left = 460;
            imgShape.top = contentTop;
            imgShape.width = 440;
            imgShape.height = 330;
            logToPPTConsole(`Slide ${slideNum}: Added chart image shape.`);
          } catch (imgErr) {
            console.warn("shapes.addImage failed with clean base64, trying raw:", imgErr);
            try {
              const imgShape2 = newSlide.shapes.addImage(rawImg);
              imgShape2.left = 460;
              imgShape2.top = contentTop;
              imgShape2.width = 440;
              imgShape2.height = 330;
            } catch (fallbackErr) {
              console.error("shapes.addImage failed:", fallbackErr);
            }
          }
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

  // 3. If in replace mode, find currently selected slide ID so first slide replaces in-place
  let activeSlideId = null;
  if (isReplace) {
    try {
      await PowerPoint.run(async (context) => {
        if (context.presentation.getSelectedSlides) {
          const selected = context.presentation.getSelectedSlides();
          selected.load("items/id");
          await context.sync();
          if (selected.items && selected.items.length > 0) {
            activeSlideId = selected.items[0].id;
          }
        }
      });
    } catch (selErr) {
      console.warn("Could not determine selected slide for replace mode:", selErr);
    }
  }

  // 4. Build each slide sequentially
  for (let i = 0; i < totalSlides; i++) {
    const slideData = slideStructures[i];
    const slideNum = i + 1;
    const targetSlideId = (isReplace && i === 0 && activeSlideId) ? activeSlideId : null;

    if (typeof onProgress === "function") {
      onProgress({
        current: slideNum,
        total: totalSlides,
        title: slideData.title
      });
    }

    try {
      await createSingleSlide(slideData, slideNum, blankLayoutOptions, targetSlideId);
    } catch (slideErr) {
      logToPPTConsole(`Slide ${slideNum} Error: ${slideErr.message}`, true);
      console.error(`[PPTBuilder] Slide ${slideNum} Error:`, slideErr);
      throw slideErr;
    }

    // Yield event loop for 400ms to allow PowerPoint host to finalize layout before next slide
    await new Promise(resolve => setTimeout(resolve, 400));
  }

  logToPPTConsole(`🎉 All ${totalSlides} slide(s) ${isReplace ? 'replaced' : 'created'} successfully!`);
}
