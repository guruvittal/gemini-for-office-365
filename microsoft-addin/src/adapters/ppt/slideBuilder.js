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
 * Discovers the Blank layout (or best clean layout) from the presentation's active theme/master.
 */
async function getThemeBlankLayoutOptions() {
  try {
    return await PowerPoint.run(async (context) => {
      const slideMasters = context.presentation.slideMasters;
      slideMasters.load("id, name, layouts/items/name, layouts/items/id");
      await context.sync();

      if (!slideMasters.items || slideMasters.items.length === 0) {
        return null;
      }

      // Use the active slide master
      const master = slideMasters.items[0];
      if (!master.layouts || !master.layouts.items || master.layouts.items.length === 0) {
        return null;
      }

      // 1. Look for a layout named "blank" (case-insensitive)
      let targetLayout = master.layouts.items.find(l => (l.name || "").toLowerCase().includes("blank"));

      // 2. If no "blank", look for "empty" or "clean"
      if (!targetLayout) {
        targetLayout = master.layouts.items.find(l => {
          const n = (l.name || "").toLowerCase();
          return n.includes("empty") || n.includes("clean") || n.includes("custom");
        });
      }

      if (targetLayout) {
        return {
          slideMasterId: master.id,
          layoutId: targetLayout.id
        };
      }

      return null;
    });
  } catch (err) {
    console.warn("Could not query slide masters/layouts:", err);
    return null;
  }
}

/**
 * Populates a native Microsoft PowerPoint table using PowerPoint.js shapes.addTable().
 */
function populateSlideTable(newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, tableData, slideNum, tableTop = 90) {
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

  const tableHeight = Math.min(400, Math.max(100, rowCount * 32));

  try {
    if (typeof newSlide.shapes.addTable === "function") {
      newSlide.shapes.addTable(rowCount, colCount, {
        left: 50,
        top: tableTop,
        width: 860,
        height: tableHeight,
        values: tableValues
      });
      logToPPTConsole(`Slide ${slideNum}: Added native PowerPoint table (${rowCount} rows x ${colCount} cols).`);
      return;
    }
  } catch (err) {
    console.warn("shapes.addTable with options failed, trying basic addTable:", err);
  }

  try {
    const shape = newSlide.shapes.addTable(rowCount, colCount);
    const table = shape.getTable();
    for (let r = 0; r < tableValues.length; r++) {
      for (let c = 0; c < colCount; c++) {
        const cell = table.getCellOrNullObject(r, c);
        if (cell) cell.text = tableValues[r][c];
      }
    }
    logToPPTConsole(`Slide ${slideNum}: Added native PowerPoint table via getCell.`);
  } catch (fallbackErr) {
    console.error("Native table shape creation failed:", fallbackErr);
    logToPPTConsole(`Slide ${slideNum}: ⚠️ Table shape notice: ${fallbackErr.message}`);
  }
}

/**
 * Creates a single slide atomically in PowerPoint with title, body bullets, native tables, or images.
 */
async function createSingleSlide(slideData, slideNum, layoutOptions = null) {
  let cleanTitle = (slideData.title || `Slide ${slideNum}`).replace(/\*\*/g, "").trim();
  if (/^(?:here\s+(?:is|are)\s+(?:the\s+)?slide|the\s+slide|slide)$/i.test(cleanTitle)) {
    cleanTitle = `Executive Briefing`;
  }
  const subtitle = slideData.subtitle || "";
  const titleSize = slideData.titleSize || 40;
  const subtitleSize = slideData.subtitleSize || 20;
  const color = slideData.color || null;
  const bodyTextContent = slideData.body || "• Executive slide content";
  const tableData = slideData.tableData || null;

  const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
    ? slideData.compressedImages
    : (slideData.base64Images || []);
  const hasImages = imagesToInsert.length > 0;

  logToPPTConsole(`Slide ${slideNum}: Preparing "${cleanTitle.substring(0, 32)}..."`);

  // Atomic PowerPoint slide creation with clean shape management
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    if (layoutOptions) {
      try {
        slides.add(layoutOptions);
      } catch (lErr) {
        slides.add();
      }
    } else {
      slides.add();
    }
    await context.sync();

    slides.load("items");
    await context.sync();

    const newSlide = slides.items[slides.items.length - 1];
    newSlide.shapes.load("items");
    await context.sync();

    // Delete any default template placeholders (e.g. "Click to add title", "Click to add text")
    if (newSlide.shapes.items && newSlide.shapes.items.length > 0) {
      for (let i = newSlide.shapes.items.length - 1; i >= 0; i--) {
        try {
          newSlide.shapes.items[i].delete();
        } catch (dErr) {}
      }
      await context.sync();
    }

    // Add Clean Title at Top using Gemini's requested title font size
    const titleBox = newSlide.shapes.addTextBox(cleanTitle, {
      left: 50,
      top: 25,
      width: 860,
      height: 50
    });
    titleBox.textFrame.textRange.font.size = titleSize;
    titleBox.textFrame.textRange.font.bold = true;
    try {
      titleBox.textFrame.wordWrap = false;
    } catch (wErr) {}
    if (color) titleBox.textFrame.textRange.font.color = color;

    let contentTop = 85;

    // Add Subtitle if present
    if (subtitle) {
      const subtitleBox = newSlide.shapes.addTextBox(subtitle, {
        left: 50,
        top: 75,
        width: 860,
        height: 30
      });
      subtitleBox.textFrame.textRange.font.size = subtitleSize;
      subtitleBox.textFrame.textRange.font.italic = true;
      if (color) subtitleBox.textFrame.textRange.font.color = color;
      contentTop = 112;
    }

    // If tableData is present, create native PowerPoint table
    if (tableData && tableData.rows && tableData.rows.length > 0) {
      populateSlideTable(newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, tableData, slideNum, contentTop);
    } else {
      // Body text / bullets and images
      const bodyTop = contentTop;
      const bodyBox = newSlide.shapes.addTextBox(bodyTextContent, {
        left: 50,
        top: bodyTop,
        width: hasImages ? 400 : 860,
        height: 380
      });
      bodyBox.textFrame.textRange.font.size = 18;

      if (hasImages) {
        for (const rawImg of imagesToInsert) {
          const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
          if (clean.length > 50) {
            try {
              newSlide.shapes.addImage(clean, {
                left: 480,
                top: bodyTop,
                width: 380,
                height: 300
              });
            } catch (imgErr) {}
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
  logToPPTConsole(`=== Starting Generation of ${totalSlides} Slides ===`);

  // 1. Discover Theme Blank Layout ONCE upfront to preserve presentation theme
  const layoutOptions = await getThemeBlankLayoutOptions();
  if (layoutOptions) {
    logToPPTConsole(`ℹ️ Using Theme Blank Layout.`);
  }

  // 2. Pre-process images
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

  // 3. Build each slide sequentially
  for (let i = 0; i < totalSlides; i++) {
    const slideData = slideStructures[i];
    const slideNum = i + 1;

    if (typeof onProgress === "function") {
      onProgress({
        current: slideNum,
        total: totalSlides,
        title: slideData.title
      });
    }

    try {
      await createSingleSlide(slideData, slideNum, layoutOptions);
    } catch (slideErr) {
      logToPPTConsole(`Slide ${slideNum} Error: ${slideErr.message}`, true);
      console.error(`[PPTBuilder] Slide ${slideNum} Error:`, slideErr);
      throw slideErr;
    }

    // Yield event loop for 400ms to allow PowerPoint host to finalize layout before next slide
    await new Promise(resolve => setTimeout(resolve, 400));
  }

  logToPPTConsole(`🎉 All ${totalSlides} slides created successfully!`);
}
