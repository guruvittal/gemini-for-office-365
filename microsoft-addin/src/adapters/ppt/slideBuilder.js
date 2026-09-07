/**
 * PowerPoint Slide Builder
 * 
 * Standard Office.js Slide Generation:
 * - Uses official Microsoft pattern: slides.add() -> slides.getCount() -> slides.getItemAt(count - 1)
 * - Appends new slides to the end of the presentation without displacing existing slides
 * - Automatically neutralizes and deletes default placeholders ("Click to add title", "Click to add subtitle")
 * - Strictly isolates tables: when a slide has a table, only the native table is rendered (no overlapping bullets)
 * - Renders native Title, Subtitle, Table / Bullets, and side-by-side visual images
 * - Zero artificial card backgrounds, zero borders, respecting native presentation theme
 * - Zero GetItem(id) COM calls to eliminate InvalidParam passed to GetItem(id) errors
 * 
 * @author Sathya AG, Principal Architect, Google
 */

function logToPPTConsole(msg, isError = false) {
  const prefix = isError ? "❌ [PPT] " : "ℹ️ [PPT] ";
  console.log(`${prefix}${msg}`);

  const debugConsole = document.getElementById("debugConsole");
  if (debugConsole) {
    const time = new Date().toLocaleTimeString();
    const entry = document.createElement("div");
    entry.style.color = isError ? "#a4262c" : "#242424";
    entry.style.marginBottom = "2px";
    entry.innerText = `[${time}] ${prefix}${msg}`;
    debugConsole.appendChild(entry);
    debugConsole.scrollTop = debugConsole.scrollHeight;
  }
}

/**
 * Compresses base64 images down to fit comfortably within PowerPoint memory limits.
 */
export function compressImageForPowerPoint(base64Str, maxWidth = 800, maxHeight = 550) {
  return new Promise((resolve) => {
    let resolved = false;
    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true;
        resolve(base64Str ? base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim() : "");
      }
    }, 1500);

    try {
      const cleanRaw = base64Str ? base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim() : "";
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
        resolve(base64Str ? base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim() : "");
      }
    }
  });
}

/**
 * Creates a single slide in PowerPoint with title, body bullets, native table, or optional images.
 */
async function createSingleSlide(slideData, slideNum) {
  const cleanTitle = (slideData.title || `Slide ${slideNum}`).replace(/\*\*/g, "").trim();
  const subtitle = slideData.subtitle || "";
  const titleSize = slideData.titleSize || 36;
  const subtitleSize = slideData.subtitleSize || 18;
  const color = slideData.color || null;
  const bodyTextContent = slideData.body || "• Executive slide content";

  const tableData = slideData.tableData || null;
  const hasTable = Boolean(tableData && tableData.rows && tableData.rows.length > 0);

  const imagesToInsert = hasTable ? [] : (
    (slideData.compressedImages && slideData.compressedImages.length > 0)
      ? slideData.compressedImages
      : (slideData.base64Images || [])
  );
  const hasImages = imagesToInsert.length > 0;

  logToPPTConsole(`Slide ${slideNum}: Preparing "${cleanTitle.substring(0, 32)}..."`);

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;

    // 1. Add slide to presentation tail and sync
    slides.add();
    await context.sync();

    // 2. Fetch total count to locate newly added slide at the end
    const countResult = slides.getCount();
    await context.sync();

    const slideCount = countResult.value;
    logToPPTConsole(`Slide ${slideNum}: Appended slide at index ${slideCount - 1} (Total: ${slideCount}).`);

    const newSlide = slides.getItemAt(slideCount - 1);

    // 3. Clear and delete default template placeholders ("Click to add title", "Click to add subtitle")
    try {
      newSlide.shapes.load("items");
      await context.sync();

      if (newSlide.shapes.items && newSlide.shapes.items.length > 0) {
        for (let i = newSlide.shapes.items.length - 1; i >= 0; i--) {
          const s = newSlide.shapes.items[i];
          try {
            if (s.textFrame && s.textFrame.textRange) {
              s.textFrame.textRange.text = "";
            }
          } catch (_) {}
          try {
            s.delete();
          } catch (_) {}
        }
        await context.sync();
      }
    } catch (cleanErr) {
      console.warn("[PPTBuilder] Notice clearing placeholders:", cleanErr.message);
    }

    // 4. Add Title TextBox at Top
    const titleBox = newSlide.shapes.addTextBox(cleanTitle, {
      left: 50,
      top: 30,
      width: 860,
      height: 50
    });
    titleBox.textFrame.textRange.font.size = titleSize;
    titleBox.textFrame.textRange.font.bold = true;
    if (color) {
      titleBox.textFrame.textRange.font.color = color;
    }

    // 5. Add Subtitle if exists directly under Title
    let contentTop = 85;
    if (subtitle) {
      const subtitleBox = newSlide.shapes.addTextBox(subtitle, {
        left: 50,
        top: 80,
        width: 860,
        height: 35
      });
      subtitleBox.textFrame.textRange.font.size = subtitleSize;
      subtitleBox.textFrame.textRange.font.italic = true;
      if (color) {
        subtitleBox.textFrame.textRange.font.color = color;
      }
      contentTop = 120;
    }

    // 6. Render Native Table OR Body Bullets / Images
    if (hasTable) {
      // Table Slide: Render ONLY the native table (zero overlapping bullets/takeaways)
      const numRows = tableData.rows.length + 1;
      const numCols = Math.max(
        tableData.headers ? tableData.headers.length : 0,
        ...tableData.rows.map(r => r.length),
        1
      );
      const tableValues = [
        tableData.headers || [],
        ...tableData.rows
      ];
      for (let r = 0; r < tableValues.length; r++) {
        while (tableValues[r].length < numCols) {
          tableValues[r].push("");
        }
      }

      const tableHeight = Math.min(380, Math.max(100, numRows * 36));
      const tableShape = newSlide.shapes.addTable(numRows, numCols, {
        left: 50,
        top: contentTop,
        width: 860,
        height: tableHeight,
        values: tableValues
      });

      try {
        tableShape.table.format = PowerPoint.TableFormat.lightStyle1;
      } catch (_) {}
      logToPPTConsole(`Slide ${slideNum}: Added native table (${numRows} rows x ${numCols} cols).`);
    } else {
      // Standard clean body bullets
      const cleanBullets = bodyTextContent.replace(/\*\*/g, "").replace(/__/g, "");
      const bodyBox = newSlide.shapes.addTextBox(cleanBullets, {
        left: 50,
        top: contentTop,
        width: hasImages ? 400 : 860,
        height: 360
      });
      bodyBox.textFrame.textRange.font.size = 18;
      bodyBox.textFrame.wordWrap = true;

      // Add Image if available (only on non-table slides)
      if (hasImages) {
        for (const rawImg of imagesToInsert) {
          const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
          if (clean.length > 50) {
            try {
              newSlide.shapes.addImage(clean, {
                left: 480,
                top: contentTop,
                width: 380,
                height: 300
              });
              logToPPTConsole(`Slide ${slideNum}: Attached image.`);
            } catch (imgErr) {
              logToPPTConsole(`Slide ${slideNum}: Image notice: ${imgErr.message}`);
            }
          }
        }
      }
    }

    // 7. Commit all shapes in single batch
    await context.sync();
    logToPPTConsole(`Slide ${slideNum}: ✅ Created with Title, ${subtitle ? 'Subtitle, ' : ''}${hasTable ? 'Native Table' : 'Bullets'}.`);
  });
}

/**
 * Builds all parsed slides sequentially, appending them to the end of the presentation.
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
  logToPPTConsole(`=== Starting Generation of ${totalSlides} Slide(s) ===`);

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

  // 2. Build each slide sequentially, appending to end of presentation
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
      await createSingleSlide(slideData, slideNum);
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
