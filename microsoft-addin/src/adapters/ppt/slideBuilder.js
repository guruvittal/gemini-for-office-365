/**
 * PowerPoint Slide Builder
 * 
 * Standard Office.js Slide Generation:
 * - Discovers Theme Blank Layout dynamically via slideMasters to eliminate template placeholders
 * - Appends new slides to presentation tail using official pattern: slides.add() -> slides.getCount() -> slides.getItemAt(count - 1)
 * - Zero shapes.load("items") or shape deletion calls to eliminate InvalidParam passed to GetItem(id) COM errors
 * - Strictly isolates tables: when a slide has a table, only the native table is rendered (no overlapping bullets)
 * - Renders native Title, Subtitle, Table / Bullets, and side-by-side visual images
 * - Zero artificial card backgrounds, zero borders, respecting native presentation theme
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

      for (const master of slideMasters.items) {
        if (!master.layouts || !master.layouts.items || master.layouts.items.length === 0) {
          continue;
        }

        // 1. Look for a layout named "blank", "em branco", "en blanco", "vide", "leer"
        let targetLayout = master.layouts.items.find(l => {
          const n = (l.name || "").toLowerCase();
          return n.includes("blank") || n.includes("branco") || n.includes("blanco") || n.includes("vide") || n.includes("leer");
        });

        // 2. Fallback: look for "empty" or "clean"
        if (!targetLayout) {
          targetLayout = master.layouts.items.find(l => {
            const n = (l.name || "").toLowerCase();
            return n.includes("empty") || n.includes("clean");
          });
        }

        if (targetLayout) {
          return {
            slideMasterId: master.id,
            layoutId: targetLayout.id
          };
        }
      }

      return null;
    });
  } catch (err) {
    console.warn("[PPTBuilder] Could not query slide masters/layouts:", err);
    return null;
  }
}

/**
 * Populates a native Microsoft PowerPoint table using PowerPoint.js shapes.addTable().
 */
function populateSlideTable(newSlide, tableData, slideNum, tableTop = 90) {
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

  const tableHeight = Math.min(380, Math.max(100, rowCount * 36));

  try {
    if (typeof newSlide.shapes.addTable === "function") {
      const tableShape = newSlide.shapes.addTable(rowCount, colCount, {
        left: 50,
        top: tableTop,
        width: 860,
        height: tableHeight,
        values: tableValues
      });
      try {
        tableShape.table.format = PowerPoint.TableFormat.lightStyle1;
      } catch (_) {}
      logToPPTConsole(`Slide ${slideNum}: Added native table (${rowCount} rows x ${colCount} cols).`);
      return;
    }
  } catch (err) {
    console.warn("[PPTBuilder] shapes.addTable with options failed, trying basic addTable:", err);
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
    logToPPTConsole(`Slide ${slideNum}: Added native table via getCell.`);
  } catch (fallbackErr) {
    console.error("[PPTBuilder] Native table shape creation failed:", fallbackErr);
    logToPPTConsole(`Slide ${slideNum}: ⚠️ Table shape notice: ${fallbackErr.message}`);
  }
}

/**
 * Creates a single slide in PowerPoint with title, body bullets, native table, or optional images.
 */
async function createSingleSlide(slideData, slideNum, layoutOptions = null) {
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

    // 1. Add slide using Theme Blank Layout if available, falling back to standard add
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

    // 2. Fetch total count to locate newly added slide at the tail
    const countResult = slides.getCount();
    await context.sync();

    const slideCount = countResult.value;
    logToPPTConsole(`Slide ${slideNum}: Appended slide at index ${slideCount - 1} (Total: ${slideCount}).`);

    const newSlide = slides.getItemAt(slideCount - 1);

    // 3. Add Title TextBox at Top
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

    // 4. Add Subtitle directly under Title if present
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

    // 5. Render Native Table OR Body Bullets / Images
    if (hasTable) {
      // Table slide: render ONLY the table! No overlapping bullets or takeaway boxes.
      populateSlideTable(newSlide, tableData, slideNum, contentTop);
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

    // 6. Commit all shapes in single batch
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

  // 1. Discover Theme Blank Layout ONCE upfront to preserve presentation theme and eliminate placeholders
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

  // 3. Build each slide sequentially, appending to end of presentation
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
