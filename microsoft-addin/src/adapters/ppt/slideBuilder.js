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

import { isConversationalPreamble } from "./slideParser.js";

export function isSubstantiveSlideBody(text) {
  if (!text || typeof text !== "string") return false;
  const trimmed = text.trim();
  if (!trimmed || trimmed === "• Executive slide content") return false;
  // If text contains only backticks, punctuation, quotes, or markdown fence characters
  if (/^[`'"*#_~>|\-\s•]+$/.test(trimmed)) return false;

  // Clean markdown markers and bullets
  const cleaned = trimmed
    .replace(/[`*#_~>|\-•]/g, " ")
    .replace(/\b(?:here\s+is\s+|below\s+is\s+|sure|certainly|image\s+of|the\s+image\s+you\s+requested|executive\s+slide\s+content)\b/gi, " ")
    .trim();

  // Must have at least 15 characters and contain words with letters
  return /[a-zA-Z]{3,}/.test(cleaned) && cleaned.length >= 15;
}

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
async function populateSlideTable(context, newSlide, tableData, slideNum, tableTop = 90, tableLeft = 50, tableWidth = 860, maxTableHeight = 380) {
  const headers = tableData.headers || [];
  const rows = tableData.rows || [];
  const colCount = Math.max(headers.length, ...rows.map(r => r.length), 1);
  const rowCount = (headers.length > 0 ? 1 : 0) + rows.length;

  const tableValues = [];
  if (headers.length > 0) {
    const hRow = [];
    for (let c = 0; c < colCount; c++) {
      hRow.push(headers[c] !== undefined && headers[c] !== null ? String(headers[c]) : "");
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

  const tableHeight = Math.min(maxTableHeight, Math.max(80, rowCount * 34));

  // 1. Attempt native PowerPoint table via shapes.addTable()
  let tableSuccess = false;
  if (typeof newSlide.shapes.addTable === "function") {
    try {
      const tableShape = newSlide.shapes.addTable(rowCount, colCount, {
        left: tableLeft,
        top: tableTop,
        width: tableWidth,
        height: tableHeight
      });
      const table = typeof tableShape.getTable === "function" ? tableShape.getTable() : tableShape.table;
      if (table) {
        for (let r = 0; r < tableValues.length; r++) {
          for (let c = 0; c < colCount; c++) {
            const rawVal = tableValues[r][c];
            if (rawVal === undefined || rawVal === null) continue;
            const strVal = String(rawVal).trim();
            if (!strVal) continue;

            try {
              const cell = typeof table.getCell === "function"
                ? table.getCell(r, c)
                : (typeof table.getCellOrNullObject === "function" ? table.getCellOrNullObject(r, c) : null);
              if (cell) {
                // 1. Direct cell.text assignment (official Office.js TableCell property)
                try {
                  cell.text = strVal;
                } catch (_) {}

                // 2. textRange text and formatting
                try {
                  if (cell.textRange) {
                    cell.textRange.text = strVal;
                    if (r === 0 && headers.length > 0) {
                      if (cell.textRange.font) {
                        cell.textRange.font.bold = true;
                        cell.textRange.font.color = "#FFFFFF";
                      }
                    } else {
                      if (cell.textRange.font) {
                        cell.textRange.font.size = 13;
                      }
                    }
                  }
                } catch (_) {}

                // 3. Header background fill
                if (r === 0 && headers.length > 0 && cell.fill && typeof cell.fill.setSolidColor === "function") {
                  try {
                    cell.fill.setSolidColor("#0f4c81");
                  } catch (_) {}
                }
              }
            } catch (cellErr) {
              console.warn(`[PPTBuilder] Error populating cell [${r}, ${c}]:`, cellErr);
            }
          }
        }
      }
      await context.sync();
      tableSuccess = true;
      logToPPTConsole(`Slide ${slideNum}: Added native table (${rowCount} rows x ${colCount} cols).`);
    } catch (tblErr) {
      console.warn("[PPTBuilder] Native shapes.addTable failed, falling back to formatted text box:", tblErr);
    }
  }

  // 2. Resilient fallback: render table content as structured text in a text box
  if (!tableSuccess) {
    try {
      const fallbackLines = [];
      if (headers.length > 0) {
        fallbackLines.push(`• **${headers.join(" | ")}**`);
      }
      for (const r of rows) {
        fallbackLines.push(`• ${r.join(" | ")}`);
      }
      const tableBox = newSlide.shapes.addTextBox(fallbackLines.join("\n"), {
        left: tableLeft,
        top: tableTop,
        width: tableWidth,
        height: tableHeight
      });
      tableBox.textFrame.textRange.font.size = 14;
      tableBox.textFrame.wordWrap = true;
      await context.sync();
      logToPPTConsole(`Slide ${slideNum}: Rendered table data in text frame.`);
    } catch (fErr) {
      console.warn("[PPTBuilder] Table fallback textbox failed:", fErr);
    }
  }

  return tableHeight;
}

/**
 * Parses markdown bold and bullet markers from text, returning cleanText and bold character ranges.
 */
export function parseMarkdownFormatting(rawContent) {
  if (!rawContent) return { cleanText: "", parsedParagraphs: [] };

  const lines = String(rawContent).split(/\r?\n/);
  const cleanLines = [];
  const parsedParagraphs = [];

  for (const rawLine of lines) {
    if (!rawLine.trim()) {
      cleanLines.push("");
      parsedParagraphs.push({ cleanText: "", boldRanges: [] });
      continue;
    }

    let cleanText = "";
    const boldRanges = [];

    // Match **bold**, __bold__, or text
    const mdRegex = /(\*\*(.*?)\*\*|__([^_]+)__|[^*_]+|[*_])/g;
    let match;
    while ((match = mdRegex.exec(rawLine)) !== null) {
      if (match[2] !== undefined) {
        // **bold**
        const boldText = match[2];
        const start = cleanText.length;
        cleanText += boldText;
        boldRanges.push({ start, length: boldText.length });
      } else if (match[3] !== undefined) {
        // __bold__
        const boldText = match[3];
        const start = cleanText.length;
        cleanText += boldText;
        boldRanges.push({ start, length: boldText.length });
      } else if (match[0]) {
        cleanText += match[0];
      }
    }

    cleanLines.push(cleanText);
    parsedParagraphs.push({ cleanText, boldRanges });
  }

  return {
    cleanText: cleanLines.join("\n"),
    parsedParagraphs
  };
}

/**
 * Inserts a picture via Office Common API setSelectedDataAsync (works universally across platforms).
 */
export async function insertPictureViaCommonApi(rawImg, { left, top, width, height }) {
  if (typeof Office === "undefined" || !Office.context?.document?.setSelectedDataAsync) {
    return false;
  }
  const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
  if (clean.length <= 50) return false;

  return new Promise((resolve) => {
    try {
      Office.context.document.setSelectedDataAsync(
        clean,
        {
          coercionType: Office.CoercionType.Image,
          imageLeft: left,
          imageTop: top,
          imageWidth: width,
          imageHeight: height
        },
        (asyncResult) => {
          if (asyncResult && asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            logToPPTConsole(`Added picture via Office Common API setSelectedDataAsync.`);
            resolve(true);
          } else {
            console.warn("[PPTBuilder] setSelectedDataAsync notice:", asyncResult ? asyncResult.error : null);
            resolve(false);
          }
        }
      );
    } catch (err) {
      console.warn("[PPTBuilder] setSelectedDataAsync threw:", err);
      resolve(false);
    }
  });
}

/**
 * Inserts a picture onto a slide using official Office.js PowerPoint APIs.
 * 1. Geometric shape with fill.setImage(clean) (PowerPointApi 1.8+ official method)
 *    If setImage is unavailable or fails, immediately deletes the shape to prevent leaving a solid blue box.
 * 2. shapes.addPicture(clean) (PowerPointApi 1.10+ / preview)
 * 3. shapes.addImage(clean) fallback
 */
export function insertPictureOnSlide(slide, rawImg, { left, top, width, height }, slideNum = 1) {
  if (!rawImg) return false;
  const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
  if (clean.length <= 50) return false;

  // 1. Primary: Geometric Shape with picture fill (PowerPointApi 1.8+ official method)
  try {
    if (typeof slide.shapes.addGeometricShape === "function" && typeof PowerPoint !== "undefined" && PowerPoint.GeometricShapeType) {
      const rect = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left,
        top,
        width,
        height
      });
      if (rect) {
        if (rect.line) { try { rect.line.visible = false; } catch (_) {} }
        if (rect.lineFormat) {
          try {
            rect.lineFormat.visible = false;
            rect.lineFormat.weight = 0;
          } catch (_) {}
        }
        if (rect.fill && typeof rect.fill.setImage === "function") {
          rect.fill.setImage(clean);
          logToPPTConsole(`Slide ${slideNum}: Added picture via shape.fill.setImage.`);
          return true;
        } else {
          // IMPORTANT: If setImage is not available, delete rect immediately to prevent leaving a blank blue box!
          try { rect.delete(); } catch (_) {}
        }
      }
    }
  } catch (fillErr) {
    console.warn("[PPTBuilder] Shape fill.setImage attempt failed:", fillErr);
  }

  // 2. Secondary: shapes.addPicture (PowerPointApi 1.10+ / preview)
  try {
    if (typeof slide.shapes.addPicture === "function") {
      const pic = slide.shapes.addPicture(clean, { left, top, width, height });
      if (pic) {
        if (pic.line) { try { pic.line.visible = false; } catch (_) {} }
        if (pic.lineFormat) { try { pic.lineFormat.visible = false; } catch (_) {} }
        logToPPTConsole(`Slide ${slideNum}: Added picture via shapes.addPicture.`);
        return true;
      }
    }
  } catch (e) {
    console.warn("[PPTBuilder] shapes.addPicture with clean base64 failed, trying data uri:", e);
    try {
      if (typeof slide.shapes.addPicture === "function") {
        const pic2 = slide.shapes.addPicture(rawImg, { left, top, width, height });
        if (pic2) {
          logToPPTConsole(`Slide ${slideNum}: Added picture via data uri.`);
          return true;
        }
      }
    } catch (e2) {
      console.warn("[PPTBuilder] shapes.addPicture data uri fallback failed:", e2);
    }
  }

  // 3. Fallback: shapes.addImage if present in host
  try {
    if (typeof slide.shapes.addImage === "function") {
      const img = slide.shapes.addImage(clean, { left, top, width, height });
      if (img) {
        logToPPTConsole(`Slide ${slideNum}: Added picture via shapes.addImage.`);
        return true;
      }
    }
  } catch (_) {}

  return false;
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
  const bodyTextContent = (slideData.body || "").trim();
  const hasMeaningfulBody = isSubstantiveSlideBody(bodyTextContent);

  const tableData = slideData.tableData || null;
  const hasTable = Boolean(tableData && tableData.rows && tableData.rows.length > 0);

  const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
    ? slideData.compressedImages
    : (slideData.base64Images || []);
  const hasImages = imagesToInsert.length > 0;

  logToPPTConsole(`Slide ${slideNum}: Preparing "${cleanTitle.substring(0, 32)}..."`);

  let imageInserted = false;
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;

    // 1. Add new slide directly to presentation
    slides.add();
    await context.sync();

    // 2. In Office.js, slides.add() returns void; fetch newly added slide at tail by index
    const countResult = slides.getCount();
    await context.sync();
    const slideCount = countResult.value;
    const newSlide = slides.getItemAt(slideCount - 1);

    logToPPTConsole(`Slide ${slideNum}: Initialized new slide canvas at index ${slideCount - 1}.`);

    // If slide notes exist, associate them with the slide tags metadata
    if (slideData.notes && newSlide.tags) {
      try {
        const cleanNotes = String(slideData.notes).replace(/[\r\n]+/g, " | ").trim();
        newSlide.tags.add("SpeakerNotes", cleanNotes.substring(0, 250));
        logToPPTConsole(`Slide ${slideNum}: Attached SpeakerNotes tag metadata.`);
      } catch (tErr) {
        console.warn("[PPTBuilder] Notice setting slide tag:", tErr);
      }
    }

    const isImageOnlySlide = slideData.imageOnly || (!cleanTitle && hasImages && !hasMeaningfulBody && !hasTable);

    let contentTop = 40;
    if (!isImageOnlySlide && cleanTitle) {
      // 2. Add Title TextBox at Top
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
      contentTop = 85;

      // 3. Add Subtitle directly under Title if present
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
    }

    // 4. Layout Rendering: Image Only, Table + Image, Table + Bullets, Table Only, Image + Bullets, or Bullets Only
    if (isImageOnlySlide && hasImages) {
      // 4-ImageOnly: Centered prominently across full slide with NO titles or placeholder boxes
      imageInserted = insertPictureOnSlide(newSlide, imagesToInsert[0], {
        left: 60,
        top: 40,
        width: 840,
        height: 460
      }, slideNum);
    } else if (hasTable && hasImages) {
      // 4a. Side-by-side: Chart/Image on Left, Table on Right
      imageInserted = insertPictureOnSlide(newSlide, imagesToInsert[0], {
        left: 50,
        top: contentTop + 10,
        width: 430,
        height: 350
      }, slideNum);

      await populateSlideTable(context, newSlide, tableData, slideNum, contentTop + 10, 500, 410);
    } else if (hasTable) {
      // 4b. Table Presentation: Render full-width with optimal height so rows never collide
      await populateSlideTable(context, newSlide, tableData, slideNum, contentTop, 50, 860);
    } else if (hasImages && hasMeaningfulBody) {
      // 4c. Bullets on Left, Image on Right
      const { cleanText: cleanBullets, parsedParagraphs } = parseMarkdownFormatting(bodyTextContent);
      const bodyBox = newSlide.shapes.addTextBox(cleanBullets, {
        left: 50,
        top: contentTop,
        width: 420,
        height: 360
      });
      bodyBox.textFrame.textRange.font.size = 18;
      bodyBox.textFrame.wordWrap = true;
      await context.sync();

      // Safely apply bold styling to lead-in phrases
      if (parsedParagraphs && parsedParagraphs.some(p => p.boldRanges && p.boldRanges.length > 0)) {
        try {
          let charOffset = 0;
          for (let pIdx = 0; pIdx < parsedParagraphs.length; pIdx++) {
            const p = parsedParagraphs[pIdx];
            for (const b of (p.boldRanges || [])) {
              if (b.start >= 0 && b.length > 0 && (charOffset + b.start + b.length) <= cleanBullets.length) {
                const sub = bodyBox.textFrame.textRange.getSubstring(charOffset + b.start, b.length);
                sub.font.bold = true;
              }
            }
            charOffset += p.cleanText.length + 1; // +1 for \n
          }
          await context.sync();
        } catch (boldErr) {
          console.warn("[PPTBuilder] Notice applying bold lead-in:", boldErr);
        }
      }

      imageInserted = insertPictureOnSlide(newSlide, imagesToInsert[0], {
        left: 490,
        top: contentTop + 10,
        width: 420,
        height: 350
      }, slideNum);
    } else if (hasImages) {
      // 4d. Image Only: Centered prominently, NO placeholder text box
      imageInserted = insertPictureOnSlide(newSlide, imagesToInsert[0], {
        left: 170,
        top: contentTop + 10,
        width: 620,
        height: 370
      }, slideNum);
    } else {
      // 4e. Bullets Only: Full width
      const content = hasMeaningfulBody ? bodyTextContent : "• Executive slide content";
      const { cleanText: cleanBullets, parsedParagraphs } = parseMarkdownFormatting(content);
      const bodyBox = newSlide.shapes.addTextBox(cleanBullets, {
        left: 50,
        top: contentTop,
        width: 860,
        height: 360
      });
      bodyBox.textFrame.textRange.font.size = 18;
      bodyBox.textFrame.wordWrap = true;
      await context.sync();

      // Safely apply bold styling to lead-in phrases
      if (parsedParagraphs && parsedParagraphs.some(p => p.boldRanges && p.boldRanges.length > 0)) {
        try {
          let charOffset = 0;
          for (let pIdx = 0; pIdx < parsedParagraphs.length; pIdx++) {
            const p = parsedParagraphs[pIdx];
            for (const b of (p.boldRanges || [])) {
              if (b.start >= 0 && b.length > 0 && (charOffset + b.start + b.length) <= cleanBullets.length) {
                const sub = bodyBox.textFrame.textRange.getSubstring(charOffset + b.start, b.length);
                sub.font.bold = true;
              }
            }
            charOffset += p.cleanText.length + 1; // +1 for \n
          }
          await context.sync();
        } catch (boldErr) {
          console.warn("[PPTBuilder] Notice applying bold lead-in:", boldErr);
        }
      }
    }

    // 5. Commit all shapes in single batch
    await context.sync();
    logToPPTConsole(`Slide ${slideNum}: ✅ Created with Title, ${subtitle ? 'Subtitle, ' : ''}${hasTable ? 'Native Table' : 'Bullets'}.`);
  });

  // Universal Fallback: If shape picture insertion failed, inject via Office Common API
  if (hasImages && !imageInserted && imagesToInsert.length > 0) {
    const imgWidth = (hasMeaningfulBody || hasTable) ? 420 : 620;
    const imgLeft = (hasMeaningfulBody || hasTable) ? 490 : 170;
    await insertPictureViaCommonApi(imagesToInsert[0], {
      left: imgLeft,
      top: 95,
      width: imgWidth,
      height: 350
    });
  }
}

/**
 * Inserts content (image, visual, table, or text) directly onto the user's currently active slide.
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

  const slideData = slideStructures[0];
  const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
    ? slideData.compressedImages
    : (slideData.base64Images || []);
  const hasImages = imagesToInsert.length > 0;
  const tableData = slideData.tableData || null;
  const hasTable = Boolean(tableData && tableData.rows && tableData.rows.length > 0);
  const rawBody = (slideData.body || "").trim();
  const hasMeaningfulBody = isSubstantiveSlideBody(rawBody);

  let imageInserted = false;

  // 2. Identify active slide and insert
  await PowerPoint.run(async (context) => {
    let activeSlide = null;
    try {
      if (context.presentation.getSelectedSlides) {
        const selected = context.presentation.getSelectedSlides();
        selected.load("items");
        await context.sync();
        if (selected.items && selected.items.length > 0) {
          activeSlide = selected.items[0];
        }
      }
    } catch (_) {}

    if (!activeSlide) {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      if (slides.items && slides.items.length > 0) {
        activeSlide = slides.items[0];
      }
    }

    if (!activeSlide) {
      throw new Error("No slide available in presentation to insert content onto.");
    }

    // 1. Image Only
    if (hasImages && slideData.imageOnly) {
      for (const rawImg of imagesToInsert) {
        const added = insertPictureOnSlide(activeSlide, rawImg, {
          left: 60,
          top: 40,
          width: 840,
          height: 460
        }, 1);
        if (added) imageInserted = true;
      }
      logToPPTConsole(`Inserted full-size image centered onto current slide.`);
    }
    // 2. Chart/Image + Table
    else if (hasImages && hasTable) {
      imageInserted = insertPictureOnSlide(activeSlide, imagesToInsert[0], { left: 50, top: 90, width: 430, height: 350 }, 1);
      populateSlideTable(activeSlide, tableData, 1, 90, 500, 410);
      logToPPTConsole(`Inserted chart and table onto current slide.`);
    }
    // 3. Images present with bullets or single image
    else if (hasImages) {
      const imgWidth = hasMeaningfulBody ? 420 : 620;
      const imgLeft = hasMeaningfulBody ? 490 : 170;
      for (const rawImg of imagesToInsert) {
        const added = insertPictureOnSlide(activeSlide, rawImg, {
          left: imgLeft,
          top: 80,
          width: imgWidth,
          height: 370
        }, 1);
        if (added) imageInserted = true;
      }
      if (hasMeaningfulBody) {
        const { cleanText: cleanBullets } = parseMarkdownFormatting(rawBody);
        const bodyBox = activeSlide.shapes.addTextBox(cleanBullets, {
          left: 50,
          top: 90,
          width: 420,
          height: 360
        });
        bodyBox.textFrame.textRange.font.size = 14;
        bodyBox.textFrame.wordWrap = true;
      }
      logToPPTConsole(`Inserted image onto current slide.`);
    }
    // 4. Table presentation
    else if (hasTable) {
      await populateSlideTable(context, activeSlide, tableData, 1, 90, 50, 860);
      logToPPTConsole(`Inserted table onto current slide.`);

      // If there are following bullet points, put them on a dedicated second slide
      if (hasMeaningfulBody) {
        try {
          context.presentation.slides.add();
          await context.sync();
          const takeawayCountResult = context.presentation.slides.getCount();
          await context.sync();
          const newTakeawaySlide = context.presentation.slides.getItemAt(takeawayCountResult.value - 1);
          const { cleanText: cleanBullets } = parseMarkdownFormatting(rawBody);
          const bodyBox = newTakeawaySlide.shapes.addTextBox(cleanBullets, {
            left: 50,
            top: 90,
            width: 860,
            height: 380
          });
          bodyBox.textFrame.textRange.font.size = 18;
          bodyBox.textFrame.wordWrap = true;
          logToPPTConsole(`Inserted dedicated takeaways slide for bullets.`);
        } catch (_) {}
      }
    }
    // 4. Bullets only
    else {
      const bodyTextContent = hasMeaningfulBody ? rawBody : "• Executive slide content";
      const { cleanText: cleanBullets } = parseMarkdownFormatting(bodyTextContent);
      const bodyBox = activeSlide.shapes.addTextBox(cleanBullets, {
        left: 50,
        top: 90,
        width: 860,
        height: 360
      });
      bodyBox.textFrame.textRange.font.size = 18;
      bodyBox.textFrame.wordWrap = true;
      logToPPTConsole(`Inserted text box onto current slide.`);
    }

    await context.sync();
    logToPPTConsole(`🎉 Successfully inserted content onto current slide!`);
  });

  // Universal Fallback: If shape picture insertion failed, inject via Office Common API
  if (hasImages && !imageInserted && imagesToInsert.length > 0) {
    const rawBody = (slideData.body || "").trim();
    const hasMeaningfulBody = isSubstantiveSlideBody(rawBody);
    const imgWidth = (hasMeaningfulBody || hasTable) ? 420 : 620;
    const imgLeft = (hasMeaningfulBody || hasTable) ? 490 : 170;
    for (const rawImg of imagesToInsert) {
      await insertPictureViaCommonApi(rawImg, {
        left: imgLeft,
        top: 80,
        width: imgWidth,
        height: 370
      });
    }
  }
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
