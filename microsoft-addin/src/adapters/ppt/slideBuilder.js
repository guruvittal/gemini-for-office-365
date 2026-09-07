/**
 * PowerPoint Slide Builder
 * 
 * Generates executive presentation slides directly in Microsoft PowerPoint via Office.js.
 * - Always appends new slides to the end of the presentation
 * - Discovers and uses the presentation theme's Blank layout (no placeholders, inherits master artwork)
 * - Renders clean typography (title, subtitle, bullets) respecting the user's presentation theme
 * - Supports native PowerPoint tables and bold lead-in formatting without artificial card backgrounds or borders
 * - Uses isolated, atomic PowerPoint.run execution contexts with resilient fallbacks
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
 * Compresses a base64 image down to fit comfortably within PowerPoint memory limits.
 */
export function compressImageForPowerPoint(base64Str, maxWidth = 1920, maxHeight = 1080) {
  return new Promise((resolve) => {
    if (!base64Str || typeof base64Str !== "string") {
      return resolve(null);
    }

    if (!base64Str.startsWith("data:image/") && !/^[A-Za-z0-9+/=]+$/.test(base64Str.substring(0, 100))) {
      return resolve(null);
    }

    const src = base64Str.startsWith("data:image/")
      ? base64Str
      : `data:image/png;base64,${base64Str.trim()}`;

    let resolved = false;
    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true;
        resolve(base64Str.replace(/^data:image\/[^;]+;base64,/, "").replace(/[\r\n\s]+/g, "").trim());
      }
    }, 4000);

    try {
      const img = new Image();
      img.crossOrigin = "anonymous";
      img.onload = () => {
        if (resolved) return;
        resolved = true;
        clearTimeout(timer);

        try {
          let { width, height } = img;
          if (width <= maxWidth && height <= maxHeight && src.length < 2 * 1024 * 1024) {
            return resolve(src.replace(/^data:image\/[^;]+;base64,/, "").replace(/[\r\n\s]+/g, "").trim());
          }

          if (width > maxWidth) {
            height = Math.round((height * maxWidth) / width);
            width = maxWidth;
          }
          if (height > maxHeight) {
            width = Math.round((width * maxHeight) / height);
            height = maxHeight;
          }

          const canvas = document.createElement("canvas");
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext("2d");
          ctx.drawImage(img, 0, 0, width, height);

          const compressedDataUrl = canvas.toDataURL("image/jpeg", 0.85);
          resolve(compressedDataUrl.replace(/^data:image\/[^;]+;base64,/, "").replace(/[\r\n\s]+/g, "").trim());
        } catch (_) {
          resolve(base64Str.replace(/^data:image\/[^;]+;base64,/, "").replace(/[\r\n\s]+/g, "").trim());
        }
      };

      img.onerror = () => {
        if (!resolved) {
          resolved = true;
          clearTimeout(timer);
          resolve(base64Str.replace(/^data:image\/[^;]+;base64,/, "").replace(/[\r\n\s]+/g, "").trim());
        }
      };

      img.src = src;
    } catch (_) {
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
export async function getThemeBlankLayoutOptions() {
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
 * Parses markdown bold and italic formatting from body text.
 */
export function parseMarkdownFormatting(rawContent) {
  if (!rawContent || typeof rawContent !== "string") {
    return { cleanText: "", parsedParagraphs: [] };
  }

  const lines = rawContent.split("\n");
  const parsedParagraphs = [];
  const cleanLines = [];

  for (const rawLine of lines) {
    let cleanText = "";
    const boldRanges = [];
    const italicRanges = [];

    const boldRegex = /(\*\*|__)(.*?)\1/g;
    let lastIndex = 0;
    let match;

    while ((match = boldRegex.exec(rawLine)) !== null) {
      const beforeMatch = rawLine.substring(lastIndex, match.index);
      cleanText += beforeMatch;
      const boldText = match[2];
      const start = cleanText.length;
      cleanText += boldText;
      boldRanges.push({ start, length: boldText.length });
      lastIndex = match.index + match[0].length;
    }
    cleanText += rawLine.substring(lastIndex);

    // Expand bold range to include trailing colon or dash
    for (const b of boldRanges) {
      if (cleanText.charAt(b.start + b.length) === ":" || cleanText.charAt(b.start + b.length) === "—") {
        b.length += 1;
      }
    }

    // If no bold ranges found from markdown, auto-detect colon/dash lead-in (e.g. "• Key Point: details")
    if (boldRanges.length === 0) {
      const colonIdx = cleanText.indexOf(":");
      const dashIdx = cleanText.indexOf("—");
      const sepIdx = colonIdx > 0 ? colonIdx : (dashIdx > 0 ? dashIdx : -1);
      if (sepIdx > 0 && sepIdx < 50) {
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
 * Applies native bold and italic font styling to paragraph ranges inside a text frame.
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
          for (const b of parsed.boldRanges) {
            try {
              if (b.length > 0 && b.start + b.length <= (p.text || "").length) {
                const sub = p.getSubstring(b.start, b.length);
                sub.font.bold = true;
              }
            } catch (_) {}
          }
        }
        pIdx++;
      }
      await context.sync();
    }
  } catch (err) {
    console.warn("Paragraph formatting notice:", err);
  }
}

/**
 * Populates title, subtitle, native table, bullets, and images cleanly on a new slide.
 * Respects the existing presentation theme without any artificial card shapes or borders.
 */
async function populateSlideShapes(newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, bodyTextContent, hasImages, imagesToInsert, slideNum, slideData = null, context = null) {
  // 1. Add Title TextBox at Top
  const titleBox = newSlide.shapes.addTextBox(cleanTitle, {
    left: 50,
    top: 35,
    width: 860,
    height: 50
  });
  titleBox.textFrame.textRange.font.size = titleSize || 36;
  titleBox.textFrame.textRange.font.bold = true;
  if (color) {
    titleBox.textFrame.textRange.font.color = color;
  }

  // 2. Add Subtitle TextBox directly under Title
  if (subtitle) {
    const subtitleBox = newSlide.shapes.addTextBox(subtitle, {
      left: 50,
      top: 85,
      width: 860,
      height: 35
    });
    subtitleBox.textFrame.textRange.font.size = subtitleSize || 18;
    subtitleBox.textFrame.textRange.font.italic = true;
    if (color) {
      subtitleBox.textFrame.textRange.font.color = color;
    }
  }

  const bodyTop = subtitle ? 130 : 95;

  // 3. Render Table if table data is provided
  const tableData = slideData ? slideData.tableData : null;
  const hasTable = Boolean(tableData && tableData.rows && tableData.rows.length > 0);

  if (hasTable) {
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

    // Ensure all rows match column count
    for (let r = 0; r < tableValues.length; r++) {
      while (tableValues[r].length < numCols) {
        tableValues[r].push("");
      }
    }

    const tableShape = newSlide.shapes.addTable(numRows, numCols, {
      left: 50,
      top: bodyTop,
      height: Math.min(320, numRows * 36),
      values: tableValues
    });

    try {
      tableShape.table.format = PowerPoint.TableFormat.lightStyle1;
    } catch (_) {}

    // If takeaway or narrative notes exist, render clean text box underneath the table
    const takeaway = (slideData && (slideData.takeaway || slideData.additionalBody)) ? (slideData.takeaway || slideData.additionalBody) : "";
    if (takeaway && takeaway.trim().length > 0) {
      const takeawayTop = Math.min(460, bodyTop + Math.min(320, numRows * 36) + 16);
      const takeawayBox = newSlide.shapes.addTextBox(takeaway.trim(), {
        left: 50,
        top: takeawayTop,
        width: 860,
        height: 50
      });
      takeawayBox.textFrame.textRange.font.size = 14;
      takeawayBox.textFrame.textRange.font.italic = true;
    }
  } else {
    // Standard Body Text Box
    const { cleanText, parsedParagraphs } = parseMarkdownFormatting(bodyTextContent);
    const bodyBox = newSlide.shapes.addTextBox(cleanText, {
      left: 50,
      top: bodyTop,
      width: hasImages ? 440 : 860,
      height: 360
    });
    bodyBox.textFrame.textRange.font.size = 17;
    bodyBox.textFrame.wordWrap = true;

    // Apply bold formatting to lead-ins if context is available
    if (context) {
      try {
        await applyParagraphFormatting(bodyBox, parsedParagraphs, context);
      } catch (_) {}
    }
  }

  // 4. Attach Image if available (and not a table slide)
  if (hasImages && !hasTable) {
    for (const rawImg of imagesToInsert) {
      const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
      if (clean.length > 50) {
        try {
          newSlide.shapes.addImage(clean, {
            left: 510,
            top: bodyTop,
            width: 370,
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

/**
 * Creates a single slide atomically in PowerPoint, appending it to the end of the presentation.
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

  let addedSuccessfully = false;

  // 1. Attempt addition with Theme Blank Layout if available
  if (layoutOptions) {
    try {
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.add(layoutOptions);
        await context.sync();

        slides.load("items");
        await context.sync();

        const newSlide = slides.items[slides.items.length - 1];
        await populateSlideShapes(newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, bodyTextContent, hasImages, imagesToInsert, slideNum, slideData, context);
        await context.sync();
        addedSuccessfully = true;
      });
    } catch (layoutErr) {
      console.warn(`[PPTBuilder] Theme layout add failed (${layoutErr.message}), falling back to standard slide add.`);
    }
  }

  // 2. Reliable Fallback: Add standard slide if theme blank layout was unavailable or failed
  if (!addedSuccessfully) {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.add();
      await context.sync();

      slides.load("items");
      await context.sync();

      const newSlide = slides.items[slides.items.length - 1];
      await populateSlideShapes(newSlide, cleanTitle, subtitle, titleSize, subtitleSize, color, bodyTextContent, hasImages, imagesToInsert, slideNum, slideData, context);
      await context.sync();
    });
  }

  logToPPTConsole(`Slide ${slideNum}: ✅ Created with Title, ${subtitle ? "Subtitle, " : ""}${hasTable ? "Native Table" : "Bullets"}.`);
}

/**
 * Builds all parsed slides sequentially, appending them to the end of the presentation.
 * @param {Array} slideStructures - Array of parsed slide objects
 * @param {Object} options - { mode: "insert" | "replace" }
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

  // 3. Build each slide sequentially, appending to the end
  let successfulSlides = 0;
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
      successfulSlides++;
    } catch (slideErr) {
      logToPPTConsole(`Slide ${slideNum} Error: ${slideErr.message}`, true);
      console.error(`[PPTBuilder] Slide ${slideNum} Error:`, slideErr);
      throw slideErr;
    }

    // Yield event loop for 400ms to allow PowerPoint host to finalize layout before next slide
    await new Promise(resolve => setTimeout(resolve, 400));
  }

  logToPPTConsole(`🎉 All ${successfulSlides} slide(s) created successfully at the end of the presentation!`);
}
