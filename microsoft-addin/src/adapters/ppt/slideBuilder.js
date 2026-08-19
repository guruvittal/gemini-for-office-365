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
 * Creates a single slide in PowerPoint with title at top, subtitle directly under title, and body bullets.
 */
async function createSingleSlide(slideData, slideNum) {
  const cleanTitle = (slideData.title || `Slide ${slideNum}`).replace(/\*\*/g, "").trim();
  const subtitle = slideData.subtitle || "";
  const titleSize = slideData.titleSize || 40;
  const subtitleSize = slideData.subtitleSize || 22;
  const color = slideData.color || null;
  const bodyTextContent = slideData.body || "• Executive slide content";

  const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
    ? slideData.compressedImages
    : (slideData.base64Images || []);
  const hasImages = imagesToInsert.length > 0;

  logToPPTConsole(`Slide ${slideNum}: Preparing "${cleanTitle.substring(0, 32)}..."`);

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;

    // 1. Add standard slide and load slides collection safely
    const addedSlide = slides.add();
    await context.sync();

    slides.load("items");
    await context.sync();

    // Use returned reference if available, otherwise pull from loaded items array
    const newSlide = addedSlide || (slides.items && slides.items.length > 0 ? slides.items[slides.items.length - 1] : null);
    if (!newSlide) {
      throw new Error("Unable to obtain reference to newly created slide.");
    }
    logToPPTConsole(`Slide ${slideNum}: Added slide to presentation.`);

    // 3. Intelligently map content to existing template placeholders
    let titlePopulated = false;
    let subtitlePopulated = false;
    let bodyPopulated = false;

    try {
      newSlide.shapes.load("items");
      await context.sync();

      if (newSlide.shapes.items && newSlide.shapes.items.length > 0) {
        for (let i = 0; i < newSlide.shapes.items.length; i++) {
          const s = newSlide.shapes.items[i];
          const sName = (s.name || "").toLowerCase();

          try {
            if (!titlePopulated && sName.includes("title")) {
              s.textFrame.textRange.text = cleanTitle;
              await context.sync();
              titlePopulated = true;
            } else if (!subtitlePopulated && sName.includes("subtitle") && subtitle) {
              s.textFrame.textRange.text = subtitle;
              await context.sync();
              subtitlePopulated = true;
            } else if (!bodyPopulated && (sName.includes("content") || sName.includes("body") || sName.includes("text") || sName.includes("placeholder") || sName.includes("object"))) {
              s.textFrame.textRange.text = bodyTextContent;
              await context.sync();
              bodyPopulated = true;
            } else if (sName.includes("subtitle") && !subtitle) {
              // Neutralize empty subtitle placeholder watermark
              s.textFrame.textRange.text = "\u00A0";
              await context.sync();
            }
          } catch (shapeErr) {
            // Shape did not accept text; fallback addTextBox will trigger below
          }
        }
      }
    } catch (e) {
      console.warn("Notice reading template placeholders:", e.message);
    }

    // 4. Fallbacks for missing placeholders (add custom boxes)
    if (!titlePopulated) {
      const titleBox = newSlide.shapes.addTextBox(cleanTitle, { left: 50, top: 40, width: 860, height: 60 });
      titleBox.textFrame.textRange.font.size = titleSize;
      titleBox.textFrame.textRange.font.bold = true;
      if (color) titleBox.textFrame.textRange.font.color = color;
      await context.sync();
    }

    if (subtitle && !subtitlePopulated) {
      const subtitleBox = newSlide.shapes.addTextBox(subtitle, { left: 50, top: 110, width: 860, height: 40 });
      subtitleBox.textFrame.textRange.font.size = subtitleSize;
      subtitleBox.textFrame.textRange.font.italic = true;
      if (color) subtitleBox.textFrame.textRange.font.color = color;
      await context.sync();
    }

    const bodyTop = subtitle ? 160 : 110;
    if (!bodyPopulated) {
      const bodyBox = newSlide.shapes.addTextBox(bodyTextContent, { left: 50, top: bodyTop, width: hasImages ? 400 : 860, height: 340 });
      bodyBox.textFrame.textRange.font.size = 18;
      await context.sync();
    }

    // 5. Add Image if available
    if (hasImages) {
      for (const rawImg of imagesToInsert) {
        const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
        if (clean.length > 50) {
          try {
            newSlide.shapes.addImage(clean, { left: 480, top: bodyTop, width: 380, height: 300 });
            await context.sync();
            logToPPTConsole(`Slide ${slideNum}: Attached image.`);
          } catch (imgErr) {
            logToPPTConsole(`Slide ${slideNum}: Image notice: ${imgErr.message}`);
          }
        }
      }
    }

    logToPPTConsole(`Slide ${slideNum}: ✅ Created with Title, ${subtitle ? 'Subtitle, ' : ''}and Bullets.`);
  });
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

  // 2. Build each slide sequentially
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
