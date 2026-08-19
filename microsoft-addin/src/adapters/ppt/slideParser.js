/**
 * Slide Parser for Microsoft PowerPoint Adapter
 * 
 * Intelligently extracts slide structures (Title, Body Bullets, Visuals, Metadata)
 * from various LLM response formats:
 *   1. Markdown / HTML Headings (## Slide 1: Title, ## Title)
 *   2. Structured Outline Tables (| Slide 1 | Title | Content |)
 *   3. Bulleted / Numbered Outline Lists (• Slide 1: Title — Details)
 *   4. Text delimiter blocks
 *   5. Single slide fallback
 * 
 * @author Sathya AG, Principal Architect, Google
 */

export function extractSlideMetadataAndBullets(rawLines) {
  let subtitle = "";
  let visualConcept = "";
  let color = null;
  let titleSize = 44;
  let subtitleSize = 24;
  const contentBullets = [];

  for (const rawLine of rawLines) {
    if (!rawLine) continue;
    const line = rawLine.replace(/^[-•*]\s*/, "").replace(/\*\*/g, "").trim();
    if (!line) continue;

    const subMatch = line.match(/^(?:Sub-?title|Subtitle\s*Text):\s*(.*)$/i);
    if (subMatch) {
      subtitle = subMatch[1].trim();
      continue;
    }

    const colorMatch = line.match(/^(?:Color|Colour|Color\s*Scheme|Palette|Theme\s*Color):\s*(.*)$/i);
    if (colorMatch) {
      const colorVal = colorMatch[1].trim();
      const hex = colorVal.match(/#[A-Fa-f0-9]{6}/);
      if (hex) {
        color = hex[0];
      } else {
        const lower = colorVal.toLowerCase();
        if (lower.includes("blue")) color = "#004E8C";
        else if (lower.includes("green")) color = "#107C10";
        else if (lower.includes("red") || lower.includes("crimson")) color = "#A80000";
        else if (lower.includes("gold") || lower.includes("yellow")) color = "#C19C00";
        else if (lower.includes("purple") || lower.includes("violet")) color = "#5C2D91";
        else if (lower.includes("teal") || lower.includes("cyan")) color = "#008272";
        else if (lower.includes("orange") || lower.includes("brown") || lower.includes("sepia")) color = "#D83B01";
      }
      continue;
    }

    const tSizeMatch = line.match(/^Title\s*(?:Font\s*)?Size(?:\s*\(pt\))?:\s*(\d+)/i);
    if (tSizeMatch) {
      titleSize = parseInt(tSizeMatch[1], 10);
      continue;
    }

    const sSizeMatch = line.match(/^Subtitle\s*(?:Font\s*)?Size(?:\s*\(pt\))?:\s*(\d+)/i);
    if (sSizeMatch) {
      subtitleSize = parseInt(sSizeMatch[1], 10);
      continue;
    }

    const visMatch = line.match(/^(?:Visual(?:\s*Concept|\s*Description|\s*Prompt|\s*Idea)?|Image(?:\s*Prompt|\s*Concept|\s*Description)?):\s*(.*)$/i);
    if (visMatch) {
      visualConcept = visMatch[1].trim();
      continue;
    }

    if (/^(?:Main\s*)?Content:?$/i.test(line)) {
      continue;
    }

    if (/^(?:Layout|Slide\s*Layout|Template|Design\s*Theme):/i.test(line)) {
      continue;
    }

    if (/^`+$/.test(line)) continue;

    contentBullets.push(`•  ${line.replace(/^[-•*]\s*/, "")}`);
  }

  return {
    subtitle,
    visualConcept,
    color,
    titleSize: titleSize || 44,
    subtitleSize: subtitleSize || 24,
    body: contentBullets.length > 0 ? contentBullets.join("\n\n") : "• Executive slide content"
  };
}

/**
 * Parses HTML or raw Markdown text into an array of slide objects:
 * [{ title: string, subtitle: string, body: string, color: string, titleSize: number, subtitleSize: number, base64Images: string[], slideNumber: number }]
 */
export function parseSlides(htmlContent, rawText = "") {
  if (!htmlContent && !rawText) return [];

  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlContent || rawText;

  // Extract all base64 images upfront
  const allImages = Array.from(tempDiv.querySelectorAll("img"))
    .map(img => img.src || img.getAttribute("src") || "")
    .filter(s => s && s.length > 50);

  // Clean citation callouts, action toolbars, and preview containers from slide body text
  const noteCallouts = tempDiv.querySelectorAll("blockquote, .note, .ppt-deck-preview-container, .response-actions-container, [style*='background-color:#f0f6ff']");
  noteCallouts.forEach(n => n.remove());

  // -------------------------------------------------------------
  // Strategy 1: Explicit Slide Headings (H1, H2, H3)
  // -------------------------------------------------------------
  const headerEls = Array.from(tempDiv.querySelectorAll("h1, h2, h3"));
  if (headerEls.length >= 2) {
    const slides = [];
    for (let i = 0; i < headerEls.length; i++) {
      const h = headerEls[i];
      const rawTitle = (h.innerText || h.textContent || "").trim();
      const title = cleanSlideTitle(rawTitle, i + 1);

      const bodyLines = [];
      const sectionImgs = [];
      let curr = h.nextElementSibling;

      while (curr && !["H1", "H2", "H3"].includes(curr.tagName)) {
        // Collect images inside this section
        const currImgs = Array.from(curr.querySelectorAll("img"))
          .map(img => img.src || img.getAttribute("src") || "")
          .filter(s => s && s.length > 50);

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

        // Process tables inside a slide section
        if (curr.tagName === "TABLE") {
          const rows = Array.from(curr.querySelectorAll("tr"));
          rows.forEach((tr, rIdx) => {
            const cells = Array.from(tr.querySelectorAll("th, td"))
              .map(c => (c.innerText || c.textContent || "").trim())
              .filter(Boolean);
            if (cells.length >= 2) {
              if (rIdx === 0 && tr.querySelector("th")) {
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
            bodyLines.push(txt);
          }
        }
        curr = curr.nextElementSibling;
      }

      const parsed = extractSlideMetadataAndBullets(bodyLines);

      slides.push({
        slideNumber: i + 1,
        title: title,
        subtitle: parsed.subtitle,
        visualConcept: parsed.visualConcept,
        color: parsed.color,
        titleSize: parsed.titleSize,
        subtitleSize: parsed.subtitleSize,
        body: parsed.body,
        base64Images: sectionImgs
      });
    }

    if (slides.length >= 2) return slides;
  }

  // -------------------------------------------------------------
  // Strategy 2: Outline Table Unpacker (e.g. | Slide 1 | Title | Content |)
  // -------------------------------------------------------------
  const tableEls = Array.from(tempDiv.querySelectorAll("table"));
  for (const table of tableEls) {
    const rows = Array.from(table.querySelectorAll("tr"));
    const tableSlides = [];

    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      const cells = Array.from(row.querySelectorAll("td, th"))
        .map(c => (c.innerText || c.textContent || "").trim());

      if (cells.length < 2) continue;

      // Skip header row if it contains generic labels like "Slide #", "Slide Title"
      const firstCell = cells[0].toLowerCase();
      const secondCell = cells[1].toLowerCase();
      if (firstCell.includes("slide #") || firstCell.includes("slide number") || secondCell === "slide title") {
        continue;
      }

      // Check if this row represents a slide (e.g. "Slide 1", "1", "Slide 1: Intro")
      const isSlideRow = /^(?:slide\s*\d+|\d+)$/i.test(cells[0]) || 
                         /^slide\s*\d+/i.test(cells[1]) ||
                         cells.some(c => /^slide\s*\d+[:\-]/i.test(c));

      if (isSlideRow || (rows.length >= 3 && cells.length >= 2 && r > 0)) {
        let slideTitle = "";
        let contentParts = [];

        if (cells.length >= 3 && /^(?:slide\s*\d+|\d+)$/i.test(cells[0])) {
          // Format: [Slide 1] | [Title] | [Content / Visuals]
          slideTitle = cells[1];
          contentParts = cells.slice(2);
        } else if (cells.length >= 2 && /^slide\s*\d+[:\-]/i.test(cells[0])) {
          // Format: [Slide 1: Title] | [Content]
          slideTitle = cells[0].replace(/^slide\s*\d+[:\-]?\s*/i, "").trim();
          contentParts = cells.slice(1);
        } else {
          // Format: [Title] | [Content]
          slideTitle = cells[0];
          contentParts = cells.slice(1);
        }

        slideTitle = cleanSlideTitle(slideTitle, tableSlides.length + 1);

        // Format content parts into clean bullet points
        const bodyBullets = [];
        contentParts.forEach(part => {
          if (!part) return;
          // Split by semicolons, arrows ($\rightarrow$, ->, →), or periods if structured
          const cleanPart = part
            .replace(/\\rightarrow|\$\\rightarrow\$/g, "→")
            .replace(/\*\*(.*?)\*\*/g, "$1");

          // Check if part has sub-items (e.g. "Comparison: Autopilot vs Standard")
          const lines = cleanPart.split(/\r?\n|•\s*/).map(l => l.trim()).filter(Boolean);
          lines.forEach(line => {
            if (line.includes(" — ")) {
              const subSplit = line.split(" — ");
              bodyBullets.push(`• **${subSplit[0].trim()}:** ${subSplit.slice(1).join(" — ").trim()}`);
            } else if (line.includes(": ") && !line.startsWith("http")) {
              const subSplit = line.split(": ");
              bodyBullets.push(`• **${subSplit[0].trim()}:** ${subSplit.slice(1).join(": ").trim()}`);
            } else {
              bodyBullets.push(`• ${line}`);
            }
          });
        });

        tableSlides.push({
          slideNumber: tableSlides.length + 1,
          title: slideTitle,
          body: bodyBullets.join("\n\n"),
          base64Images: allImages[tableSlides.length] ? [allImages[tableSlides.length]] : []
        });
      }
    }

    if (tableSlides.length >= 2) {
      return tableSlides;
    }
  }

  // -------------------------------------------------------------
  // Strategy 3: Bulleted / Numbered Outline List Unpacker
  // -------------------------------------------------------------
  const listItems = Array.from(tempDiv.querySelectorAll("li"));
  const outlineSlides = [];

  for (let i = 0; i < listItems.length; i++) {
    const text = (listItems[i].innerText || listItems[i].textContent || "").trim();
    // Match "Slide 1: Title — Description" or "Slide 1 - Title: Description"
    const slideMatch = text.match(/^(?:Slide\s*(\d+)|\*\*Slide\s*(\d+)\*\*)\s*[:\-–—]\s*(.*)$/i);

    if (slideMatch) {
      const rest = slideMatch[3].trim();
      let title = "";
      let bodyText = "";

      if (rest.includes(" — ")) {
        const parts = rest.split(" — ");
        title = parts[0].trim();
        bodyText = parts.slice(1).join(" — ").trim();
      } else if (rest.includes(" - ")) {
        const parts = rest.split(" - ");
        title = parts[0].trim();
        bodyText = parts.slice(1).join(" - ").trim();
      } else if (rest.includes(": ")) {
        const parts = rest.split(": ");
        title = parts[0].trim();
        bodyText = parts.slice(1).join(": ").trim();
      } else {
        title = rest;
      }

      title = cleanSlideTitle(title, outlineSlides.length + 1);

      // Clean body text into structured bullet points
      const bodyBullets = [];
      if (bodyText) {
        const subItems = bodyText
          .replace(/\\rightarrow|\$\\rightarrow\$/g, "→")
          .split(/\s*;\s*|\.\s+(?=[A-Z])/)
          .filter(Boolean);

        subItems.forEach(item => {
          const cleanItem = item.replace(/^[-•*]\s*/, "").trim();
          if (cleanItem) bodyBullets.push(`• ${cleanItem}`);
        });
      }

      outlineSlides.push({
        slideNumber: outlineSlides.length + 1,
        title: title,
        body: bodyBullets.length > 0 ? bodyBullets.join("\n\n") : (bodyText ? `• ${bodyText}` : ""),
        base64Images: allImages[outlineSlides.length] ? [allImages[outlineSlides.length]] : []
      });
    }
  }

  if (outlineSlides.length >= 2) {
    return outlineSlides;
  }

  // -------------------------------------------------------------
  // Strategy 4: Raw Text Block Splitting
  // -------------------------------------------------------------
  const textContent = tempDiv.innerText || tempDiv.textContent || rawText;
  const blocks = textContent
    .split(/(?:^|\n)(?=(?:#{1,3}\s+|Slide\s*\d+[:\-]|(?:\d+\.\s+\*\*Slide)))/gi)
    .map(b => b.trim())
    .filter(b => b.length > 0);

  if (blocks.length >= 2) {
    const textSlides = [];
    blocks.forEach((block, idx) => {
      const lines = block.split("\n").map(l => l.trim()).filter(Boolean);
      if (lines.length > 0) {
        const rawTitle = lines[0]
          .replace(/^[#*\s:]+/, "")
          .replace(/^\d+\.\s*/, "")
          .replace(/^Slide\s*\d+[:\-]?\s*/i, "")
          .trim();

        const title = cleanSlideTitle(rawTitle, idx + 1);
        const parsed = extractSlideMetadataAndBullets(lines.slice(1));

        textSlides.push({
          slideNumber: idx + 1,
          title: title,
          subtitle: parsed.subtitle,
          visualConcept: parsed.visualConcept,
          color: parsed.color,
          titleSize: parsed.titleSize,
          subtitleSize: parsed.subtitleSize,
          body: parsed.body,
          base64Images: allImages[idx] ? [allImages[idx]] : []
        });
      }
    });

    if (textSlides.length >= 2) return textSlides;
  }

  // -------------------------------------------------------------
  // Strategy 5: Single Slide Fallback
  // -------------------------------------------------------------
  const lines = textContent.split("\n").map(l => l.trim()).filter(Boolean);
  const singleTitle = lines[0] ? cleanSlideTitle(lines[0], 1) : "Presentation Overview";
  const parsed = extractSlideMetadataAndBullets(lines.slice(1));

  return [{
    slideNumber: 1,
    title: singleTitle,
    subtitle: parsed.subtitle,
    visualConcept: parsed.visualConcept,
    color: parsed.color,
    titleSize: parsed.titleSize,
    subtitleSize: parsed.subtitleSize,
    body: parsed.body,
    base64Images: allImages
  }];
}

function cleanSlideTitle(rawTitle, defaultNum = 1) {
  if (!rawTitle) return `Slide ${defaultNum}`;
  let clean = rawTitle
    .replace(/^[#*\s:]+/, "")
    .replace(/^\d+[\.\)]\s*/, "")
    .replace(/^Slide\s*\d+[:\-–—]?\s*/i, "")
    .replace(/\*\*/g, "")
    .trim();
  return clean || `Slide ${defaultNum}`;
}
