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

/**
 * Detects whether a string is conversational preamble or tool disclaimer rather than slide content.
 */
export function isConversationalPreamble(line) {
  if (!line) return false;
  const l = line.trim().toLowerCase();
  return (
    /^(?:here\s+(?:is|are)\s+(?:the|your|a|this)?\s*(?:content|slide|slides|presentation|deck|breakdown|summary|overview)?|below\s+(?:is|are)\s+(?:the|your|a)?\s*(?:content|slide|slides)?)/i.test(l) ||
    /^(?:sure[!,.]?|certainly[!,.]?|of\s+course[!,.]?|absolutely[!,.]?|great[!,.]?|i'\''d\s+be\s+happy\s+to|i\s+have\s+(?:created|prepared|generated))/i.test(l) ||
    /^(?:based\s+on\s+your\s+(?:request|focus|documents|data)|as\s+requested|per\s+your\s+request)/i.test(l) ||
    /^note:\s*i\s+don'\''t\s+have\s+a\s+direct\s+slide/i.test(l) ||
    /^\*?note:\s*/i.test(l) ||
    /^---+$/.test(l)
  );
}

export function extractSlideMetadataAndBullets(rawLines) {
  let subtitle = "";
  let visualConcept = "";
  let color = null;
  let titleSize = 44;
  let subtitleSize = 24;
  const contentBullets = [];

  for (const rawLine of rawLines) {
    if (!rawLine) continue;
    if (isConversationalPreamble(rawLine)) continue;

    // Remove leading bullet/hash symbols and bold markers
    let line = rawLine
      .replace(/^[-•*]\s*/, "")
      .replace(/^[#\s]+/, "")
      .trim();
    if (!line) continue;

    const subMatch = line.match(/^(?:Sub-?title|Subtitle\s*Text|Effective|Target|Date|Author|Quarter):\s*(.*)$/i);
    if (subMatch && !subtitle) {
      subtitle = (subMatch[1] ? subMatch[1].trim() : line).replace(/\*\*/g, "").replace(/^[*_]+|[*_]+$/g, "");
      continue;
    }

    const cleanLine = line.replace(/\*\*/g, "").replace(/^[*_]+|[*_]+$/g, "").trim();

    const colorMatch = cleanLine.match(/^(?:Color|Colour|Color\s*Scheme|Palette|Theme\s*Color):\s*(.*)$/i);
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

    const tSizeMatch = cleanLine.match(/^Title\s*(?:Font\s*)?Size(?:\s*\(pt\))?:\s*(\d+)/i);
    if (tSizeMatch) {
      titleSize = parseInt(tSizeMatch[1], 10);
      continue;
    }

    const sSizeMatch = cleanLine.match(/^Subtitle\s*(?:Font\s*)?Size(?:\s*\(pt\))?:\s*(\d+)/i);
    if (sSizeMatch) {
      subtitleSize = parseInt(sSizeMatch[1], 10);
      continue;
    }

    const visMatch = cleanLine.match(/^(?:Visual(?:\s*Concept|\s*Description|\s*Prompt|\s*Idea)?|Image(?:\s*Prompt|\s*Concept|\s*Description)?):\s*(.*)$/i);
    if (visMatch) {
      visualConcept = visMatch[1].trim();
      continue;
    }

    if (/^(?:Main\s*)?Content:?$/i.test(cleanLine)) {
      continue;
    }

    if (/^(?:Layout|Slide\s*Layout|Template|Design\s*Theme):/i.test(cleanLine)) {
      continue;
    }

    if (/^`+$/.test(cleanLine)) continue;

    // Clean bullet formatting
    const formattedBullet = cleanLine.replace(/^[-•*]\s*/, "").replace(/^[*_]+|[*_]+$/g, "").trim();
    if (formattedBullet) {
      contentBullets.push(`•  ${formattedBullet}`);
    }
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
 * Extracts clean structured table content and formats it into executive comparison bullets.
 */
export function extractTableContent(tableEl) {
  if (!tableEl) return null;
  const rows = Array.from(tableEl.querySelectorAll("tr"));
  if (rows.length === 0) return null;

  const headers = [];
  const headerRow = tableEl.querySelector("thead tr") || rows[0];
  if (headerRow) {
    const ths = Array.from(headerRow.querySelectorAll("th, td"))
      .map(c => (c.innerText || c.textContent || "").trim());
    if (ths.length > 0) headers.push(...ths);
  }

  const dataRows = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (row === headerRow && tableEl.querySelector("th")) continue;
    const cells = Array.from(row.querySelectorAll("td, th"))
      .map(c => (c.innerText || c.textContent || "").trim());
    if (cells.length > 0 && cells.some(c => c.length > 0)) {
      dataRows.push(cells);
    }
  }

  // Format into clean structured comparison bullets for presentation slides
  const formattedBullets = [];
  dataRows.forEach(row => {
    if (row.length === 1) {
      formattedBullets.push(`• ${row[0]}`);
    } else if (row.length === 2) {
      formattedBullets.push(`• **${row[0]}:** ${row[1]}`);
    } else {
      const itemName = row[0];
      const details = [];
      for (let c = 1; c < row.length; c++) {
        const h = headers[c] ? `${headers[c]}: ` : "";
        details.push(`${h}${row[c]}`);
      }
      formattedBullets.push(`• **${itemName}:** ${details.join("  |  ")}`);
    }
  });

  return {
    headers,
    dataRows,
    bullets: formattedBullets,
    bulletText: formattedBullets.join("\n\n")
  };
}

/**
 * Parses raw markdown table (| Header 1 | Header 2 |) into structured columns and rows.
 */
export function parseMarkdownTable(text) {
  if (!text || !text.includes("|")) return null;
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.startsWith("|") && l.endsWith("|"));
  if (lines.length < 2) return null;

  const validRows = [];
  for (const line of lines) {
    const cells = line.split("|").slice(1, -1).map(c => c.trim().replace(/\*\*/g, ""));
    // Filter out separator lines like |---|---|
    if (cells.length > 0 && !cells.every(c => /^[-:\s]+$/.test(c))) {
      validRows.push(cells);
    }
  }

  if (validRows.length < 2) return null;

  const headers = validRows[0];
  const dataRows = validRows.slice(1);

  const formattedBullets = [];
  dataRows.forEach(row => {
    if (row.length === 1) {
      formattedBullets.push(`• ${row[0]}`);
    } else if (row.length === 2) {
      formattedBullets.push(`• **${row[0]}:** ${row[1]}`);
    } else {
      const itemName = row[0];
      const details = [];
      for (let c = 1; c < row.length; c++) {
        const h = headers[c] ? `${headers[c]}: ` : "";
        details.push(`${h}${row[c]}`);
      }
      formattedBullets.push(`• **${itemName}:** ${details.join("  |  ")}`);
    }
  });

  return {
    headers,
    dataRows,
    bullets: formattedBullets,
    bulletText: formattedBullets.join("\n\n")
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
  let headerEls = Array.from(tempDiv.querySelectorAll("h1, h2"));
  let boundaryTagNames = ["H1", "H2"];

  // If no H1 or H2, fallback to H3
  if (headerEls.length === 0) {
    headerEls = Array.from(tempDiv.querySelectorAll("h3"));
    boundaryTagNames = ["H1", "H2", "H3"];
  }

  if (headerEls.length >= 1) {
    const slides = [];
    for (let i = 0; i < headerEls.length; i++) {
      const h = headerEls[i];
      const rawTitle = (h.innerText || h.textContent || "").trim();
      const title = cleanSlideTitle(rawTitle, slides.length + 1);
      if (!title) continue;

      const bodyLines = [];
      const sectionImgs = [];
      let sectionTableData = null;
      let curr = h.nextElementSibling;

      while (curr && !boundaryTagNames.includes(curr.tagName)) {
        if (curr.tagName === "HR") {
          curr = curr.nextElementSibling;
          continue;
        }

        // Collect images inside this section
        const currImgs = Array.from(curr.querySelectorAll("img"))
          .map(img => img.src || img.getAttribute("src") || "")
          .filter(s => s && s.length > 50);

        if (curr.tagName === "IMG") {
          const imgSrc = curr.src || curr.getAttribute("src") || "";
          if (imgSrc.length > 50) currImgs.push(imgSrc);
        }
        if (currImgs.length > 0) sectionImgs.push(...currImgs);

        // Skip verified source footers or preamble/notes
        if (curr.innerText && (curr.innerText.includes("Verified Sources") || isConversationalPreamble(curr.innerText))) {
          curr = curr.nextElementSibling;
          continue;
        }

        // Process tables inside a slide section
        if (curr.tagName === "TABLE") {
          const tbl = extractTableContent(curr);
          if (tbl && tbl.bullets.length > 0) {
            bodyLines.push(...tbl.bullets);
            sectionTableData = {
              headers: tbl.headers,
              rows: tbl.dataRows
            };
          }
        } else if (curr.tagName === "H3" && !boundaryTagNames.includes("H3")) {
          // H3 inside an H1/H2 slide: format as subsection heading
          const subHeading = (curr.innerText || curr.textContent || "").trim();
          if (subHeading && !isConversationalPreamble(subHeading)) {
            bodyLines.push(`### ${subHeading}`);
          }
        } else if (curr.tagName === "UL" || curr.tagName === "OL") {
          Array.from(curr.querySelectorAll("li")).forEach(li => {
            const txt = (li.innerText || li.textContent || "").trim();
            if (txt && !isConversationalPreamble(txt)) {
              bodyLines.push(`• ${txt.replace(/^[-•*]\s*/, "")}`);
            }
          });
        } else {
          const txt = (curr.innerText || curr.textContent || "").trim();
          if (txt && !isConversationalPreamble(txt)) {
            bodyLines.push(txt);
          }
        }
        curr = curr.nextElementSibling;
      }

      const parsed = extractSlideMetadataAndBullets(bodyLines);

      slides.push({
        slideNumber: slides.length + 1,
        title: title,
        subtitle: parsed.subtitle,
        visualConcept: parsed.visualConcept,
        color: parsed.color,
        titleSize: parsed.titleSize,
        subtitleSize: parsed.subtitleSize,
        body: parsed.body,
        tableData: sectionTableData,
        base64Images: sectionImgs
      });
    }

    if (slides.length >= 1) return slides;
  }

  // -------------------------------------------------------------
  // Strategy 2: Outline Table Unpacker (e.g. | Slide # | Title | Content |)
  // ONLY activates when table explicitly describes a presentation outline!
  // -------------------------------------------------------------
  const tableEls = Array.from(tempDiv.querySelectorAll("table"));
  for (const table of tableEls) {
    const rows = Array.from(table.querySelectorAll("tr"));
    if (rows.length < 2) continue;

    // Check if this table is explicitly a slide outline table
    const headerText = (rows[0].innerText || rows[0].textContent || "").toLowerCase();
    const isExplicitOutlineHeader = 
      headerText.includes("slide #") || 
      headerText.includes("slide number") || 
      headerText.includes("slide title") || 
      headerText.includes("deck outline") ||
      headerText.includes("presentation outline");

    let explicitSlideRowCount = 0;
    for (let r = 0; r < rows.length; r++) {
      const cells = Array.from(rows[r].querySelectorAll("td, th"))
        .map(c => (c.innerText || c.textContent || "").trim());
      if (cells.length >= 2 && (/^slide\s*\d+/i.test(cells[0]) || /^slide\s*\d+/i.test(cells[1]))) {
        explicitSlideRowCount++;
      }
    }

    // ONLY treat as slide outline if headers or multiple cells explicitly identify slides!
    if (!isExplicitOutlineHeader && explicitSlideRowCount < 2) {
      // This is a data/comparison table, NOT a slide outline! Skip Strategy 2.
      continue;
    }

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

      // Check if this row represents a slide (e.g. "Slide 1", "Slide 1: Intro")
      const isSlideRow = /^slide\s*\d+/i.test(cells[0]) || 
                         /^slide\s*\d+/i.test(cells[1]) ||
                         cells.some(c => /^slide\s*\d+[:\-]/i.test(c));

      if (isSlideRow || (isExplicitOutlineHeader && r > 0)) {
        let slideTitle = "";
        let contentParts = [];

        if (cells.length >= 3 && /^slide\s*\d+/i.test(cells[0])) {
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
          const cleanPart = part
            .replace(/\\rightarrow|\$\\rightarrow\$/g, "→")
            .replace(/\*\*(.*?)\*\*/g, "$1");

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

  if (outlineSlides.length >= 1) {
    return outlineSlides;
  }

  // -------------------------------------------------------------
  // Strategy 4: Standalone Data Table (HTML <table> or Markdown |...|)
  // When a table is present without multiple explicit slide markers,
  // produce a dedicated slide with full tabular layout data.
  // -------------------------------------------------------------
  const standaloneTable = tempDiv.querySelector("table");
  if (standaloneTable) {
    const tbl = extractTableContent(standaloneTable);
    if (tbl && tbl.dataRows.length > 0) {
      let tableTitle = "";
      let prevEl = standaloneTable.previousElementSibling;
      while (prevEl) {
        const txt = (prevEl.innerText || prevEl.textContent || "").trim();
        if (txt && !txt.startsWith("Verified Sources") && !isConversationalPreamble(txt)) {
          tableTitle = txt;
          break;
        }
        prevEl = prevEl.previousElementSibling;
      }

      if (!tableTitle) {
        const firstLine = (tempDiv.innerText || tempDiv.textContent || "")
          .split("\n")
          .map(l => l.trim())
          .find(l => l.length > 0 && !l.startsWith("|") && !isConversationalPreamble(l));
        if (firstLine) tableTitle = firstLine;
      }

      if (!tableTitle && tbl.headers.length > 0) {
        tableTitle = `${tbl.headers[0]} Comparison`;
      }

      const cleanTitle = cleanSlideTitle(tableTitle || "Comparison Table", 1) || "Comparison Table";
      return [{
        slideNumber: 1,
        title: cleanTitle,
        subtitle: "",
        visualConcept: "",
        color: null,
        titleSize: 36,
        subtitleSize: 20,
        body: tbl.bulletText,
        tableData: {
          headers: tbl.headers,
          rows: tbl.dataRows
        },
        base64Images: allImages
      }];
    }
  }

  // Check for raw markdown table in text
  const textContent = rawText || tempDiv.textContent || tempDiv.innerText || "";
  const mdTable = parseMarkdownTable(textContent);
  if (mdTable && mdTable.dataRows.length > 0) {
    const firstLine = textContent
      .split("\n")
      .map(l => l.trim())
      .find(l => l.length > 0 && !l.startsWith("|") && !isConversationalPreamble(l));
    const cleanTitle = cleanSlideTitle(firstLine || "Comparison Table", 1) || "Comparison Table";
    return [{
      slideNumber: 1,
      title: cleanTitle,
      subtitle: "",
      visualConcept: "",
      color: null,
      titleSize: 36,
      subtitleSize: 20,
      body: mdTable.bulletText,
      tableData: {
        headers: mdTable.headers,
        rows: mdTable.dataRows
      },
      base64Images: allImages
    }];
  }

  // -------------------------------------------------------------
  // Strategy 5: Raw Text Block Splitting
  // -------------------------------------------------------------
  const hasPrimaryHeaders = /(?:^|\n)#{1,2}\s+/i.test(textContent);
  const splitRegex = hasPrimaryHeaders
    ? /(?:^|\n)(?=(?:#{1,2}\s+|Slide\s*\d+[:\-]|(?:\d+\.\s+\*\*Slide)))/gi
    : /(?:^|\n)(?=(?:#{1,3}\s+|Slide\s*\d+[:\-]|(?:\d+\.\s+\*\*Slide)))/gi;

  const blocks = textContent
    .split(splitRegex)
    .map(b => b.trim())
    .filter(b => {
      if (b.length < 5) return false;
      if (isConversationalPreamble(b) && !b.includes("\n#") && !b.includes("\n•")) {
        return false;
      }
      return true;
    });

  if (blocks.length >= 1) {
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
        if (!title) return; // Skip conversational intro blocks

        const blockMdTable = parseMarkdownTable(block);
        const parsed = extractSlideMetadataAndBullets(lines.slice(1));

        textSlides.push({
          slideNumber: textSlides.length + 1,
          title: title,
          subtitle: parsed.subtitle,
          visualConcept: parsed.visualConcept,
          color: parsed.color,
          titleSize: parsed.titleSize,
          subtitleSize: parsed.subtitleSize,
          body: blockMdTable ? blockMdTable.bulletText : parsed.body,
          tableData: blockMdTable ? { headers: blockMdTable.headers, rows: blockMdTable.dataRows } : null,
          base64Images: allImages[idx] ? [allImages[idx]] : []
        });
      }
    });

    if (textSlides.length >= 1) return textSlides;
  }

  // -------------------------------------------------------------
  // Strategy 6: Single Slide Fallback
  // -------------------------------------------------------------
  const allLines = textContent.split("\n").map(l => l.trim()).filter(Boolean);
  const cleanLines = allLines.filter(l => !isConversationalPreamble(l));

  let titleIndex = -1;
  let singleTitle = "Executive Briefing";

  // Check for a heading line (# or ##)
  for (let i = 0; i < cleanLines.length; i++) {
    const l = cleanLines[i];
    if (l.startsWith("#")) {
      const candidate = cleanSlideTitle(l, 1);
      if (candidate) {
        singleTitle = candidate;
        titleIndex = i;
        break;
      }
    }
  }

  // If no # heading, pick first non-bullet line that is a valid title
  if (titleIndex === -1) {
    for (let i = 0; i < cleanLines.length; i++) {
      const l = cleanLines[i];
      if (!l.startsWith("•") && !l.startsWith("-") && !l.startsWith("*") && !l.startsWith("|")) {
        const candidate = cleanSlideTitle(l, 1);
        if (candidate) {
          singleTitle = candidate;
          titleIndex = i;
          break;
        }
      }
    }
  }

  const remainingLines = titleIndex >= 0 
    ? cleanLines.filter((_, idx) => idx !== titleIndex)
    : cleanLines;

  const parsed = extractSlideMetadataAndBullets(remainingLines);

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

export function cleanSlideTitle(rawTitle, defaultNum = 1) {
  if (!rawTitle) return `Slide ${defaultNum}`;
  if (isConversationalPreamble(rawTitle)) return null;

  let clean = rawTitle
    .replace(/^[#*\s:]+/, "")
    .replace(/^\d+[\.\)]\s*/, "")
    .replace(/^Slide\s*\d+[:\-–—]?\s*/i, "")
    .replace(/^(?:of\s+course[.,]?\s*|sure[.,]?\s*|certainly[.,]?\s*|absolutely[.,]?\s*)/i, "")
    .replace(/^(?:here\s+(?:is|are)\s+(?:the|your|a)?\s*(?:content|slide|slides|presentation)?\s*(?:for|of|on|about)?|below\s+(?:is|are)\s+(?:the|your|a)?\s*(?:content|slide|slides)?\s*(?:for|of|on|about)?)\s*/i, "")
    .replace(/^(?:a\s+table\s+of|a\s+comparison\s+of|table\s+of|comparison\s+of)\s*/i, "")
    .replace(/[:.]+$/, "")
    .replace(/\*\*/g, "")
    .trim();

  if (!clean || isConversationalPreamble(clean) || /^(?:the\s+)?slide$/i.test(clean)) {
    return null;
  }

  // Enforce maximum 5 words if emoji, else 4 words (under 40 chars)
  const words = clean.split(/\s+/);
  if (words.length > 5) {
    const hasEmoji = /^\p{Extended_Pictographic}/u.test(words[0]);
    const maxWords = hasEmoji ? 5 : 4;
    if (words.length > maxWords) {
      clean = words.slice(0, maxWords).join(" ");
    }
  }
  if (clean.length > 40) {
    clean = clean.substring(0, 40).replace(/\s+\S*$/, "").trim();
  }

  clean = clean.replace(/[:.\-–—\s]+$/, "").trim();
  return clean || null;
}
