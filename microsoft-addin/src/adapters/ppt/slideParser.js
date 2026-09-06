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
  let takeaway = "";
  let visualConcept = "";
  let color = null;
  let titleSize = 36;
  let subtitleSize = 18;
  const contentBullets = [];

  for (const rawLine of rawLines) {
    if (!rawLine) continue;
    let line = rawLine.trim();
    if (!line) continue;

    // Check for Subtitle:
    // Matches "#### Subtitle", "### Subtitle", "Subtitle: ...", "*Subtitle:* ...", "**Subtitle:** ..."
    const subMatch = line.match(/^(?:#{3,6}\s*|(?:Sub-?title|Subtitle\s*Text):\s*)(.*)$/i);
    if (subMatch && !subtitle) {
      const cleanSub = subMatch[1].replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      if (cleanSub) {
        subtitle = cleanSub;
        continue;
      }
    }

    // Check for Takeaway / Key Takeaway:
    // Matches "_Takeaway: text_", "**Takeaway:** text", "Takeaway: text", "Key Takeaway: text"
    const takeawayMatch = line.match(/^[-•*]*\s*[_*`\s]*(?:Key\s*)?Takeaway[_*`\s]*:\s*(.*)$/i);
    if (takeawayMatch && !takeaway) {
      const cleanTakeaway = takeawayMatch[1].replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      if (cleanTakeaway) {
        takeaway = cleanTakeaway;
        continue;
      }
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

    // Clean bullet text:
    let cleanBullet = line
      .replace(/^[-•*]\s*/, "")
      .replace(/^#{1,6}\s*/, "")
      .trim();

    // Check if it's a Subtitle or Takeaway after stripping bullet marker
    if (/^(?:Sub-?title|Subtitle\s*Text):\s*/i.test(cleanBullet) && !subtitle) {
      subtitle = cleanBullet.replace(/^(?:Sub-?title|Subtitle\s*Text):\s*/i, "").replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      continue;
    }
    if (/^[_*`\s]*(?:Key\s*)?Takeaway[_*`\s]*:\s*/i.test(cleanBullet) && !takeaway) {
      takeaway = cleanBullet.replace(/^[_*`\s]*(?:Key\s*)?Takeaway[_*`\s]*:\s*/i, "").replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      continue;
    }

    // Detect section heading or subtitle line (e.g. ### Strategic Analysis and Key Metrics)
    if (!subtitle && (/^#{2,4}\s+/.test(line) || (!line.startsWith("-") && !line.startsWith("*") && !line.startsWith("•") && cleanBullet.length < 50 && !/[.!?]$/.test(cleanBullet)))) {
      subtitle = cleanBullet.replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      continue;
    }

    // Strip wrapping markdown italic/bold underscores or asterisks
    if ((cleanBullet.startsWith("_") && cleanBullet.endsWith("_") && cleanBullet.length > 2) ||
        (cleanBullet.startsWith("*") && cleanBullet.endsWith("*") && cleanBullet.length > 2)) {
      cleanBullet = cleanBullet.slice(1, -1).trim();
    }

    if (cleanBullet) {
      contentBullets.push(`•  ${cleanBullet}`);
    }
  }

  return {
    subtitle,
    takeaway,
    visualConcept,
    color,
    titleSize: titleSize || 36,
    subtitleSize: subtitleSize || 18,
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

  // Extract all images upfront from HTML DOM, markdown images, and data URIs
  const highResChartEl = tempDiv.querySelector(".rendered-chart-container img");
  const highResChartSrc = highResChartEl ? (highResChartEl.src || highResChartEl.getAttribute("src") || "") : "";

  const rawDomImages = Array.from(tempDiv.querySelectorAll("img"))
    .map(img => img.src || img.getAttribute("src") || "")
    .filter(s => s && s.length > 50);

  const combinedSearch = ((htmlContent || '') + ' ' + (rawText || ''));
  const hasDistinctNonChartImageIntent = /(?:photo|photograph|portrait|illustration|logo|camera|scenery|picture of|image of a)/i.test(combinedSearch);

  const allImages = [];
  if (highResChartSrc) {
    allImages.push(highResChartSrc);
    // If the request requested distinct non-chart images (e.g. photos, logos), include them as well
    if (hasDistinctNonChartImageIntent) {
      for (const img of rawDomImages) {
        if (img !== highResChartSrc && !allImages.includes(img)) {
          allImages.push(img);
        }
      }
    }
  } else {
    for (const img of rawDomImages) {
      if (!allImages.includes(img)) allImages.push(img);
    }
    const mdImgMatches = combinedSearch.match(/!\[.*?\]\(\s*<?(data:image\/[^;]+;base64,[A-Za-z0-9+/=]+|https?:\/\/[^\s\)>]+)/gi) || [];
    for (const m of mdImgMatches) {
      const u = m.replace(/^!\[.*?\]\(\s*<?/i, '').replace(/>?\s*$/i, '').trim();
      if (u && !allImages.includes(u)) {
        allImages.push(u);
      }
    }

    const rawDataMatches = combinedSearch.match(/data:image\/(?:png|jpeg|jpg|webp|gif);base64,[A-Za-z0-9+/=]{100,}/gi) || [];
    for (const d of rawDataMatches) {
      if (!allImages.includes(d)) {
        allImages.push(d);
      }
    }
  }

  // Clean citation callouts, action toolbars, and preview containers from slide body text
  const noteCallouts = tempDiv.querySelectorAll("blockquote, .note, .ppt-deck-preview-container, .response-actions-container, [style*='background-color:#f0f6ff']");
  noteCallouts.forEach(n => n.remove());

  // -------------------------------------------------------------
  // Strategy 0: Executive Visual JSON (3-Column Metric Grid, Before/After)
  // -------------------------------------------------------------
  const fullSearchText = (rawText || "") + "\n" + (tempDiv.innerText || "");
  const jsonBlocks = fullSearchText.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/gi) || [];
  let visualPayload = null;

  for (const block of jsonBlocks) {
    try {
      const cleanBlock = block.replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
      const parsed = JSON.parse(cleanBlock);
      if (parsed && (parsed.visualType === "metric_grid_3col" || parsed.visualType === "before_after")) {
        visualPayload = parsed;
        break;
      }
    } catch (_) {}
  }

  if (!visualPayload) {
    const rawJsonMatch = fullSearchText.match(/\{[\s\S]*?"visualType"\s*:\s*"(?:metric_grid_3col|before_after)"[\s\S]*?\}/);
    if (rawJsonMatch) {
      try {
        visualPayload = JSON.parse(rawJsonMatch[0]);
      } catch (_) {}
    }
  }

  if (visualPayload && (visualPayload.visualType === "metric_grid_3col" || visualPayload.visualType === "before_after")) {
    const isMetric = visualPayload.visualType === "metric_grid_3col";
    let formattedBody = "";
    if (isMetric) {
      formattedBody = (visualPayload.cards || []).map(c => `• ${c.metric || ''} ${c.title || ''}: ${(c.bullets || []).join('; ')}`).join('\n');
    } else {
      const beforeList = (visualPayload.before?.bullets || []).map(b => `  - ${b}`).join('\n');
      const afterList = (visualPayload.after?.bullets || []).map(b => `  - ${b}`).join('\n');
      formattedBody = `• 🔴 BEFORE: ${visualPayload.before?.title || 'Current State'}\n${beforeList}\n\n• 🟢 AFTER: ${visualPayload.after?.title || 'Target State'}\n${afterList}`;
    }

    return [{
      slideNumber: 1,
      title: visualPayload.title || "Executive Visual",
      subtitle: visualPayload.subtitle || "",
      visualType: visualPayload.visualType,
      visualData: visualPayload,
      body: formattedBody,
      base64Images: allImages
    }];
  }

  // -------------------------------------------------------------
  // Strategy 1: Explicit Slide Headings (H1, H2, H3)
  // -------------------------------------------------------------
  const headerEls = Array.from(tempDiv.querySelectorAll("h1, h2, h3"));
  
  // Check if there is an explicit multi-slide outline table or list in the document
  const hasOutlineTable = Array.from(tempDiv.querySelectorAll("table")).some(t => {
    const rows = Array.from(t.querySelectorAll("tr"));
    if (rows.length < 2) return false;
    const hText = (rows[0].innerText || rows[0].textContent || "").toLowerCase();
    return hText.includes("slide #") || hText.includes("slide number") || hText.includes("slide title") ||
      rows.some(r => Array.from(r.querySelectorAll("td, th")).some(c => /^slide\s*\d+/i.test(c.textContent.trim())));
  });

  const hasOutlineList = Array.from(tempDiv.querySelectorAll("li")).filter(li => 
    /^(?:Slide\s*\d+|\*\*Slide\s*\d+\*\*)\s*[:\-–—]/i.test((li.innerText || li.textContent || "").trim())
  ).length >= 2;

  // Activate Strategy 1 if multiple headers exist, OR if 1 header exists without a multi-slide outline table/list
  const shouldRunStrategy1 = headerEls.length >= 2 || (headerEls.length === 1 && !hasOutlineTable && !hasOutlineList);

  if (shouldRunStrategy1) {
    const slides = [];
    for (let i = 0; i < headerEls.length; i++) {
      const h = headerEls[i];
      const rawTitle = (h.innerText || h.textContent || "").trim();
      const title = cleanSlideTitle(rawTitle, i + 1);

      const allBodyLines = [];
      const additionalBodyLines = [];
      const sectionImgs = [];
      let sectionSubtitle = "";
      let sectionTableData = null;

      // If this is the first slide, collect any chart images or illustrations that appear before the first header
      if (i === 0) {
        let prev = h.previousElementSibling;
        while (prev) {
          const prevImgs = Array.from(prev.querySelectorAll("img"))
            .map(img => img.src || img.getAttribute("src") || "")
            .filter(s => s && s.length > 50);
          if (prev.tagName === "IMG") {
            const pSrc = prev.src || prev.getAttribute("src") || "";
            if (pSrc.length > 50) prevImgs.push(pSrc);
          }
          if (prevImgs.length > 0) sectionImgs.unshift(...prevImgs);
          prev = prev.previousElementSibling;
        }
      }

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

        // Process H4, H5, H6 immediately as Subtitle if not already assigned
        if (["H4", "H5", "H6"].includes(curr.tagName)) {
          const subText = (curr.innerText || curr.textContent || "")
            .replace(/^#{1,6}\s*/, "")
            .replace(/^(?:Sub-?title|Subtitle\s*Text):\s*/i, "")
            .replace(/^[#*_`\s]+|[#*_`\s]+$/g, "")
            .trim();
          if (subText && !sectionSubtitle) {
            sectionSubtitle = subText;
            curr = curr.nextElementSibling;
            continue;
          }
        }

        // Process tables inside a slide section
        if (curr.tagName === "TABLE") {
          const tbl = extractTableContent(curr);
          if (tbl && tbl.bullets.length > 0) {
            allBodyLines.push(...tbl.bullets);
            sectionTableData = {
              headers: tbl.headers,
              rows: tbl.dataRows
            };
          }
        } else if (curr.tagName === "UL" || curr.tagName === "OL") {
          Array.from(curr.querySelectorAll("li")).forEach(li => {
            const txt = (li.innerText || li.textContent || "").trim();
            if (txt) {
              const b = `• ${txt.replace(/^[-•*]\s*/, "")}`;
              allBodyLines.push(b);
              additionalBodyLines.push(b);
            }
          });
        } else {
          const txt = (curr.innerText || curr.textContent || "").trim();
          if (txt) {
            allBodyLines.push(txt);
            if (!/^(?:Subtitle|Color|Visual|Title Size|Subtitle Size):/i.test(txt)) {
              additionalBodyLines.push(txt);
            }
          }
        }
        curr = curr.nextElementSibling;
      }

      const parsedAll = extractSlideMetadataAndBullets(allBodyLines);
      const parsedAdditional = extractSlideMetadataAndBullets(additionalBodyLines);

      slides.push({
        slideNumber: i + 1,
        title: title,
        subtitle: sectionSubtitle || parsedAll.subtitle,
        takeaway: parsedAll.takeaway,
        visualConcept: parsedAll.visualConcept,
        color: parsedAll.color,
        titleSize: parsedAll.titleSize,
        subtitleSize: parsedAll.subtitleSize,
        body: parsedAll.body,
        additionalBody: (sectionTableData && parsedAdditional.body !== "• Executive slide content") ? parsedAdditional.body : "",
        tableData: sectionTableData,
        base64Images: sectionImgs.length > 0 ? sectionImgs : (allImages.length > 0 && i === 0 ? allImages : [])
      });
    }

    if (slides.length >= 1) {
      return consolidateExecutiveSummarySlides(slides, allImages);
    }
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
      return consolidateExecutiveSummarySlides(tableSlides, allImages);
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
    return consolidateExecutiveSummarySlides(outlineSlides, allImages);
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
      
      // Check if there is a rendered chart container or chart image in tempDiv
      const chartEl = tempDiv.querySelector(".rendered-chart-container, [data-chart-title]");
      if (chartEl && chartEl.getAttribute("data-chart-title")) {
        tableTitle = chartEl.getAttribute("data-chart-title");
      } else {
        const chartImg = tempDiv.querySelector("img[alt*='chart' i], img[alt]");
        if (chartImg && chartImg.getAttribute("alt")) {
          const altText = chartImg.getAttribute("alt").replace(/\s*\([^)]*chart\)/i, "").trim();
          if (altText && altText !== "Image" && altText !== "Generated Chart") {
            tableTitle = altText;
          }
        }
      }

      if (!tableTitle) {
        let prevEl = standaloneTable.previousElementSibling;
        while (prevEl) {
          const txt = (prevEl.innerText || prevEl.textContent || "").trim();
          if (txt && !txt.startsWith("Verified Sources")) {
            tableTitle = txt;
            break;
          }
          prevEl = prevEl.previousElementSibling;
        }
      }

      if (!tableTitle) {
        const firstLine = (tempDiv.innerText || tempDiv.textContent || "").split("\n").map(l => l.trim()).find(l => l.length > 0);
        if (firstLine && !firstLine.startsWith("|")) tableTitle = firstLine;
      }

      if (!tableTitle && tbl.headers.length > 0) {
        tableTitle = `${tbl.headers[0]} Comparison`;
      }

      const cleanTitle = cleanSlideTitle(tableTitle || "Comparison Table", 1);

      // Extract any extra commentary/bullets before or after standalone table
      const extraLines = [];
      Array.from(tempDiv.querySelectorAll("h1, h2, h3, h4, p, li")).forEach(el => {
        const txt = (el.innerText || el.textContent || "").trim();
        if (txt && !txt.startsWith("Verified Sources") && txt !== tableTitle && txt !== cleanTitle && !/^(?:Subtitle|Color|Visual):/i.test(txt)) {
          extraLines.push(txt);
        }
      });
      const parsedExtra = extractSlideMetadataAndBullets(extraLines);

      return [{
        slideNumber: 1,
        title: cleanTitle,
        subtitle: parsedExtra.subtitle || "",
        visualConcept: "",
        color: null,
        titleSize: 36,
        subtitleSize: 20,
        body: tbl.bulletText,
        additionalBody: parsedExtra.body !== "• Executive slide content" ? parsedExtra.body : "",
        tableData: {
          headers: tbl.headers,
          rows: tbl.dataRows
        },
        base64Images: allImages
      }];
    }
  }

  // Check for raw markdown table in text
  const textContent = tempDiv.innerText || tempDiv.textContent || rawText;
  const mdTable = parseMarkdownTable(textContent);
  if (mdTable && mdTable.dataRows.length > 0) {
    const textLines = textContent.split("\n").map(l => l.trim()).filter(Boolean);
    const firstLine = textLines.find(l => !l.startsWith("|"));
    const cleanTitle = cleanSlideTitle(firstLine || "Comparison Table", 1);
    const nonTableLines = textLines.slice(1).filter(l => !l.startsWith("|") && !l.endsWith("|") && l !== firstLine);
    const parsedExtra = extractSlideMetadataAndBullets(nonTableLines);
    return [{
      slideNumber: 1,
      title: cleanTitle,
      subtitle: parsedExtra.subtitle || "",
      visualConcept: "",
      color: null,
      titleSize: 36,
      subtitleSize: 20,
      body: mdTable.bulletText,
      additionalBody: parsedExtra.body !== "• Executive slide content" ? parsedExtra.body : "",
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
  const blocks = textContent
    .split(/(?:^|\n)(?=(?:#{1,3}\s+|Slide\s*\d+[:\-]|(?:\d+\.\s+\*\*Slide)))/gi)
    .map(b => b.trim())
    .filter(b => {
      if (b.length < 5) return false;
      // Skip conversational intro preambles like "Here is a comparison..." if they precede real headings
      if (/^(?:here\s+is\s+|below\s+is\s+|sure|certainly|i've\s+prepared)/i.test(b) && !b.includes("\n#") && !b.includes("\n•")) {
        return false;
      }
      return true;
    });

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
        const blockMdTable = parseMarkdownTable(block);
        const nonTableLines = lines.slice(1).filter(l => !l.startsWith("|") && !l.endsWith("|"));
        const parsed = extractSlideMetadataAndBullets(lines.slice(1));
        const parsedNonTable = extractSlideMetadataAndBullets(nonTableLines);

        textSlides.push({
          slideNumber: idx + 1,
          title: title,
          subtitle: parsed.subtitle,
          takeaway: parsed.takeaway,
          visualConcept: parsed.visualConcept,
          color: parsed.color,
          titleSize: parsed.titleSize,
          subtitleSize: parsed.subtitleSize,
          body: blockMdTable ? blockMdTable.bulletText : parsed.body,
          additionalBody: (blockMdTable && parsedNonTable.body !== "• Executive slide content") ? parsedNonTable.body : "",
          tableData: blockMdTable ? { headers: blockMdTable.headers, rows: blockMdTable.dataRows } : null,
          base64Images: allImages[idx] ? [allImages[idx]] : []
        });
      }
    });

    if (textSlides.length >= 2) return consolidateExecutiveSummarySlides(textSlides, allImages);
  }

  // -------------------------------------------------------------
  // Strategy 6: Single Slide Fallback
  // -------------------------------------------------------------
  const cloneDiv = tempDiv.cloneNode(true);
  cloneDiv.querySelectorAll("h1, h2, h3, p, li, tr, div, blockquote").forEach(el => {
    el.insertAdjacentText("afterend", "\n");
  });
  const fallbackRaw = (cloneDiv.textContent || cloneDiv.innerText || rawText || "").trim();
  const sanitizedFallback = sanitizeAiResponse(fallbackRaw);
  const lines = sanitizedFallback.split("\n").map(l => l.trim()).filter(Boolean);

  let singleTitle = "🎯 Key Highlights";
  let contentLines = lines;

  if (lines.length > 0) {
    const candidateTitle = cleanSlideTitle(lines[0], 1);
    if (candidateTitle && !isConversationalPreamble(lines[0])) {
      singleTitle = candidateTitle;
      contentLines = lines.slice(1);
    } else {
      contentLines = lines;
    }
  }

  const parsed = extractSlideMetadataAndBullets(contentLines);

  return [{
    slideNumber: 1,
    title: singleTitle,
    subtitle: parsed.subtitle,
    takeaway: parsed.takeaway,
    visualConcept: parsed.visualConcept,
    color: parsed.color,
    titleSize: parsed.titleSize,
    subtitleSize: parsed.subtitleSize,
    body: parsed.body,
    base64Images: allImages
  }];
}

/**
 * Consolidates multi-section executive summary slides into a single executive slide,
 * preserving all subtitles, bullets, tables, takeaways, and generated images.
 */
function consolidateExecutiveSummarySlides(slides, allImages = []) {
  if (!slides || slides.length === 0) return slides;
  const firstTitle = (slides[0].title || "").toLowerCase();
  const isExecutiveSummary = slides.length > 1 && (
    firstTitle.includes("executive slide summary") ||
    firstTitle.includes("executive summary") ||
    firstTitle.includes("slide summary") ||
    (firstTitle.includes("executive") && firstTitle.includes("summary")) ||
    (firstTitle.includes("summary") && !firstTitle.includes("slide 1") && !firstTitle.includes("slide #"))
  );

  if (!isExecutiveSummary) return slides;

  const merged = { ...slides[0] };
  const allBullets = [];

  if (merged.body && merged.body !== "• Executive slide content") {
    allBullets.push(merged.body);
  }
  if (merged.additionalBody && merged.additionalBody !== "• Executive slide content") {
    allBullets.push(merged.additionalBody);
  }

  for (let s = 1; s < slides.length; s++) {
    const sub = slides[s];
    if (!merged.subtitle && sub.title && !sub.title.toLowerCase().startsWith("slide")) {
      merged.subtitle = sub.title;
    }
    if (sub.body && sub.body !== "• Executive slide content") {
      allBullets.push(sub.body);
    }
    if (sub.additionalBody && sub.additionalBody !== "• Executive slide content") {
      allBullets.push(sub.additionalBody);
    }
    if (!merged.tableData && sub.tableData) {
      merged.tableData = sub.tableData;
    }
    if (sub.base64Images && sub.base64Images.length > 0) {
      if (!merged.base64Images) merged.base64Images = [];
      for (const img of sub.base64Images) {
        if (!merged.base64Images.includes(img)) {
          merged.base64Images.push(img);
        }
      }
    }
    if (!merged.takeaway && sub.takeaway) {
      merged.takeaway = sub.takeaway;
    }
  }

  if (allImages && allImages.length > 0) {
    if (!merged.base64Images) merged.base64Images = [];
    for (const img of allImages) {
      if (!merged.base64Images.includes(img)) {
        merged.base64Images.push(img);
      }
    }
  }

  // Deduplicate and structure bullets
  const uniqueBullets = [];
  const seen = new Set();
  for (const bBlock of allBullets) {
    const lines = bBlock.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    for (const line of lines) {
      const cleanLine = line.replace(/^[-•*]\s*/, "").trim();
      if (cleanLine.startsWith("|") || cleanLine.startsWith("---") || /^(?:Subtitle|Color|Visual):/i.test(cleanLine)) {
        continue;
      }
      if (cleanLine.toLowerCase().startsWith("the chart and table below") || cleanLine.toLowerCase().startsWith("here is the")) {
        continue;
      }
      if (cleanLine.length > 5 && !seen.has(cleanLine.toLowerCase())) {
        seen.add(cleanLine.toLowerCase());
        uniqueBullets.push(line.startsWith("•") ? line : `• ${cleanLine}`);
      }
    }
  }

  merged.body = uniqueBullets.join("\n\n").trim() || "• Executive slide content";
  merged.additionalBody = merged.body;
  return [merged];
}

export function sanitizeAiResponse(text) {
  if (!text) return "";
  let clean = text.trim();

  // 1. Remove conversational preamble line(s) at start
  clean = clean.replace(/^(?:here\s+(?:is|are)|below\s+(?:is|are)|sure[.,!]?|certainly[.,!]?|of\s+course[.,!]?|as\s+requested[.,:]?|optimized\s+version[.,:]?|revised\s+version[.,:]?|i've\s+created|in\s+summary[.,:]?)[^\n]*\n+/i, "");

  // 2. Truncate alternative options (keep only Option 1)
  const altMatch = clean.match(/\n+\s*(?:(?:or,?\s+(?:for\s+)?(?:an\s+)?(?:even\s+)?(?:more\s+)?(?:minimalist|concise|alternative|compact)[^\n]*)|(?:option\s*2\b[^\n]*)|(?:alternative\s*(?:option|\d)?\b[^\n]*))\s*[:\n]/i);
  if (altMatch && altMatch.index !== undefined) {
    clean = clean.substring(0, altMatch.index).trim();
  }

  // 3. Remove conversational closing remarks at the end
  clean = clean.replace(/\n+(?:hope\s+this\s+helps|let\s+me\s+know|feel\s+free\s+to|let\s+me\s+know\s+if)[^\n]*$/i, "");

  return clean.trim();
}

export function isConversationalPreamble(line) {
  if (!line) return false;
  const l = line.toLowerCase().trim();
  if (/^(?:here\s+(?:is|are)|below\s+(?:is|are)|sure|certainly|of\s+course|as\s+requested|optimized|revised|concise|punchy|in\s+summary|to\s+make)/i.test(l)) return true;
  if (l.includes("optimized for a presentation") || l.includes("version optimized") || l.includes("concise and punchy version") || l.includes("minimalist layout") || l.includes("presentation slide")) return true;
  if (/^[-•*]/.test(l)) return true; // Starts with bullet
  return false;
}

export function extractCleanBulletPoints(htmlContent, rawText = "") {
  let text = (rawText || htmlContent || "").trim();
  if (text.includes("<") && text.includes(">")) {
    const tmp = document.createElement("div");
    tmp.innerHTML = text;
    tmp.querySelectorAll("blockquote, .note, .ppt-deck-preview-container, .response-actions-container").forEach(el => el.remove());
    text = (tmp.textContent || tmp.innerText || "").trim();
  }

  const sanitized = sanitizeAiResponse(text);
  const lines = sanitized.split("\n").map(l => l.trim()).filter(Boolean);
  const bullets = [];

  for (const line of lines) {
    if (isConversationalPreamble(line)) continue;
    const cleanLine = line.replace(/^[-•*]\s*/, "").replace(/\*\*/g, "").trim();
    if (cleanLine.length > 0) {
      bullets.push(`• ${cleanLine}`);
    }
  }

  return bullets.length > 0 ? bullets.join("\n") : sanitized;
}

export function cleanSlideTitle(rawTitle, defaultNum = 1) {
  if (!rawTitle) return `Slide ${defaultNum}`;

  let clean = rawTitle
    .replace(/^[#*\s:]+/, "")
    .replace(/^\d+[\.\)]\s*/, "")
    .replace(/^Slide\s*\d+[:\-–—]?\s*/i, "")
    .replace(/^(?:of\s+course[.,]?\s*|sure[.,]?\s*|certainly[.,]?\s*|absolutely[.,]?\s*)/i, "")
    .replace(/^(?:here\s+is\s+(?:a\s+)?(?:table|comparison|list|breakdown|summary)?\s+(?:of|comparing|for)?|below\s+is\s+(?:a\s+)?(?:table|comparison|list|breakdown|summary)?\s+(?:of|comparing|for)?)\s*/i, "")
    .replace(/^(?:a\s+table\s+of|a\s+comparison\s+of|table\s+of|comparison\s+of)\s*/i, "")
    .replace(/[:.]+$/, "")
    .replace(/\*\*/g, "")
    .trim();

  // If preamble was detected on the cleaned string, but it contains a colon, extract candidate after colon
  if (isConversationalPreamble(clean)) {
    if (clean.includes(":")) {
      const parts = clean.split(":");
      const candidate = parts.slice(1).join(":").trim();
      if (candidate && !isConversationalPreamble(candidate)) {
        clean = candidate;
      } else {
        return `Slide ${defaultNum}`;
      }
    } else {
      return `Slide ${defaultNum}`;
    }
  }

  // Preserve meaningful titles up to 8 words or 60 chars
  const words = clean.split(/\s+/);
  if (words.length > 8) {
    clean = words.slice(0, 8).join(" ");
  }
  if (clean.length > 60) {
    clean = clean.substring(0, 60).replace(/\s+\S*$/, "").trim();
  }

  return clean || `Slide ${defaultNum}`;
}
