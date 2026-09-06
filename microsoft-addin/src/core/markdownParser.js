/**
 * Enterprise Markdown & Multimodal Visual Component Parser for Microsoft Office
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { renderChartHtml } from './chartRenderer.js';

/**
 * Strips internal developer and slide layout metadata lines
 * (Visual Concept, Color, Title Size, Subtitle Size, Layout) from raw LLM text
 * so that business users see only clean slide titles, subtitles, bullets, and tables.
 */
export function cleanSlideDisplayMarkdown(text) {
  if (!text) return "";
  const lines = text.split(/\r?\n/);
  const cleanLines = [];

  for (const rawLine of lines) {
    const stripped = rawLine.replace(/^[-•*]\s*/, "").replace(/\*\*/g, "").trim();
    if (
      /^(?:Visual(?:\s*Concept|\s*Description|\s*Prompt|\s*Idea)?|Image(?:\s*Prompt|\s*Concept|\s*Description)?):/i.test(stripped) ||
      /^(?:Color|Colour|Color\s*Scheme|Palette|Theme\s*Color):/i.test(stripped) ||
      /^(?:Title|Subtitle)\s*(?:Font\s*)?Size(?:\s*\(pt\))?:/i.test(stripped) ||
      /^(?:Slide\s*)?Layout:/i.test(stripped) ||
      /^(?:Design\s*)?Theme:/i.test(stripped)
    ) {
      continue;
    }
    cleanLines.push(rawLine);
  }

  return cleanLines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

/**
 * Renders structured executive visual JSON (3-column metric grid or before/after comparison)
 * into a rich visual HTML card directly inside the chat interface.
 */
export function renderExecutiveVisualHtml(jsonString) {
  let data;
  try {
    data = typeof jsonString === 'object' ? jsonString : JSON.parse(jsonString);
  } catch (_) {
    return null;
  }
  if (!data || !data.visualType) return null;

  if (data.visualType === "metric_grid_3col") {
    const cards = data.cards || [];
    const cardHtml = cards.map(c => `
      <div style="background:#f8fafc; border:1px solid #cbd5e1; border-radius:8px; padding:10px; display:flex; flex-direction:column; gap:4px; flex:1; min-width:140px; box-sizing:border-box;">
        <div style="font-size:22px; font-weight:800; color:#0284c7; line-height:1.1;">${c.metric || ''}</div>
        <div style="font-size:12px; font-weight:700; color:#0f172a;">${c.title || ''}</div>
        ${c.subtitle ? `<div style="font-size:10px; color:#64748b; font-style:italic;">${c.subtitle}</div>` : ''}
        <ul style="margin:6px 0 0 14px; padding:0; font-size:11px; color:#334155; line-height:1.35;">
          ${(c.bullets || []).map(b => `<li>${b}</li>`).join('')}
        </ul>
      </div>
    `).join('');

    return `
      <div class="rendered-visual-container" style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:12px; margin:14px 0; box-shadow:0 2px 8px rgba(0,0,0,0.04);">
        <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:4px;">
          <div style="font-size:13.5px; font-weight:700; color:#0f172a;">${data.title || 'Executive Performance Metrics'}</div>
          <span style="background:#e0f2fe; color:#0369a1; font-size:9.5px; font-weight:700; padding:2px 6px; border-radius:4px; text-transform:uppercase;">📊 Metric Grid</span>
        </div>
        ${data.subtitle ? `<div style="font-size:11px; color:#64748b; margin-bottom:10px; font-style:italic;">${data.subtitle}</div>` : ''}
        <div style="display:flex; gap:8px; flex-wrap:wrap;">
          ${cardHtml}
        </div>
      </div>
    `;
  }

  if (data.visualType === "before_after") {
    const before = data.before || { title: "Current State", bullets: [] };
    const after = data.after || { title: "Target State", bullets: [] };

    return `
      <div class="rendered-visual-container" style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:12px; margin:14px 0; box-shadow:0 2px 8px rgba(0,0,0,0.04);">
        <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:4px;">
          <div style="font-size:13.5px; font-weight:700; color:#0f172a;">${data.title || 'Operational Transformation'}</div>
          <span style="background:#fee2e2; color:#991b1b; font-size:9.5px; font-weight:700; padding:2px 6px; border-radius:4px; text-transform:uppercase;">⚖️ Comparison</span>
        </div>
        ${data.subtitle ? `<div style="font-size:11px; color:#64748b; margin-bottom:10px; font-style:italic;">${data.subtitle}</div>` : ''}
        <div style="display:flex; gap:10px; flex-direction:column;">
          <div style="background:#fff8f8; border:1px solid #fecaca; border-radius:8px; padding:10px;">
            <div style="font-size:11.5px; font-weight:700; color:#991b1b; margin-bottom:4px;">🔴 BEFORE: ${before.title || 'Current State'}</div>
            <ul style="margin:2px 0 0 14px; padding:0; font-size:11px; color:#450a0a; line-height:1.35;">
              ${(before.bullets || []).map(b => `<li>${b}</li>`).join('')}
            </ul>
          </div>
          <div style="background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px; padding:10px;">
            <div style="font-size:11.5px; font-weight:700; color:#166534; margin-bottom:4px;">🟢 AFTER: ${after.title || 'Target State'}</div>
            <ul style="margin:2px 0 0 14px; padding:0; font-size:11px; color:#052e16; line-height:1.35;">
              ${(after.bullets || []).map(b => `<li>${b}</li>`).join('')}
            </ul>
          </div>
        </div>
      </div>
    `;
  }

  return null;
}

export function parseMarkdown(text) {
  if (!text) return "";

  // Strip internal slide/layout metadata for clean display
  const cleanedText = cleanSlideDisplayMarkdown(text);

  // 1. Sanitize raw scripts and broken SVGs
  let sanitized = cleanedText.replace(/<script[\s\S]*?<\/script>/gi, '');
  sanitized = sanitized.replace(/<svg[\s\S]*?<\/svg>/gi, '');
  sanitized = sanitized.replace(/<!--[\s\S]*?-->/g, '');

  const visualTokens = [];

  // 2. Pre-extract existing HTML <img> or <div style="..."><img ...></div> blocks
  sanitized = sanitized.replace(/<div[^>]*>[\s\S]*?<img[^>]+>[\s\S]*?<\/div>/gi, (match) => {
    const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
    visualTokens.push(match);
    return `\n\n${token}\n\n`;
  });

  sanitized = sanitized.replace(/<img[^>]+>/gi, (match) => {
    const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
    const wrapped = `<div class="office-visual-image-container" style="margin:16px 0; text-align:center;">${match}</div>`;
    visualTokens.push(wrapped);
    return `\n\n${token}\n\n`;
  });

  // 3. Pre-extract Markdown images ![alt](url)
  sanitized = sanitized.replace(/!\[([^\]]*)\]\(\s*<?(https?:\/\/[^\s\)>]+|data:image\/[^\s\)>]+)>?\s*(?:"[^"]*")?\s*\)/g, (match, alt, url) => {
    const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
    const imgHtml = `<div class="office-visual-image-container" style="margin:16px 0; text-align:center;"><img src="${url}" alt="${alt || 'Image'}" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 2px 8px rgba(0,0,0,0.06);" /></div>`;
    visualTokens.push(imgHtml);
    return `\n\n${token}\n\n`;
  });

  // 3b. Pre-extract any standalone data:image URIs that were output directly without markdown or HTML tags
  sanitized = sanitized.replace(/(?:^|\n)(data:image\/(?:png|jpeg|jpg|webp|gif);base64,[A-Za-z0-9+/=]{100,})(?:\n|$)/gi, (match, dataUri) => {
    const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
    const imgHtml = `<div class="office-visual-image-container" style="margin:16px 0; text-align:center;"><img src="${dataUri}" alt="Generated Image" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 2px 8px rgba(0,0,0,0.06);" /></div>`;
    visualTokens.push(imgHtml);
    return `\n\n${token}\n\n`;
  });

  // Check if an image version is already present in visualTokens (e.g. high-resolution image generated by Gemini Enterprise)
  const hasGeneratedChartImage = visualTokens.some(tok => tok.includes('<img'));

  // 4. Pre-extract and render structured Chart JSON blocks (Pie, Doughnut, Bar, Column, Line)
  // RULE: If an image version is already present, prioritize that image version and omit the lower-resolution duplicate client canvas chart.
  sanitized = sanitized.replace(/```(?:json|chart|pie|bar|line|doughnut|donut|column|vega|vega-lite)?\s*([\s\S]*?)```/gi, (fullMatch, codeContent) => {
    const trimmedCode = (codeContent || '').trim();
    if (trimmedCode.startsWith('{') && (
      trimmedCode.includes('chartType') ||
      trimmedCode.includes('chart_type') ||
      trimmedCode.includes('"pie"') ||
      trimmedCode.includes('"bar"') ||
      trimmedCode.includes('"doughnut"') ||
      trimmedCode.includes('"line"') ||
      (trimmedCode.includes('"data"') && trimmedCode.includes('"value"'))
    )) {
      try {
        const chartHtml = renderChartHtml(trimmedCode);
        if (chartHtml) {
          const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
          visualTokens.push(chartHtml);
          return `\n\n${token}\n\n`;
        }
      } catch (e) {
        console.warn("Client chart rendering failed:", e);
      }
    }
    return fullMatch;
  });

  // 5. Pre-extract standalone JSON chart objects without code fences
  sanitized = sanitized.replace(/\{\s*"(?:chartType|chart_type|type)"\s*:\s*"(?:pie|bar|line|doughnut|donut|column)"[\s\S]*?\n\s*\}/gi, (match) => {
    if (hasGeneratedChartImage) {
      return '';
    }
    try {
      const chartHtml = renderChartHtml(match);
      if (chartHtml) {
        const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
        visualTokens.push(chartHtml);
        return `\n\n${token}\n\n`;
      }
    } catch (_) {}
    return match;
  });

  // 6. Pre-extract and render structured Executive Visual JSON blocks (3-Column Metric Grid or Before/After)
  sanitized = sanitized.replace(/```(?:json)?\s*(\{[\s\S]*?"visualType"\s*:\s*"(?:metric_grid_3col|before_after)"[\s\S]*?\})\s*```/gi, (fullMatch, codeContent) => {
    try {
      const visualHtml = renderExecutiveVisualHtml(codeContent);
      if (visualHtml) {
        const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
        visualTokens.push(visualHtml);
        return `\n\n${token}\n\n`;
      }
    } catch (_) {}
    return fullMatch;
  });

  // 7. Pre-extract standalone Executive Visual JSON objects without code fences
  sanitized = sanitized.replace(/\{\s*"visualType"\s*:\s*"(?:metric_grid_3col|before_after)"[\s\S]*?\n\s*\}/gi, (match) => {
    try {
      const visualHtml = renderExecutiveVisualHtml(match);
      if (visualHtml) {
        const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
        visualTokens.push(visualHtml);
        return `\n\n${token}\n\n`;
      }
    } catch (_) {}
    return match;
  });


  // 3. Block-level parsing (Headings, Lists, Tables, Blockquotes, Paragraphs)
  const lines = sanitized.split(/\r?\n/);
  const outputBlocks = [];
  let currentList = null; // 'ul' or 'ol'
  let listItems = [];
  let currentTable = [];
  let currentBlockquote = [];

  function flushList() {
    if (currentList && listItems.length > 0) {
      const tag = currentList;
      const listHtml = `<${tag} style="font-family:'Segoe UI',Calibri,sans-serif; font-size:11pt; color:#201f1e; line-height:1.6; margin:10px 0; padding-left:24px;">` +
        listItems.map(item => `<li style="margin-bottom:6px;">${parseInline(item)}</li>`).join('') +
        `</${tag}>`;
      outputBlocks.push(listHtml);
      currentList = null;
      listItems = [];
    }
  }

  function flushTable() {
    if (currentTable.length > 0) {
      let tableHtml = `<table style="width:100%; border-collapse:collapse; margin:16px 0; font-family:'Segoe UI',Calibri,sans-serif; font-size:10pt; box-shadow:0 1px 3px rgba(0,0,0,0.05); border:1px solid #c7e0f4;">`;
      const isHeader = currentTable.length > 1 && currentTable[1].every(cell => /^[-:]+$/.test(cell.trim()));
      
      currentTable.forEach((row, rowIndex) => {
        if (rowIndex === 1 && isHeader) return; // Skip separator line
        tableHtml += `<tr>`;
        row.forEach(cell => {
          if (rowIndex === 0 && isHeader) {
            tableHtml += `<th style="background-color:#0f6cbd; color:#ffffff; padding:10px 14px; font-weight:600; text-align:left; border:1px solid #0f6cbd; font-size:10.5pt;">${parseInline(cell.trim())}</th>`;
          } else {
            tableHtml += `<td style="padding:9px 13px; border:1px solid #e1dfdd; color:#201f1e; vertical-align:top; line-height:1.4;">${parseInline(cell.trim())}</td>`;
          }
        });
        tableHtml += `</tr>`;
      });
      tableHtml += `</table>`;
      outputBlocks.push(tableHtml);
      currentTable = [];
    }
  }

  function flushBlockquote() {
    if (currentBlockquote.length > 0) {
      const fullText = currentBlockquote.join(' ');
      let borderColor = "#0078d4";
      let bgColor = "#eff6fc";
      let textColor = "#004e8c";
      let icon = "💡";
      let cleanText = fullText;

      if (cleanText.includes("[!WARNING]") || cleanText.toLowerCase().includes("risk") || cleanText.toLowerCase().includes("caution")) {
        borderColor = "#d83b01";
        bgColor = "#fde7e9";
        textColor = "#a80000";
        icon = "⚠️";
        cleanText = cleanText.replace(/\[!WARNING\]/gi, "").replace(/\[!CAUTION\]/gi, "");
      } else if (cleanText.includes("[!TIP]") || cleanText.toLowerCase().includes("recommendation") || cleanText.toLowerCase().includes("action")) {
        borderColor = "#107c10";
        bgColor = "#dff6dd";
        textColor = "#0e5c0e";
        icon = "✅";
        cleanText = cleanText.replace(/\[!TIP\]/gi, "");
      } else {
        cleanText = cleanText.replace(/\[!NOTE\]/gi, "").replace(/\[!IMPORTANT\]/gi, "");
      }

      const cardHtml = `<table style="width:100%; background-color:${bgColor}; border-left:5px solid ${borderColor}; border-top:1px solid #d2d0ce; border-right:1px solid #d2d0ce; border-bottom:1px solid #d2d0ce; margin:14px 0; border-collapse:collapse; border-radius:4px;">
        <tr>
          <td style="padding:12px 16px; font-family:'Segoe UI',Calibri,sans-serif; font-size:11pt; color:${textColor}; border:none; line-height:1.5;">
            <div style="font-weight:600; margin-bottom:4px;">${icon} Executive Callout</div>
            ${parseInline(cleanText.trim())}
          </td>
        </tr>
      </table>`;
      outputBlocks.push(cardHtml);
      currentBlockquote = [];
    }
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    // Check for Visual Token
    if (trimmed.startsWith('%%OFFICE_VISUAL_TOKEN_') && trimmed.endsWith('%%')) {
      flushList();
      flushTable();
      flushBlockquote();
      const tokenIdx = parseInt(trimmed.replace(/\D/g, ''), 10);
      if (!isNaN(tokenIdx) && visualTokens[tokenIdx]) {
        outputBlocks.push(visualTokens[tokenIdx]);
      }
      continue;
    }

    // Check for Horizontal Rule
    if (/^(\*\*\*|---|___)$/.test(trimmed)) {
      flushList();
      flushTable();
      flushBlockquote();
      outputBlocks.push('<hr style="border:none; border-top:1px solid #edebe9; margin:16px 0;" />');
      continue;
    }

    // Check for Headings
    if (/^#{1,3}\s+/.test(trimmed)) {
      flushList();
      flushTable();
      flushBlockquote();
      if (trimmed.startsWith('# ')) {
        outputBlocks.push(`<h1 style="font-family:'Segoe UI',Calibri,sans-serif; font-size:18pt; font-weight:700; color:#004e8c; margin-top:20px; margin-bottom:10px; border-bottom:2px solid #0078d4; padding-bottom:6px;">${parseInline(trimmed.substring(2))}</h1>`);
      } else if (trimmed.startsWith('## ')) {
        outputBlocks.push(`<h2 style="font-family:'Segoe UI',Calibri,sans-serif; font-size:14pt; font-weight:600; color:#0f6cbd; margin-top:16px; margin-bottom:8px; border-bottom:1px solid #c7e0f4; padding-bottom:4px;">${parseInline(trimmed.substring(3))}</h2>`);
      } else if (trimmed.startsWith('### ')) {
        outputBlocks.push(`<h3 style="font-family:'Segoe UI',Calibri,sans-serif; font-size:12pt; font-weight:600; color:#115ea3; margin-top:12px; margin-bottom:6px;">${parseInline(trimmed.substring(4))}</h3>`);
      }
      continue;
    }

    // Check for Blockquotes
    if (trimmed.startsWith('>')) {
      flushList();
      flushTable();
      currentBlockquote.push(trimmed.replace(/^>\s*/, ''));
      continue;
    } else if (currentBlockquote.length > 0) {
      flushBlockquote();
    }

    // Check for Markdown Table Rows
    if (trimmed.startsWith('|') && trimmed.endsWith('|') && trimmed.includes('|')) {
      flushList();
      flushBlockquote();
      const cells = trimmed.slice(1, -1).split('|');
      currentTable.push(cells);
      continue;
    } else if (currentTable.length > 0) {
      flushTable();
    }

    // Check for Bullet Lists
    const bulletMatch = trimmed.match(/^[-*]\s+(.*)$/);
    if (bulletMatch) {
      flushTable();
      flushBlockquote();
      if (currentList !== 'ul') flushList();
      currentList = 'ul';
      listItems.push(bulletMatch[1]);
      continue;
    }

    // Check for Numbered Lists
    const numMatch = trimmed.match(/^\d+\.\s+(.*)$/);
    if (numMatch) {
      flushTable();
      flushBlockquote();
      if (currentList !== 'ol') flushList();
      currentList = 'ol';
      listItems.push(numMatch[1]);
      continue;
    }

    // Blank line
    if (!trimmed) {
      flushList();
      flushTable();
      flushBlockquote();
      continue;
    }

    // Visual Tokens (Images / Charts)
    const tokenMatch = trimmed.match(/^%%OFFICE_VISUAL_TOKEN_(\d+)%%$/);
    if (tokenMatch) {
      flushList();
      flushTable();
      flushBlockquote();
      const tokenIdx = parseInt(tokenMatch[1], 10);
      if (visualTokens[tokenIdx]) {
        outputBlocks.push(visualTokens[tokenIdx]);
      }
      continue;
    }

    // Regular Paragraph
    flushList();
    flushTable();
    flushBlockquote();
    outputBlocks.push(`<p style="font-family:'Segoe UI',Calibri,sans-serif; font-size:11pt; color:#201f1e; line-height:1.6; margin:8px 0;">${parseInline(trimmed)}</p>`);
  }

  flushList();
  flushTable();
  flushBlockquote();

  let htmlResult = `<div style="font-family:'Segoe UI','Segoe UI Web',Calibri,sans-serif; font-size:11pt; color:#201f1e; line-height:1.6;">${outputBlocks.join('')}</div>`;

  // Final pass for any inline tokens
  visualTokens.forEach((tokenHtml, idx) => {
    htmlResult = htmlResult.replaceAll(`%%OFFICE_VISUAL_TOKEN_${idx}%%`, tokenHtml);
  });

  return htmlResult;
}

function parseInline(text) {
  if (!text) return '';
  return text
    .replace(/\*\*([^*]+)\*\*/g, '<strong style="color:#111827; font-weight:600;">$1</strong>')
    .replace(/\*([^*]+)\*/g, '<em>$1</em>')
    .replace(/`([^`]+)`/g, '<code style="background-color:#f3f2f1; padding:2px 5px; border-radius:3px; font-family:Consolas,monospace; font-size:10pt; color:#a4262c;">$1</code>');
}

function renderOfficeStoryboard(title, subtitle, panels) {
  const count = Math.min(panels.length, 4);
  const borderColors = ['#0078d4', '#0f6cbd', '#107c10', '#d83b01'];
  const bgColors = ['#f0f6ff', '#f3f8fc', '#f0f9f0', '#fff8f0'];
  const defaultIcons = ['🎬', '⚡', '🚀', '💡'];

  let cells = '';
  panels.slice(0, 4).forEach((p, i) => {
    const border = p.color || borderColors[i % borderColors.length];
    const bg = bgColors[i % bgColors.length];
    const icon = p.icon || defaultIcons[i % defaultIcons.length];
    const dialogueText = p.dialogue ? `&ldquo;${escapeXml(p.dialogue)}&rdquo;` : escapeXml(p.caption || p.description || '');

    cells += `
      <td bgcolor="${bg}" style="width:${Math.floor(100 / count)}%; padding:8px; vertical-align:top; border:2px solid ${border}; background-color:${bg}; text-align:center;">
        <table style="width:100%; border-collapse:collapse; margin:0;" cellpadding="0" cellspacing="0">
          <tr>
            <td bgcolor="${border}" style="background-color:${border}; color:#ffffff; font-weight:700; font-size:9pt; padding:3px 6px; text-align:center; font-family:'Segoe UI',Calibri,sans-serif;">
              PANEL ${i + 1}: ${escapeXml(p.title || 'Scene')}
            </td>
          </tr>
          <tr>
            <td style="font-size:24pt; line-height:1.2; padding:10px 0; text-align:center;">${icon}</td>
          </tr>
          <tr>
            <td bgcolor="#ffffff" style="background-color:#ffffff; border:1px solid #e1dfdd; padding:6px 8px; font-size:9pt; color:#201f1e; line-height:1.4; text-align:left; font-style:italic; font-family:'Segoe UI',Calibri,sans-serif;">
              ${dialogueText}
            </td>
          </tr>
        </table>
      </td>`;
  });

  return `<table style="width:100%; border-collapse:separate; border-spacing:8px; margin:16px 0; font-family:'Segoe UI',Calibri,sans-serif;" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="${count}" style="padding:4px 0 8px 0; border:none;">
          <div style="font-size:13pt; font-weight:700; color:#004e8c;">🎨 ${escapeXml(title || 'Visual Storyboard')}</div>
          ${subtitle ? `<div style="font-size:10pt; color:#605e5c;">${escapeXml(subtitle)}</div>` : ''}
        </td>
      </tr>
      <tr>${cells}</tr>
    </table>`;
}

function renderOfficeBarChart(title, subtitle, data) {
  const maxVal = Math.max(...data.map(d => Number(d.value) || 0), 1);
  const colors = ['#004e8c', '#0f6cbd', '#2b88d8', '#0078d4', '#107c10', '#5c2d91', '#605e5c'];

  let rows = '';
  data.forEach((item, index) => {
    const val = Number(item.value) || 0;
    const pct = Math.max(5, Math.min(100, Math.round((val / maxVal) * 100)));
    const emptyPct = 100 - pct;
    const color = item.color || colors[index % colors.length];
    const displayVal = item.displayValue || (item.unit ? `${val} ${item.unit}` : (val <= 100 ? `${val}%` : val.toLocaleString()));

    rows += `
      <tr style="border-bottom:1px solid #edebe9;">
        <td style="width:28%; padding:8px 10px; font-weight:600; color:#201f1e; font-size:10pt; vertical-align:middle;">${escapeXml(item.label)}</td>
        <td style="width:56%; padding:8px 10px; vertical-align:middle;">
          <table style="width:100%; border-collapse:collapse; border:none; margin:0;" cellpadding="0" cellspacing="0" border="0">
            <tr style="height:16px;">
              <td bgcolor="${color}" style="background-color:${color}; width:${pct}%; height:16px; font-size:1px; line-height:1px; border:none;">&nbsp;</td>
              ${emptyPct > 0 ? `<td bgcolor="#f3f2f1" style="background-color:#f3f2f1; width:${emptyPct}%; height:16px; font-size:1px; line-height:1px; border:none;">&nbsp;</td>` : ''}
            </tr>
          </table>
        </td>
        <td style="width:16%; padding:8px 10px; text-align:right; font-weight:700; color:#004e8c; font-size:10pt; vertical-align:middle;">${escapeXml(displayVal)}</td>
      </tr>`;
  });

  return `<table style="width:100%; border-collapse:collapse; margin:16px 0; font-family:'Segoe UI',Calibri,sans-serif; background-color:#ffffff; border:1px solid #c7e0f4;" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="3" bgcolor="#f8fafd" style="background-color:#f8fafd; padding:12px 14px; border-bottom:1px solid #c7e0f4;">
          <div style="font-size:12pt; font-weight:700; color:#004e8c;">📊 ${escapeXml(title)}</div>
          ${subtitle ? `<div style="font-size:9.5pt; color:#605e5c; margin-top:2px;">${escapeXml(subtitle)}</div>` : ''}
        </td>
      </tr>
      ${rows}
    </table>`;
}

function renderOfficeProcessFlow(title, subtitle, data) {
  const count = Math.min(data.length, 5);
  let cells = '';

  data.slice(0, 5).forEach((step, i) => {
    const color = i === 0 ? '#004e8c' : (i === count - 1 ? '#107c10' : '#0f6cbd');
    cells += `
      <td bgcolor="#f8fafd" style="width:${Math.floor(100 / count)}%; padding:10px 8px; vertical-align:top; border:2px solid ${color}; background-color:#f8fafd; text-align:center;">
        <div style="font-size:9pt; font-weight:700; color:${color}; margin-bottom:4px;">STEP ${i + 1}</div>
        <div style="font-size:10pt; font-weight:600; color:#201f1e; margin-bottom:4px;">${escapeXml(step.label)}</div>
        ${step.description ? `<div style="font-size:8.5pt; color:#605e5c;">${escapeXml(step.description)}</div>` : ''}
      </td>`;
    if (i < count - 1) {
      cells += `<td style="width:20px; text-align:center; font-weight:bold; color:#0f6cbd; font-size:14pt; vertical-align:middle; border:none;">➔</td>`;
    }
  });

  return `<table style="width:100%; border-collapse:collapse; margin:16px 0; background-color:#ffffff; border:1px solid #c7e0f4; padding:12px; font-family:'Segoe UI',Calibri,sans-serif;" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="${count * 2 - 1}" bgcolor="#f8fafd" style="background-color:#f8fafd; padding:10px 12px; border-bottom:1px solid #c7e0f4;">
          <div style="font-size:12pt; font-weight:700; color:#004e8c;">⚡ ${escapeXml(title)}</div>
          ${subtitle ? `<div style="font-size:9.5pt; color:#605e5c;">${escapeXml(subtitle)}</div>` : ''}
        </td>
      </tr>
      <tr><td colspan="${count * 2 - 1}" style="height:8px; border:none;"></td></tr>
      <tr>${cells}</tr>
    </table>`;
}

function renderOfficeKpiGrid(title, subtitle, data) {
  const count = Math.min(data.length, 4);
  let cells = '';

  data.slice(0, 4).forEach((kpi) => {
    const color = kpi.color || '#0078d4';
    cells += `
      <td bgcolor="#f8fafd" style="width:${Math.floor(100 / count)}%; padding:12px 10px; vertical-align:top; border:1px solid #deecf9; border-top:4px solid ${color}; background-color:#f8fafd; text-align:center;">
        <div style="font-size:9pt; font-weight:600; color:#605e5c; margin-bottom:4px;">${escapeXml(kpi.label)}</div>
        <div style="font-size:16pt; font-weight:700; color:${color}; font-family:'Segoe UI',Calibri,sans-serif;">${escapeXml(String(kpi.value))}</div>
      </td>`;
  });

  return `<table style="width:100%; border-collapse:separate; border-spacing:8px; margin:16px 0; font-family:'Segoe UI',Calibri,sans-serif;" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="${count}" style="padding:4px 0 8px 0; border:none;">
          <div style="font-size:12pt; font-weight:700; color:#004e8c;">📈 ${escapeXml(title || 'Key Metrics')}</div>
          ${subtitle ? `<div style="font-size:9.5pt; color:#605e5c;">${escapeXml(subtitle)}</div>` : ''}
        </td>
      </tr>
      <tr>${cells}</tr>
    </table>`;
}

function escapeXml(unsafe) {
  if (!unsafe) return '';
  return String(unsafe)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
