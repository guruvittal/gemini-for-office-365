/**
 * Enterprise Markdown & Multimodal Visual Component Parser for Microsoft Office
 * 
 * @author Sathya AG, Principal Architect, Google
 */
export function parseMarkdown(text) {
  if (!text) return "";

  // 1. Sanitize raw scripts and broken SVGs
  let sanitized = text.replace(/<script[\s\S]*?<\/script>/gi, '');
  sanitized = sanitized.replace(/<svg[\s\S]*?<\/svg>/gi, '');
  sanitized = sanitized.replace(/<!--[\s\S]*?-->/g, '');

  const visualTokens = [];

  // 2. Pre-extract existing HTML <img> or <div style="..."><img ...></div> blocks
  sanitized = sanitized.replace(/<div[^>]*><img[^>]*><\/div>/gi, (match) => {
    const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
    visualTokens.push(match);
    return `\n\n${token}\n\n`;
  });

  sanitized = sanitized.replace(/<img[^>]+>/gi, (match) => {
    const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
    const wrapped = `<div style="margin:16px 0; text-align:center;">${match}</div>`;
    visualTokens.push(wrapped);
    return `\n\n${token}\n\n`;
  });

  // 3. Pre-extract Markdown images ![alt](url)
  sanitized = sanitized.replace(/!\[([^\]]*)\]\((https?:\/\/[^\s\)]+|data:image\/[^\s\)]+)\)/g, (match, alt, url) => {
    const token = `%%OFFICE_VISUAL_TOKEN_${visualTokens.length}%%`;
    const imgHtml = `<div style="margin:16px 0; text-align:center;"><img src="${url}" alt="${alt || 'Image'}" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 2px 8px rgba(0,0,0,0.06);" /></div>`;
    visualTokens.push(imgHtml);
    return `\n\n${token}\n\n`;
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
