/**
 * Client-Side Canvas Chart Rendering Engine for Microsoft Office Add-in
 * 
 * Dynamically parses structured chart JSON, Vega-Lite specs, and tabular data
 * and renders high-DPI corporate charts (Pie, Doughnut, Bar, Column, Line)
 * to Base64 PNG images for taskpane display and native PowerPoint/Word/Excel insertion.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

const PALETTE = [
  '#1a73e8', // Google Blue
  '#34a853', // Google Green
  '#fbbc04', // Google Yellow
  '#ea4335', // Google Red
  '#9334e6', // Purple
  '#12b5cb', // Teal / Cyan
  '#fa7b17', // Orange
  '#174ea6', // Navy Dark Blue
  '#e37400', // Amber
  '#a142f4'  // Light Purple
];

/**
 * Parses raw text or code block content to extract chart specification.
 * Supports:
 * 1. Standard chart JSON: { chartType: "pie", title: "...", data: [{ label, value }] }
 * 2. Variant keys: type, chart_type, labels/values, series/categories
 * 3. Array of objects: [{ label: "A", value: 10 }, ...]
 * @param {string} text - Raw string content
 * @returns {Object|null} Parsed chart spec or null
 */
export function parseChartSpec(text) {
  if (!text || typeof text !== 'string') return null;

  const trimmed = text.trim();

  // Try parsing direct JSON
  try {
    const parsed = JSON.parse(trimmed);
    return normalizeChartSpec(parsed);
  } catch (_) {}

  // Try extracting JSON from markdown code fence ```json ... ``` or ```chart ... ```
  const fenceMatch = trimmed.match(/```(?:json|chart|vega|vega-lite)?\s*([\s\S]*?)```/i);
  if (fenceMatch) {
    try {
      const parsed = JSON.parse(fenceMatch[1].trim());
      return normalizeChartSpec(parsed);
    } catch (_) {}
  }

  // Try searching for first { ... } block
  const jsonBlockMatch = trimmed.match(/\{[\s\S]*\}/);
  if (jsonBlockMatch) {
    try {
      const parsed = JSON.parse(jsonBlockMatch[0]);
      return normalizeChartSpec(parsed);
    } catch (_) {}
  }

  return null;
}

/**
 * Normalizes various JSON chart structures into a uniform schema:
 * {
 *   chartType: 'pie' | 'doughnut' | 'bar' | 'column' | 'line',
 *   title: string,
 *   subtitle: string,
 *   data: [{ label: string, value: number, color?: string }]
 * }
 */
export function normalizeChartSpec(obj) {
  if (!obj || typeof obj !== 'object') return null;

  // If top-level is an array: [{ label: "A", value: 10 }]
  if (Array.isArray(obj)) {
    const items = extractDataItems(obj);
    if (items.length > 0) {
      return {
        chartType: 'bar',
        title: 'Data Chart',
        subtitle: '',
        data: items
      };
    }
    return null;
  }

  const rawType = (obj.chartType || obj.chart_type || obj.type || obj.mark?.type || obj.mark || 'bar').toString().toLowerCase();
  let chartType = 'bar';
  if (rawType.includes('pie')) chartType = 'pie';
  else if (rawType.includes('doughnut') || rawType.includes('donut')) chartType = 'doughnut';
  else if (rawType.includes('line') || rawType.includes('trend')) chartType = 'line';
  else if (rawType.includes('column') || rawType.includes('vertical')) chartType = 'column';
  else if (rawType.includes('bar') || rawType.includes('horizontal')) chartType = 'bar';

  const title = obj.title?.text || obj.title || obj.name || 'Data Chart';
  const subtitle = obj.subtitle || obj.description || '';

  // Data can be in obj.data, obj.values, obj.items, obj.series
  let rawData = obj.data || obj.values || obj.items || obj.series || [];
  if (rawData && typeof rawData === 'object' && !Array.isArray(rawData)) {
    // Could be Vega-lite: { data: { values: [...] } }
    if (Array.isArray(rawData.values)) {
      rawData = rawData.values;
    } else if (Array.isArray(obj.labels) && Array.isArray(obj.values)) {
      // { labels: ["A", "B"], values: [10, 20] }
      rawData = obj.labels.map((l, i) => ({ label: l, value: obj.values[i] }));
    } else {
      // Key-value map: { "India": 1430000000, "China": 1410000000 }
      rawData = Object.entries(rawData).map(([k, v]) => ({ label: k, value: v }));
    }
  }

  // Handle { labels: [...], datasets: [{ data: [...] }] } (Chart.js format)
  if (Array.isArray(obj.labels) && Array.isArray(obj.datasets) && obj.datasets[0]?.data) {
    rawData = obj.labels.map((l, i) => ({ label: l, value: obj.datasets[0].data[i] }));
  }

  const dataItems = extractDataItems(rawData);
  if (dataItems.length === 0) return null;

  return {
    chartType,
    title: typeof title === 'string' ? title : 'Data Chart',
    subtitle: typeof subtitle === 'string' ? subtitle : '',
    data: dataItems
  };
}

/**
 * Extracts [{ label, value, color }] from diverse object shapes
 */
function extractDataItems(arr) {
  if (!Array.isArray(arr)) return [];

  const items = [];
  for (const item of arr) {
    if (!item) continue;
    if (typeof item !== 'object') continue;

    // Direct label and value
    const label = item.label || item.name || item.category || item.country || item.pillar || item.item || item.segment || item.x || Object.keys(item)[0];
    let rawVal = item.value !== undefined ? item.value : (item.val !== undefined ? item.val : (item.count !== undefined ? item.count : (item.amount !== undefined ? item.amount : (item.metric !== undefined ? item.metric : item.y))));
    if (rawVal === undefined) {
      const vals = Object.values(item);
      rawVal = vals.length > 1 ? vals[1] : vals[0];
    }


    const numVal = parseNumericValue(rawVal);
    if (label && !isNaN(numVal)) {
      items.push({
        label: String(label).trim(),
        value: numVal,
        color: item.color || null
      });
    }
  }
  return items;
}

function parseNumericValue(val) {
  if (typeof val === 'number') return val;
  if (!val) return NaN;
  const cleaned = String(val).replace(/[$,%BMk\s]/gi, '').trim();
  const num = parseFloat(cleaned);
  if (isNaN(num)) return NaN;
  if (String(val).toLowerCase().includes('b')) return num * 1e9;
  if (String(val).toLowerCase().includes('m')) return num * 1e6;
  if (String(val).toLowerCase().includes('k')) return num * 1e3;
  return num;
}

/**
 * Formats a number with friendly suffixes (e.g. 1.43B, 340M, 45K)
 */
export function formatValue(val) {
  if (val >= 1e9) return (val / 1e9).toFixed(val % 1e9 === 0 ? 0 : 2) + 'B';
  if (val >= 1e6) return (val / 1e6).toFixed(val % 1e6 === 0 ? 0 : 1) + 'M';
  if (val >= 1e3) return (val / 1e3).toFixed(val % 1e3 === 0 ? 0 : 1) + 'K';
  if (Number.isInteger(val)) return val.toLocaleString();
  return val.toFixed(1);
}

/**
 * Renders a chart specification onto an HTML5 Canvas and returns a Base64 PNG data URL.
 * Designed with a 2x Retina scale factor for razor-sharp presentation graphics.
 * @param {Object} spec - Normalized chart spec
 * @param {Object} [options] - Width, height, title
 * @returns {string|null} data:image/png;base64,... or null if canvas unsupported
 */
export function renderChartToDataUrl(spec, options = {}) {
  if (!spec || !spec.data || spec.data.length === 0) return null;
  if (typeof document === 'undefined' || !document.createElement) return null;

  const width = options.width || 800;
  const height = options.height || 500;
  const scale = 3; // Ultra-crisp 3x Retina scale for razor-sharp PowerPoint rendering

  const canvas = document.createElement('canvas');
  canvas.width = width * scale;
  canvas.height = height * scale;

  const ctx = canvas.getContext('2d');
  if (!ctx) return null;

  ctx.scale(scale, scale);

  // Background
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);

  // Header: Title & Subtitle
  ctx.fillStyle = '#1f1f1f';
  ctx.font = 'bold 26px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif';
  ctx.textAlign = 'left';
  ctx.fillText(spec.title || 'Data Visualization', 40, 44);

  if (spec.subtitle) {
    ctx.fillStyle = '#5f6368';
    ctx.font = 'italic 16px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif';
    ctx.fillText(spec.subtitle, 40, 72);
  }

  // Draw chart according to type
  const chartType = (spec.chartType || 'bar').toLowerCase();
  if (chartType === 'pie' || chartType === 'doughnut') {
    drawPieOrDoughnutChart(ctx, spec, width, height, chartType === 'doughnut');
  } else if (chartType === 'line') {
    drawLineChart(ctx, spec, width, height);
  } else if (chartType === 'column') {
    drawColumnChart(ctx, spec, width, height);
  } else {
    drawBarChart(ctx, spec, width, height);
  }

  return canvas.toDataURL('image/png');
}

/**
 * Draws Pie or Doughnut Chart with legend and percentages
 */
function drawPieOrDoughnutChart(ctx, spec, width, height, isDoughnut) {
  const data = spec.data;
  const total = data.reduce((sum, d) => sum + (d.value > 0 ? d.value : 0), 0) || 1;

  const startY = spec.subtitle ? 90 : 70;
  const availableHeight = height - startY - 30;
  const centerX = width * 0.38;
  const centerY = startY + availableHeight / 2;
  const radius = Math.min(centerX - 50, availableHeight / 2 - 16);

  let startAngle = -Math.PI / 2;

  // Draw Slices
  data.forEach((item, idx) => {
    const val = item.value > 0 ? item.value : 0;
    const sliceAngle = (val / total) * (2 * Math.PI);
    const endAngle = startAngle + sliceAngle;
    const color = item.color || PALETTE[idx % PALETTE.length];

    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, startAngle, endAngle);
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();

    // Slice separator
    ctx.lineWidth = 3;
    ctx.strokeStyle = '#ffffff';
    ctx.stroke();

    startAngle = endAngle;
  });

  // If Doughnut, cutout inner circle
  if (isDoughnut) {
    ctx.beginPath();
    ctx.arc(centerX, centerY, radius * 0.54, 0, 2 * Math.PI);
    ctx.fillStyle = '#ffffff';
    ctx.fill();

    // Inner total label
    ctx.fillStyle = '#202124';
    ctx.font = 'bold 22px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Total', centerX, centerY - 6);
    ctx.font = 'bold 18px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.fillStyle = '#5f6368';
    ctx.fillText(formatValue(total), centerX, centerY + 18);
  }

  // Draw Legend on Right Side
  const legendX = width * 0.66;
  const itemHeight = Math.min(42, Math.floor((availableHeight - 10) / Math.max(data.length, 1)));
  const legendStartY = centerY - (data.length * itemHeight) / 2;

  ctx.textAlign = 'left';

  data.forEach((item, idx) => {
    const y = legendStartY + idx * itemHeight + 12;
    const color = item.color || PALETTE[idx % PALETTE.length];
    const pct = ((item.value / total) * 100).toFixed(1);

    // Color Swatch
    ctx.beginPath();
    ctx.arc(legendX, y, 8, 0, 2 * Math.PI);
    ctx.fillStyle = color;
    ctx.fill();

    // Label
    ctx.fillStyle = '#202124';
    ctx.font = 'bold 16px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    const labelText = item.label.length > 20 ? item.label.substring(0, 18) + '…' : item.label;
    ctx.fillText(labelText, legendX + 18, y + 4);

    // Value & Pct
    ctx.fillStyle = '#5f6368';
    ctx.font = '14px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.fillText(`${formatValue(item.value)} (${pct}%)`, legendX + 18, y + 22);
  });
}

/**
 * Draws Horizontal Bar Chart
 */
function drawBarChart(ctx, spec, width, height) {
  const data = spec.data;
  const maxVal = Math.max(...data.map(d => d.value), 1);
  const startY = spec.subtitle ? 90 : 70;
  const availableHeight = height - startY - 40;
  const barHeight = Math.min(34, Math.floor((availableHeight - 10) / data.length) - 10);
  const leftMargin = 180;
  const maxBarWidth = width - leftMargin - 120;

  data.forEach((item, idx) => {
    const y = startY + idx * (barHeight + 12);
    const barWidth = Math.max(6, (item.value / maxVal) * maxBarWidth);
    const color = item.color || PALETTE[idx % PALETTE.length];

    // Label
    ctx.fillStyle = '#202124';
    ctx.font = 'bold 15px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'right';
    const labelText = item.label.length > 22 ? item.label.substring(0, 20) + '…' : item.label;
    ctx.fillText(labelText, leftMargin - 14, y + barHeight / 2 + 5);

    // Bar background track
    ctx.fillStyle = '#f1f3f4';
    drawRoundedRect(ctx, leftMargin, y, maxBarWidth, barHeight, 4);
    ctx.fill();

    // Bar fill
    ctx.fillStyle = color;
    drawRoundedRect(ctx, leftMargin, y, barWidth, barHeight, 4);
    ctx.fill();

    // Value label
    ctx.fillStyle = '#202124';
    ctx.font = 'bold 14px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'left';
    ctx.fillText(formatValue(item.value), leftMargin + barWidth + 10, y + barHeight / 2 + 5);
  });
}

/**
 * Draws Vertical Column Chart
 */
function drawColumnChart(ctx, spec, width, height) {
  const data = spec.data;
  const maxVal = Math.max(...data.map(d => d.value), 1);
  const startY = spec.subtitle ? 90 : 70;
  const bottomY = height - 50;
  const availableHeight = bottomY - startY;
  const leftMargin = 70;
  const availableWidth = width - leftMargin - 50;
  const colWidth = Math.min(50, Math.floor(availableWidth / data.length) - 16);
  const gap = (availableWidth - colWidth * data.length) / (data.length + 1);

  // Baseline
  ctx.strokeStyle = '#dadce0';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(leftMargin - 10, bottomY);
  ctx.lineTo(width - 40, bottomY);
  ctx.stroke();

  data.forEach((item, idx) => {
    const x = leftMargin + gap + idx * (colWidth + gap);
    const colHeight = Math.max(8, (item.value / maxVal) * availableHeight);
    const y = bottomY - colHeight;
    const color = item.color || PALETTE[idx % PALETTE.length];

    // Column fill
    ctx.fillStyle = color;
    drawRoundedRect(ctx, x, y, colWidth, colHeight, 4);
    ctx.fill();

    // Value above column
    ctx.fillStyle = '#202124';
    ctx.font = 'bold 13px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(formatValue(item.value), x + colWidth / 2, y - 8);

    // Label below column
    ctx.fillStyle = '#5f6368';
    ctx.font = '600 13px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    const labelText = item.label.length > 10 ? item.label.substring(0, 8) + '…' : item.label;
    ctx.fillText(labelText, x + colWidth / 2, bottomY + 22);
  });
}

/**
 * Draws Line Chart
 */
function drawLineChart(ctx, spec, width, height) {
  const data = spec.data;
  const maxVal = Math.max(...data.map(d => d.value), 1);
  const minVal = Math.min(...data.map(d => d.value), 0);
  const range = maxVal - minVal || 1;

  const startY = spec.subtitle ? 95 : 75;
  const bottomY = height - 55;
  const availableHeight = bottomY - startY;
  const leftMargin = 80;
  const rightMargin = 50;
  const availableWidth = width - leftMargin - rightMargin;
  const stepX = availableWidth / Math.max(data.length - 1, 1);

  // Horizontal Grid Lines
  ctx.strokeStyle = '#f1f3f4';
  ctx.lineWidth = 1;
  for (let i = 0; i <= 4; i++) {
    const gridY = bottomY - (availableHeight * i) / 4;
    ctx.beginPath();
    ctx.moveTo(leftMargin, gridY);
    ctx.lineTo(width - rightMargin, gridY);
    ctx.stroke();

    // Axis label
    ctx.fillStyle = '#80868b';
    ctx.font = '12px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'right';
    const axisVal = minVal + (range * i) / 4;
    ctx.fillText(formatValue(axisVal), leftMargin - 10, gridY + 4);
  }

  // Points coordinates
  const points = data.map((d, i) => {
    const x = leftMargin + i * stepX;
    const y = bottomY - ((d.value - minVal) / range) * availableHeight;
    return { x, y, value: d.value, label: d.label };
  });

  // Draw Area under line (subtle gradient)
  ctx.beginPath();
  ctx.moveTo(points[0].x, bottomY);
  points.forEach(p => ctx.lineTo(p.x, p.y));
  ctx.lineTo(points[points.length - 1].x, bottomY);
  ctx.closePath();
  ctx.fillStyle = 'rgba(26, 115, 232, 0.08)';
  ctx.fill();

  // Draw Line
  ctx.strokeStyle = '#1a73e8';
  ctx.lineWidth = 3;
  ctx.lineJoin = 'round';
  ctx.beginPath();
  points.forEach((p, i) => {
    if (i === 0) ctx.moveTo(p.x, p.y);
    else ctx.lineTo(p.x, p.y);
  });
  ctx.stroke();

  // Draw Dots & Labels
  points.forEach(p => {
    // Outer white dot
    ctx.beginPath();
    ctx.arc(p.x, p.y, 6, 0, 2 * Math.PI);
    ctx.fillStyle = '#ffffff';
    ctx.fill();
    ctx.strokeStyle = '#1a73e8';
    ctx.lineWidth = 2.5;
    ctx.stroke();

    // Value
    ctx.fillStyle = '#202124';
    ctx.font = 'bold 11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(formatValue(p.value), p.x, p.y - 12);

    // X-Axis Label
    ctx.fillStyle = '#5f6368';
    ctx.font = '600 11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.fillText(p.label, p.x, bottomY + 20);
  });
}

function drawRoundedRect(ctx, x, y, width, height, radius) {
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + width - radius, y);
  ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
  ctx.lineTo(x + width, y + height - radius);
  ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
  ctx.lineTo(x + radius, y + height);
  ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

function escapeHtmlAttr(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Convenience method that parses raw text/JSON and renders an HTML image wrapper.
 * @param {string} text - Raw chart JSON string
 * @returns {string|null} HTML markup with <img src="data:image/png;base64,...">
 */
export function renderChartHtml(text) {
  const spec = parseChartSpec(text);
  if (!spec) return null;

  const dataUri = renderChartToDataUrl(spec);
  if (dataUri) {
    const chartTitle = spec.title || 'Generated Chart';
    const alt = spec.title ? `${spec.title} (${spec.chartType} chart)` : 'Generated Chart';
    return `<div class="rendered-chart-container" data-chart-title="${escapeHtmlAttr(chartTitle)}" style="margin:18px 0; text-align:center;">
      <img src="${dataUri}" alt="${alt}" style="max-width:100%; border-radius:8px; border:1px solid #d2e3fc; box-shadow:0 3px 12px rgba(0,0,0,0.07);" />
    </div>`;
  }

  // Graceful fallback if Canvas is unavailable: render responsive visual HTML table
  const total = spec.data.reduce((sum, d) => sum + (d.value > 0 ? d.value : 0), 0) || 1;
  const rows = spec.data.map((d, i) => {
    const pct = ((d.value / total) * 100).toFixed(1);
    const color = PALETTE[i % PALETTE.length];
    return `<tr>
      <td style="padding:8px 12px; border-bottom:1px solid #e0e0e0;"><span style="display:inline-block; width:10px; height:10px; border-radius:50%; background:${color}; margin-right:8px;"></span>${d.label}</td>
      <td style="padding:8px 12px; border-bottom:1px solid #e0e0e0; font-weight:bold; text-align:right;">${formatValue(d.value)}</td>
      <td style="padding:8px 12px; border-bottom:1px solid #e0e0e0; color:#5f6368; text-align:right;">${pct}%</td>
    </tr>`;
  }).join('');

  return `<div class="chart-table-fallback" style="margin:16px 0; border:1px solid #d2e3fc; border-radius:8px; overflow:hidden;">
    <div style="background:#f8fafd; padding:10px 14px; font-weight:bold; color:#1a73e8; border-bottom:1px solid #d2e3fc;">📊 ${spec.title}</div>
    <table style="width:100%; border-collapse:collapse; font-size:13px;">
      <thead>
        <tr style="background:#fafafa; color:#5f6368; text-align:left;">
          <th style="padding:8px 12px;">Category</th>
          <th style="padding:8px 12px; text-align:right;">Metric</th>
          <th style="padding:8px 12px; text-align:right;">Share</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

