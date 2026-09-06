/**
 * Gemini Enterprise Backend Proxy for Microsoft 365 (Word, PowerPoint, Excel)
 * 
 * Provides grounded enterprise generative AI inference using Gemini 2.5 Flash
 * and multimodal image chart generation via Gemini 2.5 Flash Image on Vertex AI.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import functions from '@google-cloud/functions-framework';
import { GoogleAuth } from 'google-auth-library';
import express from 'express';
import corsLib from 'cors';

const cors = corsLib({ origin: true });

// Environment Configuration (Configured via .env or GCP Cloud Run environment variables)
const PROJECT_ID = process.env.GE_GCP_PROJECT_ID || process.env.GCP_PROJECT_ID || process.env.GOOGLE_CLOUD_PROJECT;
const REGION = process.env.GCP_REGION || 'us-central1';
const DATASTORE_ID = process.env.VERTEX_DATASTORE_ID || process.env.VERTEX_DATASTORE || '';

// Gemini Enterprise StreamAssist API Configuration (STRICT: All requests route through StreamAssist)
const BACKEND_MODE = (process.env.BACKEND_MODE || 'streamassist').toLowerCase();
const STREAM_ASSIST_ENDPOINT_LOCATION = process.env.STREAM_ASSIST_ENDPOINT_LOCATION || 'global';
const GCP_LOCATION = process.env.GE_GCP_LOCATION || process.env.GCP_LOCATION || 'global';
const ENTERPRISE_APP_ID = process.env.GEMINI_ENTERPRISE_APP_ID || process.env.VERTEX_DATASTORE_ID || '';
const ENTERPRISE_COLLECTION_ID = process.env.GEMINI_ENTERPRISE_COLLECTION_ID || 'default_collection';
const ENTERPRISE_ASSISTANT_ID = process.env.GEMINI_ENTERPRISE_ASSISTANT_ID || 'default_assistant';
const ALLOW_SERVICE_ACCOUNT_FALLBACK = (process.env.ALLOW_SERVICE_ACCOUNT_FALLBACK || 'false').toLowerCase() === 'true';

const auth = new GoogleAuth({
  scopes: 'https://www.googleapis.com/auth/cloud-platform'
});

if (!PROJECT_ID) {
  console.warn('WARNING: GE_GCP_PROJECT_ID environment variable is not set.');
}

// In-memory session store for multi-turn chats
const sessionStore = new Map();

const SYSTEM_INSTRUCTION_TEXT = `You are Ask Gemini, a versatile, intelligent corporate AI assistant for Microsoft Office (Word, PowerPoint, and Excel).

CRITICAL FORMATTING RULES:
1. NEVER output email or memo headers (e.g., "Date:", "To:", "From:", "Subject:", "Memo:", or salutations). Never format responses as an email unless explicitly asked. Start directly with the main document title (# Title).
2. DYNAMIC VISUAL CHARTS & DIAGRAMS (MANDATORY WHEN REQUESTED):
   Whenever the user asks for charts, graphs, data visualizations, diagrams, or comparisons, you MUST embed an image using:
   ![Chart Title](image: A modern corporate financial bar chart comparing Google Services ($94.5B) vs Google Cloud ($24.7B) revenue and operating income for Q1 and Q2 2026. Clean white background, modern flat 2D graphic design, crisp typography.)
   OR
   \`\`\`image
   A modern corporate financial bar chart comparing Google Services vs Google Cloud Q1 and Q2 2026 revenue and operating income ($M). Crisp flat 2D vector style, white background.
   \`\`\`
3. EXECUTIVE BULLET POINTS & TYPOGRAPHY:
   - For executive briefings, presentations, and reports, format bullet points with bold lead-ins and highlighted metrics (e.g. * **Google Services ($94.54B):** Reported strong growth driven by Search and YouTube with **$39.54B** in operating income).
   - Use clean, punchy takeaways that C-level executives can scan and absorb immediately.
4. Use standard markdown tables (| Header 1 | Header 2 | ...) for tabular data and executive callouts (> [!NOTE] ...).`;

// Inlines images into base64 data URIs for native Office insertion (strictly without direct model calls)
async function inlineImagesInContent(text) {
  if (!text) return '';
  let processed = text;

  // 1. Fetch remote image URLs (https://...) and convert to base64 data URIs for native Office insertion
  const remoteImgRegex = /!\[([^\]]*)\]\((https:\/\/[^\s\)]+)\)/gi;
  const remoteMatches = [...processed.matchAll(remoteImgRegex)];
  for (const m of remoteMatches) {
    const [fullMatch, alt, url] = m;
    try {
      const imgRes = await fetch(url);
      if (imgRes.ok) {
        const buffer = await imgRes.arrayBuffer();
        const base64 = Buffer.from(buffer).toString('base64');
        const mime = imgRes.headers.get('content-type') || 'image/png';
        const dataUri = `data:${mime};base64,${base64}`;
        const imgTag = `<div style="margin:18px 0; text-align:center;"><img src="${dataUri}" alt="${escapeXml(alt || 'Generated Image')}" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 3px 10px rgba(0,0,0,0.08);" /></div>`;
        processed = processed.replace(fullMatch, imgTag);
      }
    } catch (fetchErr) {
      console.warn('[INLINE_IMAGES] Could not fetch remote image URL:', url, fetchErr.message);
    }
  }

  // 2. Convert markdown data:image URIs to styled HTML images
  const dataImgRegex = /!\[([^\]]*)\]\((data:image\/[^;]+;base64,[^\)]+)\)/gi;
  const dataMatches = [...processed.matchAll(dataImgRegex)];
  for (const m of dataMatches) {
    const [fullMatch, alt, dataUri] = m;
    const imgTag = `<div style="margin:18px 0; text-align:center;"><img src="${dataUri}" alt="${escapeXml(alt || 'Visual')}" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 3px 10px rgba(0,0,0,0.08);" /></div>`;
    processed = processed.replace(fullMatch, imgTag);
  }

  return processed;
}


// 1. Office-Native Visual Storyboard (100% Word & PPT Compatible)
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

// 2. Office-Native Visual Bar Chart (100% Word & PPT Compatible)
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

// 3. Office-Native Visual Process Flow (100% Word & PPT Compatible)
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

// 4. Office-Native KPI Stat Grid (100% Word & PPT Compatible)
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

function processChartsInContent(text) {
  if (!text) return '';
  return text.replace(/```(?:json)?\s*chart\s*([\s\S]*?)```/gi, (match, jsonString) => {
    try {
      const cleanJson = jsonString.trim();
      const chartSpec = JSON.parse(cleanJson);
      return renderChartToImageTag(chartSpec);
    } catch (e) {
      console.warn('Could not parse chart JSON block:', e);
      return '';
    }
  });
}

async function handleGeminiRequest(req, res) {
  // STRICT: Route all requests through Gemini Enterprise StreamAssist API
  return handleGeminiEnterpriseRequest(req, res);
}

function extractImagesFromObject(obj, targetList) {
  if (!obj || typeof obj !== 'object') return;

  // 1. Blob object with base64 data
  if (obj.blob && (obj.blob.data || obj.blob.bytesBase64Encoded)) {
    const mime = obj.blob.mimeType || '';
    if (mime.startsWith('image/')) {
      const b64 = obj.blob.data || obj.blob.bytesBase64Encoded;
      const uri = `data:${mime};base64,${b64}`;
      if (!targetList.includes(uri)) targetList.push(uri);
    }
  }

  // 2. inlineData with base64 data
  if (obj.inlineData && (obj.inlineData.data || obj.inlineData.bytesBase64Encoded)) {
    const mime = obj.inlineData.mimeType || '';
    if (mime.startsWith('image/')) {
      const b64 = obj.inlineData.data || obj.inlineData.bytesBase64Encoded;
      const uri = `data:${mime};base64,${b64}`;
      if (!targetList.includes(uri)) targetList.push(uri);
    }
  }

  // 3. Object with image mimeType and data
  if (obj.mimeType && typeof obj.mimeType === 'string' && obj.mimeType.startsWith('image/')) {
    const b64 = obj.data || obj.bytesBase64Encoded;
    if (b64 && typeof b64 === 'string') {
      const uri = `data:${obj.mimeType};base64,${b64}`;
      if (!targetList.includes(uri)) targetList.push(uri);
    }
  }

  // 4. Object with direct fileUri or imageUri
  if (obj.fileData && obj.fileData.fileUri) {
    if (!targetList.includes(obj.fileData.fileUri)) targetList.push(obj.fileData.fileUri);
  }
  if (obj.imageUri && typeof obj.imageUri === 'string') {
    if (!targetList.includes(obj.imageUri)) targetList.push(obj.imageUri);
  }

  // Recurse into arrays and child objects
  if (Array.isArray(obj)) {
    for (const item of obj) extractImagesFromObject(item, targetList);
  } else {
    for (const key of Object.keys(obj)) {
      if (typeof obj[key] === 'object' && obj[key] !== null) {
        extractImagesFromObject(obj[key], targetList);
      }
    }
  }
}

function extractReplies(chunk) {
  if (chunk.answer?.reply) {
    return [chunk.answer.reply];
  }
  if (Array.isArray(chunk.answer?.replies)) {
    return chunk.answer.replies;
  }
  if (Array.isArray(chunk.replies)) {
    return chunk.replies;
  }
  if (chunk.reply && typeof chunk.reply === 'object') {
    return [chunk.reply];
  }
  return [];
}

function findPendingFiles(obj, targetList, currentSession) {
  if (!obj || typeof obj !== 'object') return;
  if (obj.file && (obj.file.fileId || obj.file.id)) {
    const fId = obj.file.fileId || obj.file.id;
    const mime = obj.file.mimeType || 'image/png';
    if (!targetList.some(item => item.fileId === fId)) {
      targetList.push({ fileId: fId, mimeType: mime, session: currentSession });
    }
  }
  if (Array.isArray(obj)) {
    for (const item of obj) findPendingFiles(item, targetList, currentSession);
  } else {
    for (const key of Object.keys(obj)) {
      if (typeof obj[key] === 'object' && obj[key] !== null) {
        findPendingFiles(obj[key], targetList, currentSession);
      }
    }
  }
}

async function processStreamAssistChunks(parsedChunks, originalSessionId, sessionResetOccurred = false, headers = null) {
  let aggregatedText = '';
  let returnedSessionResource = null;
  const citations = [];
  const seenCitations = new Set();
  const extractedImages = [];
  const thoughtParts = [];
  const pendingFiles = [];

  for (const chunk of parsedChunks) {
    if (chunk.sessionInfo?.session) {
      returnedSessionResource = chunk.sessionInfo.session;
    }

    // Check if image generation tool was invoked
    const invokedTools = chunk.invocationTools || chunk.answer?.invocationTools || [];
    if (Array.isArray(invokedTools) && invokedTools.includes('image_generation')) {
      console.log('[STREAM_ASSIST] Image generation tool invoked by Gemini Enterprise');
    }

    // Comprehensive recursive image and file extraction from the entire chunk
    extractImagesFromObject(chunk, extractedImages);
    findPendingFiles(chunk, pendingFiles, chunk.sessionInfo?.session || returnedSessionResource || originalSessionId);

    // Defensive reply normalization (supports both singular reply and plural replies)
    const replies = extractReplies(chunk);

    for (const reply of replies) {
      const contentObj = reply.groundedContent?.content || reply.content;

      if (contentObj) {
        // Direct text fragment
        if (contentObj.text && !contentObj.thought) {
          aggregatedText += contentObj.text;
        } else if (contentObj.thought && contentObj.text) {
          thoughtParts.push(contentObj.text);
        }

        // Direct blob (Base64 image/media bytes)
        if (contentObj.blob && (contentObj.blob.data || contentObj.blob.bytesBase64Encoded)) {
          const mime = contentObj.blob.mimeType || 'image/png';
          if (mime.startsWith('image/')) {
            const b64 = contentObj.blob.data || contentObj.blob.bytesBase64Encoded;
            const uri = `data:${mime};base64,${b64}`;
            if (!extractedImages.includes(uri)) {
              extractedImages.push(uri);
              console.log(`[STREAM_ASSIST] Found content.blob image artifact (${b64.length} bytes)`);
            }
          }
        }

        // Direct file reference (fileId)
        if (contentObj.file && (contentObj.file.fileId || contentObj.file.id)) {
          const fId = contentObj.file.fileId || contentObj.file.id;
          const mime = contentObj.file.mimeType || 'image/png';
          if (!pendingFiles.some(f => f.fileId === fId)) {
            pendingFiles.push({ fileId: fId, mimeType: mime, session: chunk.sessionInfo?.session || returnedSessionResource || originalSessionId });
          }
        }

        // Content parts (if structured as parts array)
        if (Array.isArray(contentObj.parts) && contentObj.parts.length > 0) {
          for (const part of contentObj.parts) {
            if (part.text && !part.thought) {
              aggregatedText += part.text;
            } else if (part.thought && part.text) {
              thoughtParts.push(part.text);
            } else if (part.inlineData?.data) {
              const mime = part.inlineData.mimeType || 'image/png';
              if (mime.startsWith('image/')) {
                const uri = `data:${mime};base64,${part.inlineData.data}`;
                if (!extractedImages.includes(uri)) extractedImages.push(uri);
                console.log(`[STREAM_ASSIST] Found inlineData image artifact (${part.inlineData.data.length} bytes)`);
              }
            } else if (part.blob && (part.blob.data || part.blob.bytesBase64Encoded)) {
              const mime = part.blob.mimeType || 'image/png';
              if (mime.startsWith('image/')) {
                const b64 = part.blob.data || part.blob.bytesBase64Encoded;
                const uri = `data:${mime};base64,${b64}`;
                if (!extractedImages.includes(uri)) extractedImages.push(uri);
                console.log(`[STREAM_ASSIST] Found part.blob image artifact (${b64.length} bytes)`);
              }
            } else if (part.file && (part.file.fileId || part.file.id)) {
              const fId = part.file.fileId || part.file.id;
              const mime = part.file.mimeType || 'image/png';
              if (!pendingFiles.some(f => f.fileId === fId)) {
                pendingFiles.push({ fileId: fId, mimeType: mime, session: chunk.sessionInfo?.session || returnedSessionResource || originalSessionId });
              }
            } else if (part.fileData?.fileUri) {
              if (!extractedImages.includes(part.fileData.fileUri)) extractedImages.push(part.fileData.fileUri);
            } else if (part.image?.uri || part.image?.url) {
              const uri = part.image.uri || part.image.url;
              if (!extractedImages.includes(uri)) extractedImages.push(uri);
            }
          }
        }
      } else {
        // Fallback to top-level reply text properties
        if (reply.text) {
          aggregatedText += reply.text;
        } else if (reply.replyText) {
          aggregatedText += reply.replyText;
        }
      }

      // 2. Check reply media/images
      const replyMedia = reply.media || reply.images || reply.groundedContent?.media || reply.groundedContent?.images;
      if (replyMedia) {
        const mediaList = Array.isArray(replyMedia) ? replyMedia : [replyMedia];
        for (const m of mediaList) {
          if (typeof m === 'string') {
            if (!extractedImages.includes(m)) extractedImages.push(m);
          } else if (m.bytesBase64Encoded || m.data) {
            const uri = `data:${m.mimeType || 'image/png'};base64,${m.bytesBase64Encoded || m.data}`;
            if (!extractedImages.includes(uri)) extractedImages.push(uri);
          } else if (m.uri || m.url) {
            const uri = m.uri || m.url;
            if (!extractedImages.includes(uri)) extractedImages.push(uri);
          }
        }
      }

      // 3. Grounded citations
      const grounded = reply.groundedContent || {};
      if (grounded.searchChunk || grounded.web || reply.replyId) {
        const title = grounded.title || grounded.searchChunk?.title || 'Enterprise Data Source';
        const uri = grounded.uri || grounded.searchChunk?.uri || '';
        if (uri && !seenCitations.has(title)) {
          seenCitations.add(title);
          citations.push({ title, uri });
        }
      }
    }

    if (!aggregatedText) {
      const chunkText = chunk.answer?.reply?.groundedContent?.content?.text
        || chunk.answer?.reply?.content?.text
        || chunk.answer?.reply?.text
        || chunk.answer?.replyText 
        || chunk.answer?.text 
        || chunk.answer?.content?.text 
        || chunk.replyText 
        || chunk.text 
        || '';
      if (chunkText) {
        aggregatedText += chunkText;
      }
    }
  }

  // 4. Download any pending fileId resources emitted by Gemini Enterprise
  if (pendingFiles.length > 0 && headers) {
    const downloadHeaders = {
      'Authorization': headers['Authorization'] || headers['authorization'] || '',
      'Accept': '*/*'
    };
    const userProj = headers['X-Goog-User-Project'] || headers['x-goog-user-project'];
    if (userProj) {
      downloadHeaders['X-Goog-User-Project'] = userProj;
    }

    for (const file of pendingFiles) {
      const sessionPath = file.session || returnedSessionResource || originalSessionId;
      if (!sessionPath) continue;
      const regionalHost = `${STREAM_ASSIST_ENDPOINT_LOCATION}-discoveryengine.googleapis.com`;
      const encodedId = encodeURIComponent(file.fileId);

      // Multiple candidate endpoints in order of priority:
      // 1. :downloadFile with alt=media (standard Google Cloud API raw media streaming)
      // 2. :downloadFile without alt=media (standard custom method)
      // 3. snake_case file_id variants
      // 4. direct session /files/{fileId} resource path
      const candidates = [
        `https://${regionalHost}/v1alpha/${sessionPath}:downloadFile?fileId=${encodedId}&alt=media`,
        `https://${regionalHost}/v1alpha/${sessionPath}:downloadFile?fileId=${encodedId}`,
        `https://${regionalHost}/v1alpha/${sessionPath}:downloadFile?file_id=${encodedId}&alt=media`,
        `https://${regionalHost}/v1alpha/${sessionPath}:downloadFile?file_id=${encodedId}`,
        `https://${regionalHost}/v1alpha/${sessionPath}/files/${encodedId}?alt=media`,
        `https://${regionalHost}/v1alpha/${sessionPath}/files/${encodedId}`
      ];

      let downloaded = false;
      for (const downloadUrl of candidates) {
        if (downloaded) break;
        try {
          console.log(`[STREAM_ASSIST] Attempting to fetch file '${file.fileId}' from ${downloadUrl}...`);
          const fileRes = await fetch(downloadUrl, {
            method: 'GET',
            headers: downloadHeaders
          });

          const contentType = fileRes.headers.get('content-type') || '';
          console.log(`[STREAM_ASSIST] Response status: ${fileRes.status} ${fileRes.statusText}, Content-Type: ${contentType}`);

          if (!fileRes.ok) {
            const errText = await fileRes.text();
            console.warn(`[STREAM_ASSIST] Candidate failed (HTTP ${fileRes.status}):`, errText.slice(0, 200));
            continue;
          }

          const arrayBuffer = await fileRes.arrayBuffer();
          const buf = Buffer.from(arrayBuffer);
          console.log(`[STREAM_ASSIST] Downloaded payload size: ${buf.length} bytes`);

          if (buf.length === 0) {
            console.log(`[STREAM_ASSIST] Empty body returned (0 bytes), trying next candidate...`);
            continue;
          }

          // Check for image magic bytes
          let detectedMime = null;
          if (buf.length >= 8 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47) {
            detectedMime = 'image/png';
          } else if (buf.length >= 3 && buf[0] === 0xFF && buf[1] === 0xD8 && buf[2] === 0xFF) {
            detectedMime = 'image/jpeg';
          } else if (buf.length >= 6 && buf.toString('ascii', 0, 4) === 'GIF8') {
            detectedMime = 'image/gif';
          } else if (buf.length >= 12 && buf.toString('ascii', 0, 4) === 'RIFF' && buf.toString('ascii', 8, 12) === 'WEBP') {
            detectedMime = 'image/webp';
          }

          if (detectedMime || contentType.startsWith('image/')) {
            const mime = detectedMime || contentType.split(';')[0].trim() || file.mimeType || 'image/png';
            const b64 = buf.toString('base64');
            const uri = `data:${mime};base64,${b64}`;
            if (!extractedImages.includes(uri)) {
              extractedImages.push(uri);
              console.log(`[STREAM_ASSIST] Successfully extracted binary image '${file.fileId}' (${mime}, ${b64.length} base64 chars)`);
            }
            downloaded = true;
            break;
          }

          // Check if payload is JSON
          const rawText = buf.toString('utf-8').trim();
          if (rawText.startsWith('{') || rawText.startsWith('[')) {
            try {
              const fileJson = JSON.parse(rawText);
              console.log(`[STREAM_ASSIST] JSON keys returned:`, Object.keys(fileJson));
              const mime = fileJson.mimeType || file.mimeType || 'image/png';
              const b64 = fileJson.data || fileJson.bytesBase64Encoded || fileJson.fileContents || fileJson.content;
              if (b64) {
                const uri = b64.startsWith('data:') ? b64 : `data:${mime};base64,${b64}`;
                if (!extractedImages.includes(uri)) {
                  extractedImages.push(uri);
                  console.log(`[STREAM_ASSIST] Successfully extracted JSON base64 image '${file.fileId}' (${b64.length} chars)`);
                }
                downloaded = true;
                break;
              }

              // Check if a signed download URL or URI was returned
              const downloadUri = fileJson.downloadUri || fileJson.downloadUrl || fileJson.signedUrl || fileJson.uri || fileJson.url;
              if (downloadUri && typeof downloadUri === 'string') {
                console.log(`[STREAM_ASSIST] Found download URL in JSON: ${downloadUri.slice(0, 100)}... Fetching...`);
                const signedRes = await fetch(downloadUri);
                if (signedRes.ok) {
                  const sBuf = Buffer.from(await signedRes.arrayBuffer());
                  const sMime = signedRes.headers.get('content-type') || mime;
                  const uri = `data:${sMime};base64,${sBuf.toString('base64')}`;
                  if (!extractedImages.includes(uri)) {
                    extractedImages.push(uri);
                    console.log(`[STREAM_ASSIST] Successfully fetched image from signed URL (${sBuf.length} bytes)`);
                  }
                  downloaded = true;
                  break;
                }
              }
            } catch (jsonErr) {
              console.warn(`[STREAM_ASSIST] Could not parse text as JSON:`, jsonErr.message, rawText.slice(0, 200));
            }
          }
        } catch (candErr) {
          console.warn(`[STREAM_ASSIST] Exception querying candidate '${downloadUrl}':`, candErr.message);
        }
      }

      if (!downloaded) {
        console.warn(`[STREAM_ASSIST] Unable to retrieve file content for '${file.fileId}' across all candidates`);
      }
    }
  }

  // 5. Append extracted images as Markdown images if not already embedded
  if (extractedImages.length > 0) {
    if (!aggregatedText.trim()) {
      aggregatedText = "Here is the image generated based on your request:\n\n";
    }
    for (const imgUrl of extractedImages) {
      if (!aggregatedText.includes(imgUrl)) {
        aggregatedText += `\n\n![Generated Image](${imgUrl})\n\n`;
      }
    }
  }

  // Graceful fallback if no direct text was found
  if (!aggregatedText.trim()) {
    if (thoughtParts.length > 0) {
      aggregatedText = thoughtParts.join('\n\n');
    } else {
      console.warn('[STREAM_ASSIST] StreamAssist returned empty text and no images. Parsed chunk count:', parsedChunks.length, 'Chunk sample:', JSON.stringify(parsedChunks).slice(0, 1000));
      aggregatedText = "I processed your request with Gemini Enterprise, but no narrative content was generated. Please try rephrasing your request or asking for slide summaries, structured tables, or data comparisons.";
    }
  }

  let shortSessionId = returnedSessionResource;
  if (returnedSessionResource && returnedSessionResource.includes('/sessions/')) {
    shortSessionId = returnedSessionResource.split('/sessions/').pop();
  }

  return {
    resultText: aggregatedText,
    sessionResource: returnedSessionResource,
    sessionId: shortSessionId || (sessionResetOccurred ? null : originalSessionId),
    citations: citations,
    images: extractedImages
  };
}


async function callStreamAssistAPI({ prompt, sessionId, userId, userPseudoId, userGoogleToken, authMode, attachments }) {
  if (!PROJECT_ID) {
    throw new Error('GE_GCP_PROJECT_ID environment variable is required for StreamAssist');
  }
  if (!ENTERPRISE_APP_ID) {
    throw new Error('GEMINI_ENTERPRISE_APP_ID environment variable is required for StreamAssist');
  }

  const activeUserId = userPseudoId || userId || 'office_365_user';
  let bearerToken = null;

  if (userGoogleToken) {
    console.log(JSON.stringify({
      severity: 'INFO',
      message: `[AUTH] Using End-User Google Token for Discovery Engine StreamAssist (Mode: ${authMode || 'user_token'}, User: ${activeUserId})`,
      user_id: activeUserId,
      auth_mode: authMode || 'user_token',
      token_source: 'X-End-User-Google-Token'
    }));
    bearerToken = userGoogleToken;
  } else {
    if (ALLOW_SERVICE_ACCOUNT_FALLBACK) {
      console.warn(JSON.stringify({
        severity: 'WARNING',
        message: `[AUTH_FALLBACK] No end-user Google token provided for user '${activeUserId}'. ALLOW_SERVICE_ACCOUNT_FALLBACK is enabled. Falling back to Cloud Run Service Account ADC credentials.`,
        user_id: activeUserId,
        fallback_reason: 'MISSING_END_USER_GOOGLE_TOKEN',
        allow_service_account_fallback: true
      }));

      const client = await auth.getClient();
      const accessTokenObj = await client.getAccessToken();
      bearerToken = typeof accessTokenObj === 'string' ? accessTokenObj : accessTokenObj.token;
    } else {
      console.error(JSON.stringify({
        severity: 'ERROR',
        message: `[AUTH_REJECTED] Request for user '${activeUserId}' rejected: End-user Google token is required to enforce Gemini Enterprise licensing, and ALLOW_SERVICE_ACCOUNT_FALLBACK is false.`,
        user_id: activeUserId,
        allow_service_account_fallback: false
      }));

      const authErr = new Error(`End-user Google authentication token is required to access Gemini Enterprise. Service account fallback is disabled.`);
      authErr.statusCode = 403;
      throw authErr;
    }
  }

  const endpointUrl = `https://${STREAM_ASSIST_ENDPOINT_LOCATION}-discoveryengine.googleapis.com/v1alpha/projects/${PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/assistants/${ENTERPRISE_ASSISTANT_ID}:streamAssist`;
  // Upgraded to v1alpha for full toolsSpec/grounding support

  const requestBody = {
    query: {
      text: prompt
    }
  };

  // Enable Vertex AI Search and Image Generation (powered by Imagen) in toolsSpec
  requestBody.toolsSpec = {
    vertexAiSearchSpec: {},
    imageGenerationSpec: {}
  };


  if (sessionId) {
    let fullSessionName = sessionId;
    if (!sessionId.startsWith('projects/')) {
      fullSessionName = `projects/${PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions/${sessionId}`;
    }
    requestBody.session = fullSessionName;
  } else {
    // Explicitly create session tagged with the userPseudoId for user history, attribution, and context files
    try {
      const sessionCreateUrl = `https://${STREAM_ASSIST_ENDPOINT_LOCATION}-discoveryengine.googleapis.com/v1alpha/projects/${PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions`;
      const sessionRes = await fetch(sessionCreateUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${bearerToken}`,
          'Content-Type': 'application/json',
          'X-Goog-User-Project': PROJECT_ID
        },
        body: JSON.stringify({ userPseudoId: activeUserId })
      });
      if (sessionRes.ok) {
        const sessionData = await sessionRes.json();
        if (sessionData.name) {
          requestBody.session = sessionData.name;
          console.log(`[SESSION] Created new session tagged for user '${activeUserId}': ${sessionData.name}`);
        }
      }
    } catch (sessionErr) {
      console.warn(`[SESSION_WARN] Could not pre-create session with userPseudoId '${activeUserId}':`, sessionErr.message);
    }
  }

/**
 * Safely extracts raw text from document attachments (PDF, Word DOCX, Markdown, Text, CSV, JSON)
 */
async function extractDocumentText(att) {
  if (!att || !att.fileContents) return '';
  try {
    const buffer = Buffer.from(att.fileContents, 'base64');
    const mime = (att.mimeType || '').toLowerCase();
    const name = (att.fileName || '').toLowerCase();

    // Plain text / Markdown / JSON / CSV
    if (mime.includes('text') || mime.includes('markdown') || mime.includes('json') || mime.includes('csv') || name.endsWith('.txt') || name.endsWith('.md')) {
      return buffer.toString('utf-8');
    }

    // Word documents (.docx)
    if (mime.includes('word') || name.endsWith('.docx')) {
      const mammoth = await import('mammoth');
      const result = await (mammoth.default || mammoth).extractRawText({ buffer });
      return result.value || '';
    }

    // PDF documents (.pdf)
    if (mime.includes('pdf') || name.endsWith('.pdf')) {
      const pdfModule = await import('pdf-parse');
      const { PDFParse } = pdfModule.default || pdfModule;
      if (PDFParse) {
        const parser = new PDFParse({ data: buffer });
        await parser.load();
        const textResult = await parser.getText();
        return (typeof textResult === 'string' ? textResult : (textResult.text || '')).trim();
      }
    }
  } catch (err) {
    console.warn(`[EXTRACT_TEXT_WARN] Could not extract text from '${att.fileName}':`, err.message);
  }
  return '';
}

  // Handle document attachments:
  // 1. Extract text from documents and inject into query.text so StreamAssist ALWAYS has full grounded context
  // 2. Try Discovery Engine addContextFile across multiple API versions and locations
  const uploadedFileIds = [];
  const extractedDocs = [];

  if (attachments && Array.isArray(attachments) && attachments.length > 0) {
    for (const att of attachments) {
      // 1. Extract text first
      try {
        const text = await extractDocumentText(att);
        if (text && text.trim().length > 0) {
          extractedDocs.push({ fileName: att.fileName, text: text.trim() });
          console.log(`[STREAM_ASSIST] Successfully extracted text from '${att.fileName}' (${text.trim().length} chars)`);
        }
      } catch (extractErr) {
        console.warn(`[STREAM_ASSIST_WARN] Could not extract text from '${att.fileName}':`, extractErr.message);
      }

      // 2. Attempt addContextFile via Discovery Engine if session exists
      if (requestBody.session) {
        const regionalHost = STREAM_ASSIST_ENDPOINT_LOCATION && STREAM_ASSIST_ENDPOINT_LOCATION !== 'global'
          ? `${STREAM_ASSIST_ENDPOINT_LOCATION}-discoveryengine.googleapis.com`
          : 'discoveryengine.googleapis.com';

        const candidateUrls = [
          `https://${regionalHost}/v1alpha/${requestBody.session}:addContextFile`,
          `https://${regionalHost}/v1alpha/projects/${PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions/-:addContextFile`
        ];

        let fileAdded = false;
        for (const addFileUrl of candidateUrls) {
          if (fileAdded) break;
          try {
            const addRes = await fetch(addFileUrl, {
              method: 'POST',
              headers: {
                'Authorization': `Bearer ${bearerToken}`,
                'Content-Type': 'application/json',
                'X-Goog-User-Project': PROJECT_ID
              },
              body: JSON.stringify({
                fileName: att.fileName,
                mimeType: att.mimeType || 'application/pdf',
                fileContents: att.fileContents
              })
            });

            if (addRes.ok) {
              const addData = await addRes.json();
              if (addData.fileId) {
                uploadedFileIds.push(addData.fileId);
                if (addData.session) {
                  requestBody.session = addData.session;
                }
                fileAdded = true;
                console.log(`[STREAM_ASSIST] Successfully registered context file '${att.fileName}' -> fileId: ${addData.fileId}, tokenCount: ${addData.tokenCount || 'N/A'}`);
              }
            } else if (addRes.status === 404) {
              // Endpoint does not support addContextFile ("Method not found"). Break early and use reliable text grounding.
              if (process.env.VERBOSE_LOGGING === 'true') {
                console.log(`[STREAM_ASSIST] addContextFile not supported on this engine (404). Proceeding with text grounding.`);
              }
              break;
            } else {
              if (process.env.VERBOSE_LOGGING === 'true') {
                const addErrText = await addRes.text();
                console.log(`[STREAM_ASSIST] addContextFile status ${addRes.status}: ${addErrText.slice(0, 100)}`);
              }
            }
          } catch (attErr) {
            if (process.env.VERBOSE_LOGGING === 'true') {
              console.log(`[STREAM_ASSIST] addContextFile network note: ${attErr.message}`);
            }
          }
        }
      }
    }

    if (uploadedFileIds.length > 0) {
      requestBody.fileIds = uploadedFileIds;
      console.log(`[STREAM_ASSIST] Successfully bound ${uploadedFileIds.length} context fileId(s) to streamAssist request:`, uploadedFileIds);
    } else if (extractedDocs.length > 0) {
      // Fallback: inject extracted text only if addContextFile could not bind fileIds
      const docContextBlocks = extractedDocs.map(d => `--- BEGIN DOCUMENT: ${d.fileName} ---\n${d.text}\n--- END DOCUMENT: ${d.fileName} ---`).join('\n\n');
      requestBody.query.text = `${requestBody.query.text}\n\n[ATTACHED DOCUMENTS CONTEXT]\n${docContextBlocks}\n[END ATTACHED DOCUMENTS CONTEXT]`;
      console.log(`[STREAM_ASSIST] Fallback: Grounded query.text with ${extractedDocs.length} attached document(s) (${docContextBlocks.length} chars total)`);
    }
  }

  const headers = {
    'Authorization': `Bearer ${bearerToken}`,
    'Content-Type': 'application/json',
    'X-Goog-User-Project': PROJECT_ID
  };

  console.log(`Calling StreamAssist API (${endpointUrl})... Session: ${requestBody.session || 'NEW'}`);

  let sessionResetOccurred = false;

  let apiRes = await fetch(endpointUrl, {
    method: 'POST',
    headers: headers,
    body: JSON.stringify(requestBody)
  });

  if (!apiRes.ok) {
    let errText = await apiRes.text();
    
    // 1. Auto-recover if session overflowed token limit (PROMPT_TOO_LARGE) or session ownership conflict occurs
    if (requestBody.session && (
      (apiRes.status === 400 && (errText.includes('PROMPT_TOO_LARGE') || errText.includes('INVALID_ARGUMENT'))) ||
      (apiRes.status === 403 && errText.includes('Session is not owned'))
    )) {
      sessionResetOccurred = true;
      console.warn(JSON.stringify({
        severity: 'WARNING',
        message: `[SESSION_RECOVERY] Discovery Engine session '${requestBody.session}' failed (${errText.includes('PROMPT_TOO_LARGE') ? 'PROMPT_TOO_LARGE token overflow' : 'session error'}). Retrying automatically with a fresh clean session...`,
        user_id: activeUserId
      }));
      delete requestBody.session;
      apiRes = await fetch(endpointUrl, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(requestBody)
      });
      if (!apiRes.ok) {
        errText = await apiRes.text();
      }
    }

    // 2. Auto-recover if toolsSpec (e.g. vertexAiSearchSpec / imageGenerationSpec) causes downstream timeout/error
    if (!apiRes.ok && (apiRes.status === 499 || (apiRes.status >= 400 && requestBody.toolsSpec))) {
      console.warn(JSON.stringify({
        severity: 'WARNING',
        message: `[TOOLS_SPEC_FALLBACK] StreamAssist failed with HTTP ${apiRes.status} (details: ${errText.substring(0, 100)}). Retrying with default assistant tools configuration...`,
        user_id: activeUserId
      }));
      delete requestBody.toolsSpec;
      apiRes = await fetch(endpointUrl, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(requestBody)
      });
      if (!apiRes.ok) {
        errText = await apiRes.text();
      }
    }

    if (!apiRes.ok) {
      console.error(JSON.stringify({
        severity: 'ERROR',
        message: `StreamAssist API call failed with HTTP ${apiRes.status}`,
        status_code: apiRes.status,
        user_id: activeUserId,
        error_detail: errText
      }));

      if (apiRes.status === 403) {
        const forbiddenErr = new Error(`Google Cloud Discovery Engine rejected the request (HTTP 403): User '${activeUserId}' does not have an active Gemini Enterprise license or IAM permission on engine '${ENTERPRISE_APP_ID}'. Details: ${errText}`);
        forbiddenErr.statusCode = 403;
        throw forbiddenErr;
      }

      const streamErr = new Error(`StreamAssist API returned HTTP ${apiRes.status}: ${errText}`);
      streamErr.statusCode = apiRes.status;
      throw streamErr;
    }
  }

  const rawResponseBody = await apiRes.text();
  let parsedChunks = [];

  try {
    const data = JSON.parse(rawResponseBody);
    parsedChunks = Array.isArray(data) ? data : [data];
  } catch (e) {
    const lines = rawResponseBody.split('\n').filter(l => l.trim().length > 0);
    for (const line of lines) {
      try {
        parsedChunks.push(JSON.parse(line));
      } catch (err) {
        console.warn('Could not parse streaming line chunk:', line.substring(0, 80));
      }
    }
  }

  return await processStreamAssistChunks(parsedChunks, sessionId, sessionResetOccurred, headers);
}

async function handleGeminiEnterpriseRequest(req, res) {
  cors(req, res, async () => {
    try {
      if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
      }

      const { prompt, sessionId, attachments } = req.body;
      const userId = req.body.userId || req.headers['x-end-user-id'] || req.headers['x-end-user-email'];
      const userPseudoId = req.body.userPseudoId || userId || req.headers['x-end-user-id'] || req.headers['x-end-user-email'] || 'office_365_user';
      const endUserName = req.body.authenticatedUser?.name || req.headers['x-end-user-name'] || '';
      const userGoogleToken = req.headers['x-end-user-google-token'] || req.headers['x-end-user-token'] || '';
      const authMode = req.headers['x-user-auth-mode'] || '';

      if (!prompt) {
        return res.status(400).json({ error: 'Prompt is required' });
      }

      console.log(`Processing Gemini Enterprise request (Mode: ${BACKEND_MODE})... Authenticated User: ${userPseudoId} (${endUserName || 'Corporate User'}, AuthMode: ${authMode || 'default'}, Attachments: ${attachments ? attachments.length : 0})`);

      if (BACKEND_MODE === 'streamassist' && ENTERPRISE_APP_ID) {
        try {
          const streamAssistResult = await callStreamAssistAPI({ 
            prompt, 
            sessionId, 
            userId, 
            userPseudoId, 
            userGoogleToken, 
            authMode,
            attachments
          });
          let rawText = streamAssistResult.resultText || '';

          if (streamAssistResult.citations && streamAssistResult.citations.length > 0) {
            const citationItems = streamAssistResult.citations.map(c => `* 📄 **${c.title}**${c.uri ? ` \`(${c.uri})\`` : ''}`);
            rawText += `\n\n---\n> [!NOTE] **Verified Gemini Enterprise Grounded Sources:**\n> ` + citationItems.join('\n> ');
          }

          const processedText = await inlineImagesInContent(rawText);

          return res.status(200).json({
            result: processedText,
            sessionId: streamAssistResult.sessionId,
            sessionResource: streamAssistResult.sessionResource,
            citations: streamAssistResult.citations,
            backendMode: 'streamassist'
          });
        } catch (streamAssistErr) {
          console.error('StreamAssist execution failed:', streamAssistErr.message);
          return res.status(streamAssistErr.statusCode || 500).json({
            error: 'Gemini Enterprise StreamAssist failed',
            details: streamAssistErr.message,
            statusCode: streamAssistErr.statusCode || 500
          });
        }
      }

      return res.status(400).json({
        error: 'Invalid Configuration',
        details: 'BACKEND_MODE must be streamassist and GEMINI_ENTERPRISE_APP_ID must be configured.'
      });
    } catch (error) {
      console.error('Error in Gemini Enterprise handler:', error);
      return res.status(500).json({
        error: 'Failed to process Gemini Enterprise request',
        details: error.message,
      });
    }
  });
}

functions.http('askGemini', handleGeminiRequest);
functions.http('askGeminiEnterprise', handleGeminiEnterpriseRequest);
functions.http('geminiProxy', handleGeminiRequest);

// Standalone Express Server (for Cloud Run & Docker execution)
const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(cors);
app.post('/askGeminiEnterprise', handleGeminiEnterpriseRequest);
app.post('/askGemini', handleGeminiRequest);
app.post('/', handleGeminiEnterpriseRequest);
app.get('/', (req, res) => res.status(200).send('Gemini Enterprise Proxy Service Healthy'));

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Gemini Enterprise Proxy HTTP Server running on port ${PORT}`);
});


