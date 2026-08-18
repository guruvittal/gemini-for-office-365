/**
 * Gemini Enterprise Backend Proxy for Microsoft 365 (Word, PowerPoint, Excel)
 * 
 * Provides grounded enterprise generative AI inference using Gemini 2.5 Flash
 * and multimodal image chart generation via Gemini 2.5 Flash Image on Vertex AI.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import functions from '@google-cloud/functions-framework';
import { VertexAI } from '@google-cloud/vertexai';
import { GoogleAuth } from 'google-auth-library';
import express from 'express';
import corsLib from 'cors';

const cors = corsLib({ origin: true });

// Environment Configuration (Configured via .env or GCP Cloud Run environment variables)
const PROJECT_ID = process.env.GCP_PROJECT_ID || process.env.GOOGLE_CLOUD_PROJECT;
const REGION = process.env.GCP_REGION || 'us-central1';
const MODEL_NAME = process.env.GEMINI_MODEL || 'gemini-2.5-flash';
const IMAGE_MODEL_NAME = process.env.GEMINI_IMAGE_MODEL || 'gemini-2.5-flash-image';
const DATASTORE_ID = process.env.VERTEX_DATASTORE_ID || process.env.VERTEX_DATASTORE || '';

// Gemini Enterprise StreamAssist API Configuration
const BACKEND_MODE = (process.env.BACKEND_MODE || 'streamassist').toLowerCase();
const STREAM_ASSIST_ENDPOINT_LOCATION = process.env.STREAM_ASSIST_ENDPOINT_LOCATION || 'global';
const GCP_LOCATION = process.env.GCP_LOCATION || 'global';
const ENTERPRISE_APP_ID = process.env.GEMINI_ENTERPRISE_APP_ID || process.env.VERTEX_DATASTORE_ID || '';
const ENTERPRISE_COLLECTION_ID = process.env.GEMINI_ENTERPRISE_COLLECTION_ID || 'default_collection';
const ENTERPRISE_ASSISTANT_ID = process.env.GEMINI_ENTERPRISE_ASSISTANT_ID || 'default_assistant';

const auth = new GoogleAuth({
  scopes: 'https://www.googleapis.com/auth/cloud-platform'
});

if (!PROJECT_ID) {
  console.warn('WARNING: GCP_PROJECT_ID environment variable is not set. Vertex AI client will use default credentials.');
}

// Pre-warmed global VertexAI client instance (connection pooling & token caching)
let vertexAI;
function getVertexAI() {
  if (!vertexAI) {
    vertexAI = new VertexAI({
      project: PROJECT_ID || 'genai-demo-catalog',
      location: REGION,
    });
  }
  return vertexAI;
}

const modelCache = new Map();

function getCachedModel(modelName, enableGrounding = true) {
  const name = modelName || MODEL_NAME;
  const cacheKey = `${name}_grounded_${enableGrounding}`;
  if (!modelCache.has(cacheKey)) {
    const modelConfig = {
      model: name,
      systemInstruction: {
        parts: [{ text: SYSTEM_INSTRUCTION_TEXT }]
      },
      generationConfig: {
        temperature: 0.4,
        maxOutputTokens: 8192
      }
    };

    if (enableGrounding && DATASTORE_ID) {
      modelConfig.tools = [
        {
          retrieval: {
            vertexAiSearch: {
              datastore: DATASTORE_ID
            }
          }
        }
      ];
    }

    const model = getVertexAI().getGenerativeModel(modelConfig);
    modelCache.set(cacheKey, model);
  }
  return { model: modelCache.get(cacheKey), name };
}

// In-memory session store for multi-turn chats
const sessionStore = new Map();

function getImageModel() {
  return getVertexAI().getGenerativeModel({
    model: IMAGE_MODEL_NAME
  });
}

async function generateVertexImage(prompt) {
  try {
    console.log('Generating image via Vertex AI (gemini-2.5-flash-image):', prompt.substring(0, 70) + '...');
    const res = await getImageModel().generateContent(
      `Create a professional corporate financial infographic chart reflecting exact data: ${prompt}. Use exact company name and exact numbers provided. DO NOT use generic placeholder names like "Global Innovations Inc.". Flat 2D vector graphic design, crisp typography, clean white background.`
    );
    const parts = (await res.response)?.candidates?.[0]?.content?.parts || [];
    const imgPart = parts.find(p => p.inlineData);
    if (imgPart && imgPart.inlineData?.data) {
      return `data:${imgPart.inlineData.mimeType || 'image/png'};base64,${imgPart.inlineData.data}`;
    }
  } catch (err) {
    console.warn('Vertex image generation error:', err.message);
  }
  return null;
}

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

// Inlines Vertex AI generated images into base64 data URIs
async function inlineImagesInContent(text) {
  if (!text) return '';
  let processed = text;

  // 1. Process ```image or ```imagen or ```visual blocks
  const codeBlockRegex = /```(?:image|imagen|visual)\s*([\s\S]*?)```/gi;
  const codeMatches = [...processed.matchAll(codeBlockRegex)];
  for (const match of codeMatches) {
    const [fullMatch, promptText] = match;
    const cleanPrompt = promptText.trim();
    if (cleanPrompt) {
      const dataUri = await generateVertexImage(cleanPrompt);
      if (dataUri) {
        const imgTag = `<div style="margin:18px 0; text-align:center;"><img src="${dataUri}" alt="Visual Chart" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 3px 10px rgba(0,0,0,0.08);" /></div>`;
        processed = processed.replace(fullMatch, imgTag);
      }
    }
  }

  // 2. Process ![alt](image: ...) with full balanced parenthesis support
  const startMarkerRegex = /!\[([^\]]*)\]\((?:image:|image-prompt:|imagen:)\s*/gi;
  let match;
  const itemsToReplace = [];

  while ((match = startMarkerRegex.exec(processed)) !== null) {
    const fullStart = match[0];
    const alt = match[1];
    const startIndex = match.index;
    const contentStartIndex = startIndex + fullStart.length;

    let openCount = 1;
    let i = contentStartIndex;
    while (i < processed.length && openCount > 0) {
      if (processed[i] === '(') openCount++;
      else if (processed[i] === ')') openCount--;
      i++;
    }

    if (openCount === 0) {
      const prompt = processed.substring(contentStartIndex, i - 1).trim();
      const rawMatch = processed.substring(startIndex, i);
      itemsToReplace.push({ rawMatch, alt, prompt });
    }
  }

  for (const item of itemsToReplace) {
    const dataUri = await generateVertexImage(item.prompt);
    if (dataUri) {
      const imgTag = `<div style="margin:18px 0; text-align:center;"><img src="${dataUri}" alt="${escapeXml(item.alt || 'Visual')}" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 3px 10px rgba(0,0,0,0.08);" /></div>`;
      processed = processed.replace(item.rawMatch, imgTag);
    } else {
      processed = processed.replace(item.rawMatch, '');
    }
  }

  // 3. Clean up any accidental malformed chart URLs and convert to images
  const malformedChartRegex = /!\[([^\]]*)\]\(https?:\/\/quickchart\.io\/chart[^\)]*\)/gi;
  const malformedMatches = [...processed.matchAll(malformedChartRegex)];
  for (const m of malformedMatches) {
    const [fullMatch, alt] = m;
    const fallbackPrompt = alt || "Corporate financial revenue and operating income comparison bar chart";
    const dataUri = await generateVertexImage(fallbackPrompt);
    if (dataUri) {
      const imgTag = `<div style="margin:18px 0; text-align:center;"><img src="${dataUri}" alt="${escapeXml(alt)}" style="max-width:100%; border-radius:6px; border:1px solid #c7e0f4; box-shadow:0 3px 10px rgba(0,0,0,0.08);" /></div>`;
      processed = processed.replace(fullMatch, imgTag);
    } else {
      processed = processed.replace(fullMatch, '');
    }
  }

  // 4. Catch any trailing prompt fragments (e.g. 'Clean white background... modern flat 2D...')
  processed = processed.replace(/(?:and operating income[^\n]*\.\s*)?Clean white background[^\n]*\)?/gi, '');
  processed = processed.replace(/['"],?data:,?backgroundColor:[^)]+\)\)?/gi, '');

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
  cors(req, res, async () => {
    try {
      if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
      }

      const { prompt, history, sessionId, model } = req.body;
      if (!prompt) {
        return res.status(400).json({ error: 'Prompt is required' });
      }

      const { enableGrounding = true } = req.body;
      const { model: generativeModel, name: selectedModel } = getCachedModel(model, enableGrounding);
      let activeSessionId = sessionId || ('session_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7));

      let formattedHistory = [];
      if (Array.isArray(history) && history.length > 0) {
        formattedHistory = history.map(item => {
          let role = item.role === 'assistant' ? 'model' : (item.role || 'user');
          let text = '';
          if (typeof item.text === 'string') {
            text = item.text;
          } else if (Array.isArray(item.parts)) {
            text = item.parts.map(p => typeof p === 'string' ? p : (p.text || '')).join('\n');
          } else if (typeof item.parts === 'string') {
            text = item.parts;
          }
          return { role: role, parts: [{ text: text.trim() }] };
        }).filter(item => item.parts[0].text.length > 0);
      }

      const chatSession = generativeModel.startChat({ history: formattedHistory });
      console.log('Sending prompt to Gemini (' + selectedModel + ') on Vertex AI (Session: ' + activeSessionId + ', Grounding: ' + enableGrounding + ')...');
      const result = await chatSession.sendMessage(prompt);
      const response = await result.response;
      
      const candidate = response.candidates?.[0];
      const rawParts = candidate?.content?.parts || [];
      let rawText = rawParts.map(p => p.text || '').join('\n').trim();

      // Extract enterprise grounding citations
      const citations = [];
      const groundingMetadata = candidate?.groundingMetadata;
      if (groundingMetadata?.groundingChunks?.length) {
        const seenSources = new Set();
        const citationItems = [];

        for (const chunk of groundingMetadata.groundingChunks) {
          const ctx = chunk.retrievedContext;
          if (ctx) {
            const key = ctx.title || ctx.uri;
            if (key && !seenSources.has(key)) {
              seenSources.add(key);
              const title = ctx.title || (ctx.uri ? ctx.uri.split('/').pop() : 'Document');
              const cleanUri = ctx.uri ? ctx.uri.replace('gs://', '') : '';
              citations.push({ title, uri: ctx.uri });
              citationItems.push(`* 📄 **${title}**${cleanUri ? ` \`(${cleanUri})\`` : ''}`);
            }
          }
        }

        if (citationItems.length > 0) {
          rawText += `\n\n---\n> [!NOTE] **Verified Sources from Advance Auto Parts Data Store:**\n> ` + citationItems.join('\n> ');
        }
      }

      const processedText = await inlineImagesInContent(rawText);
      const updatedHistory = await chatSession.getHistory();
      sessionStore.set(activeSessionId, updatedHistory);

      return res.status(200).json({
        result: processedText,
        sessionId: activeSessionId,
        history: updatedHistory,
        citations: citations,
        groundingMetadata: groundingMetadata || null
      });
    } catch (error) {
      console.error('Error in Gemini handler:', error);
      return res.status(500).json({
        error: 'Failed to process request',
        details: error.message,
      });
    }
  });
}

async function callStreamAssistAPI({ prompt, sessionId, userId, userPseudoId }) {
  if (!PROJECT_ID) {
    throw new Error('GCP_PROJECT_ID environment variable is required for StreamAssist');
  }
  if (!ENTERPRISE_APP_ID) {
    throw new Error('GEMINI_ENTERPRISE_APP_ID environment variable is required for StreamAssist');
  }

  const client = await auth.getClient();
  const accessTokenObj = await client.getAccessToken();
  const accessToken = typeof accessTokenObj === 'string' ? accessTokenObj : accessTokenObj.token;

  const endpointUrl = `https://${STREAM_ASSIST_ENDPOINT_LOCATION}-discoveryengine.googleapis.com/v1/projects/${PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/assistants/${ENTERPRISE_ASSISTANT_ID}:streamAssist`;

  const activeUserId = userPseudoId || userId || 'office_365_user';

  const requestBody = {
    query: {
      text: prompt
    }
  };

  if (sessionId) {
    let fullSessionName = sessionId;
    if (!sessionId.startsWith('projects/')) {
      fullSessionName = `projects/${PROJECT_ID}/locations/${GCP_LOCATION}/collections/${ENTERPRISE_COLLECTION_ID}/engines/${ENTERPRISE_APP_ID}/sessions/${sessionId}`;
    }
    requestBody.session = fullSessionName;
  }

  console.log(`Calling StreamAssist API (${endpointUrl})... Session: ${requestBody.session || 'NEW'}`);

  const apiRes = await fetch(endpointUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(requestBody)
  });

  if (!apiRes.ok) {
    const errText = await apiRes.text();
    console.error('StreamAssist API error:', apiRes.status, errText);
    throw new Error(`StreamAssist API returned HTTP ${apiRes.status}: ${errText}`);
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

  let aggregatedText = '';
  let returnedSessionResource = null;
  const citations = [];
  const seenCitations = new Set();

  for (const chunk of parsedChunks) {
    if (chunk.sessionInfo?.session) {
      returnedSessionResource = chunk.sessionInfo.session;
    }

    const replies = chunk.answer?.replies || [];
    for (const reply of replies) {
      const contentObj = reply.groundedContent?.content;
      if (contentObj && contentObj.text && !contentObj.thought) {
        aggregatedText += contentObj.text;
      } else if (reply.replyText) {
        aggregatedText += reply.replyText;
      }

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
      let chunkText = chunk.answer?.replyText || chunk.replyText || '';
      if (chunkText) {
        aggregatedText += chunkText;
      }
    }
  }

  let shortSessionId = returnedSessionResource;
  if (returnedSessionResource && returnedSessionResource.includes('/sessions/')) {
    shortSessionId = returnedSessionResource.split('/sessions/').pop();
  }

  return {
    resultText: aggregatedText,
    sessionResource: returnedSessionResource,
    sessionId: shortSessionId || sessionId,
    citations: citations
  };
}

async function handleGeminiEnterpriseRequest(req, res) {
  cors(req, res, async () => {
    try {
      if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed' });
      }

      const { prompt, sessionId, userId, userPseudoId } = req.body;
      if (!prompt) {
        return res.status(400).json({ error: 'Prompt is required' });
      }

      console.log(`Processing Gemini Enterprise request (Mode: ${BACKEND_MODE})... User: ${userPseudoId || userId || 'office_365_user'}`);

      if (BACKEND_MODE === 'streamassist' && ENTERPRISE_APP_ID) {
        try {
          const streamAssistResult = await callStreamAssistAPI({ prompt, sessionId, userId, userPseudoId });
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
          console.warn('StreamAssist execution failed, falling back to direct Vertex AI:', streamAssistErr.message);
        }
      }

      // Fallback to direct Vertex AI
      return handleGeminiRequest(req, res);
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
app.use(express.json());
app.use(cors);
app.post('/askGeminiEnterprise', handleGeminiEnterpriseRequest);
app.post('/askGemini', handleGeminiRequest);
app.post('/', handleGeminiEnterpriseRequest);
app.get('/', (req, res) => res.status(200).send('Gemini Enterprise Proxy Service Healthy'));

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Gemini Enterprise Proxy HTTP Server running on port ${PORT}`);
});


