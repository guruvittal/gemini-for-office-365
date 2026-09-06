/**
 * PowerPoint Prompt Enhancer
 * 
 * Cleanly handles PowerPoint prompt interactions without polluting the user's visible prompt text.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

/**
 * Cleanly returns the user prompt without injecting visible system rules into the chat input.
 */
export function enhancePromptForPowerPoint(userPrompt) {
  const lowerPrompt = userPrompt.toLowerCase();

  // Rule set for executive visuals (metric grid, before/after comparison)
  if (lowerPrompt.includes("executive visual") || lowerPrompt.includes("metric grid") || lowerPrompt.includes("comparison card") || lowerPrompt.includes("before/after") || lowerPrompt.includes("before and after")) {
    const rules = `
CRITICAL INSTRUCTIONS FOR EXECUTIVE VISUAL GENERATION:
1. Provide EXACTLY ONE definitive visual design. NEVER output multiple options.
2. DO NOT output conversational preamble, pleasantries, or conclusions.
3. For Metric Grids, output a structured JSON code block with "visualType": "metric_grid_3col" containing "title", "subtitle", and an array of 3 "cards", each with "metric", "title", "subtitle", and "bullets".
4. For Before/After Comparisons, output a structured JSON code block with "visualType": "before_after" containing "title", "subtitle", "before" ({ "title", "bullets" }), and "after" ({ "title", "bullets" }).
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for editing / shortening / making smaller / rewriting
  if (lowerPrompt.includes("shorten") || lowerPrompt.includes("smaller") || lowerPrompt.includes("concise") || lowerPrompt.includes("rewrite") || lowerPrompt.includes("punchy") || lowerPrompt.includes("fluff") || lowerPrompt.includes("trim")) {
    const rules = `
CRITICAL INSTRUCTIONS FOR SLIDE EDITING:
1. Provide EXACTLY ONE definitive, finalized version. NEVER provide multiple alternative options, variations, or choices (e.g. NEVER output "Option 1", "Option 2", or "Or, for an even more minimalist layout...").
2. DO NOT output conversational preamble, introduction, or pleasantries (e.g. NEVER output "Here is a concise and punchy version...", "Sure!", "Here are your revised bullets:").
3. DO NOT output conversational sign-offs or questions (e.g. "Let me know if you need changes").
4. Output ONLY the finalized slide content. Format each point as a clean bullet point with a bold lead-in phrase:
   • **Key Theme**: Crisp, high-impact description.
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for chart generation (pie, bar, line, doughnut, column, graph, breakdown)
  if (lowerPrompt.includes("chart") || lowerPrompt.includes("pie") || lowerPrompt.includes("bar") || lowerPrompt.includes("graph") || lowerPrompt.includes("visualization") || lowerPrompt.includes("visualize") || lowerPrompt.includes("plot") || lowerPrompt.includes("breakdown") || lowerPrompt.includes("doughnut") || lowerPrompt.includes("column")) {
    const rules = `
CRITICAL INSTRUCTIONS FOR CHART GENERATION:
1. Output a structured JSON code block with the exact data metrics:
\`\`\`json
{
  "chartType": "pie",
  "title": "Chart Title",
  "data": [
    { "label": "Category / Label", "value": 12345 }
  ]
}
\`\`\`
Supported chartType values: "pie", "doughnut", "bar", "column", "line". Use exact numeric values (not strings).
2. Also provide a clean Markdown Table with the data metrics (| Category | Metric | Share % |).
3. Provide 2-3 executive bullet points with bold lead-ins highlighting strategic insights.
4. DO NOT output conversational preamble or pleasantries.
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for Executive Summary (strictly 1 slide only)
  if (lowerPrompt.includes("summarize") || lowerPrompt.includes("executive summary") || lowerPrompt.includes("slide summary") || lowerPrompt.includes("key takeaway")) {
    const rules = `
CRITICAL INSTRUCTIONS FOR EXECUTIVE SUMMARY GENERATION:
1. Provide EXACTLY ONE single slide. NEVER generate multiple slides or multiple '##' slide headings under any circumstances.
2. Structure with exactly one slide title: "## 📊 Executive Slide Summary".
3. Visual Chart: If the summary contains quantitative breakdown, comparisons, or metrics, output a structured JSON code block with the exact data metrics (chartType: "doughnut", title: "...", data: [...]) so our client presentation engine can render a crisp, high-resolution chart.
4. Content layout:
   - Provide a clean Markdown Table (| Category | Metric | Share / Value |) summarizing quantitative data or metrics.
   - Followed by 2 to 3 concise, high-impact executive takeaway bullets with bold lead-in phrases.
5. DO NOT output conversational preamble, pleasantries, or additional slides.
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for slide generation
  if (lowerPrompt.includes("slide") || lowerPrompt.includes("presentation") || lowerPrompt.includes("deck") || lowerPrompt.includes("table") || lowerPrompt.includes("pitch")) {
    const rules = `
IMPORTANT RULES FOR SLIDE GENERATION:
1. Provide EXACTLY ONE definitive presentation version. DO NOT output multiple alternatives or options.
2. DO NOT output conversational preamble or filler (e.g. "Here is...", "Sure!"). Output the presentation content directly.
3. Structure your response clearly using Markdown Headings (e.g. ## Slide 1: [Emoji] [Title]) for each slide.
4. For each slide, provide:
   - A short, punchy **Title** of **MAXIMUM 3 TO 4 WORDS (under 40 characters)** prefixed with a relevant **Emoji / Icon** (e.g., "📊 Financial Highlights", "🚀 Growth Strategy", "💰 Capital & Resources", "📈 Outlook & Guidance"). Put extra details (like dates or quarters) into the Subtitle.
   - A **Subtitle** (if applicable, clearly labeled).
   - **Main Content**:
     * **STRUCTURED TABLES FOR QUANTITATIVE & COMPARATIVE DATA**: When presenting dense financial results, multi-attribute comparisons, or numeric metrics, format as a clean Markdown table (e.g. | Metric | Q1 2026 | YoY Change | Impact |).
     * **EXECUTIVE BULLETS FOR STRATEGY & NARRATIVE**: For strategic vision, qualitative analysis, key initiatives, risks, and next steps, use 3 to 4 crisp, high-impact bullet points with bold lead-ins. Do NOT force a table when narrative bullets convey the insight better.
5. Do NOT output internal design metadata, font sizes (like "Title Size: 44"), hex colors (like "Color: #..."), or raw "Visual Concept:" labels. Keep the output clean, executive-ready presentation content.
6. Provide all slides in a single response, cleanly separated by headings.
`;
    return `${userPrompt}\n\n${rules}`;
  }
  return userPrompt;
}

/**
 * Initializes prompt enhancer (no-op to prevent polluting UI input text).
 */
export function initPromptEnhancer() {
  // Kept clean to avoid displaying bracketed formatting rules to the user
}
