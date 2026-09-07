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
    const wordToNumber = {
      "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
      "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10
    };
    const countMatch = lowerPrompt.match(/\b(?:create|generate|make|build|provide|give\s+me)?\s*(\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s+slides?\b/i);
    let requestedCount = null;
    if (countMatch && countMatch[1]) {
      const token = countMatch[1].toLowerCase();
      requestedCount = wordToNumber[token] || parseInt(token, 10);
    }

    let countConstraint = "";
    if (requestedCount && requestedCount > 0) {
      countConstraint = `
CRITICAL CONSTRAINT - EXACT SLIDE COUNT:
The user explicitly requested EXACTLY ${requestedCount} slides. You MUST generate EXACTLY ${requestedCount} slides (from ## Slide 1 to ## Slide ${requestedCount}). NEVER output fewer or more than ${requestedCount} slides under any circumstances.
`;
    }

    const rules = `
CRITICAL STRUCTURE CONTRACT FOR SLIDE GENERATION:${countConstraint}
1. Provide EXACTLY ONE definitive presentation version. DO NOT output multiple alternatives, conversational preamble, pleasantries, or filler. Output the presentation slides directly.
2. For EVERY slide, strictly follow this standardized structure:

---
## Slide {N}: [Relevant Emoji] [Title: MAX 3 TO 4 WORDS]
### [Contextual Subtitle / One-sentence Takeaway]

[CONTENT AREA - CHOOSE EXACTLY ONE VISUAL OR CONTENT FORMAT PER SLIDE]:
- For narrative, strategy, vision, or background: 3 to 4 crisp executive bullet points with bold lead-ins (• **Key Theme**: Impactful description).
- For quantitative data / financial metrics: An Executive Markdown Table with 3 to 4 columns (| Dimension | Metric | Strategic Impact |). Do NOT add extra bullets or takeaways on a table slide - the table itself represents the content. (NEVER put a chart on a table slide!).
- For distribution / breakdown / share metrics: A structured JSON code block (\`\`\`json { "chartType": "doughnut|bar|pie", "title": "...", "data": [...] } \`\`\`) accompanied by 2-3 narrative bullet points. (NEVER put a table on a chart slide!).
- For strategic priorities or key pillars: 3 to 4 crisp executive bullet points with bold lead-ins.

3. CRITICAL VISUAL RULES:
- PRESENTATION DIVERSITY (Decks with 3+ slides): NEVER make every slide a plain bullet slide! A high-performing executive presentation MUST use varied layouts across slides:
  * Slide 1: Executive introduction or strategic overview (clean title, subtitle, and 3 to 4 structured narrative bullets with bold lead-ins).
  * Data & Metrics Slide: At least 1 slide MUST feature an Executive Markdown Table (3-4 columns) showing core metrics, dimensions, and strategic impacts.
  * Optional Chart Slide: At most 1 slide can feature a Data Chart (\`\`\`json { "chartType": "doughnut|bar|pie" ... } \`\`\`).
  * Remaining Slides: Strategic pillars, roadmap, or narrative bullets with bold lead-ins.
- NEVER OUTPUT PSEUDO-TEXT LABELS: NEVER output labels like "Metric Grid", "Comparison Card", or raw placeholder text in the slide content.
- EACH CHART MUST BE 100% UNIQUE: NEVER repeat or duplicate the same chart, metrics, or title across multiple slides. If a deck has 5 to 10 slides, at most 2 should contain a data chart, and each MUST cover a completely different topic and metric.
- NEVER COMBINE A CHART AND A TABLE ON THE SAME SLIDE: A slide must feature EITHER a table OR a chart, NEVER both.
- SLIDE 1 MUST BE AN INTRODUCTION SLIDE: Slide 1 should contain a clean title, subtitle, and 3 to 4 introductory narrative bullet points. Do NOT put charts or tables on Slide 1.
4. Do NOT output internal design metadata, font sizes (like "Title Size: 44"), hex colors (like "Color: #..."), or raw "Visual Concept:" labels. Keep the output clean, executive-ready presentation content.
5. Separate every slide cleanly with a horizontal rule "---".
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
