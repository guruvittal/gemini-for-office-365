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

  // Rule set for image generation
  if (lowerPrompt.includes("image") || lowerPrompt.includes("picture") || lowerPrompt.includes("illustration") || lowerPrompt.includes("drawing") || lowerPrompt.includes("visual art")) {
    const rules = `
CRITICAL INSTRUCTIONS FOR IMAGE GENERATION:
1. Generate a professional high-quality corporate visual image illustration representing this concept.
2. DO NOT output conversational preamble, pleasantries, or conclusions.
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for Executive Summary (Summarize Slides allows up to 5 slides with dedicated takeaways slide; other summaries remain single-slide)
  if (lowerPrompt.includes("summarize") || lowerPrompt.includes("executive summary") || lowerPrompt.includes("slide summary") || lowerPrompt.includes("key takeaway")) {
    const isSummarizeSlides = (typeof window !== "undefined" && window.__isSummarizeSlidesAction) ||
      lowerPrompt.includes("summarize slides") ||
      lowerPrompt.includes("up to 5 slides") ||
      lowerPrompt.includes("[summarize slides]");

    if (isSummarizeSlides) {
      const rules = `
CRITICAL INSTRUCTIONS FOR SUMMARIZE SLIDES GENERATION (UP TO 5 SLIDES):
1. Provide a comprehensive Executive Summary presentation across multiple slides (UP TO 5 SLIDES):
   - You can create up to 5 slides to thoroughly cover the key information, data, metrics, comparisons, and strategic findings.
   - Separate distinct topics, tables, and visual charts into their own slides (e.g. ## Slide 1: [Executive Overview / Main Metrics Table], ## Slide 2: [Category Breakdown / Visual Chart / Details Table], etc.).
   - If there are multiple tables or data sets, place each table on its own appropriate slide.
   - If a visual chart represents data, output a structured JSON code block with the exact data metrics (chartType: "doughnut" or "bar", title: "...", data: [...]) so our client presentation engine can render a crisp chart.
   - Break down the key takeaways into a dedicated single slide titled "## 📊 Executive Summary: Key Takeaways" with impactful bullet points and bold lead-in phrases.
   - Format each slide with a clear markdown header (## Slide 1: [Title], ## Slide 2: [Title], etc.) so each section generates its own slide.
2. STRICT CLOSED-BOOK GROUNDING CONTRACT:
   - You are operating in 100% STRICT CLOSED-BOOK MODE based SOLELY on the provided slide context.
   - You must synthesize information ONLY AND EXCLUSIVELY from the text and data present in the selected slides.
   - NEVER extrapolate, bring in external industry knowledge, or introduce topics, facts, or assumptions that do not appear in the selected slides.
   - If a concept, fact, or metric is not explicitly stated in the selected slides, DO NOT mention it.
3. DO NOT output conversational preamble or pleasantries.
`;
      return `${userPrompt}\n\n${rules}`;
    }

    const rules = `
CRITICAL INSTRUCTIONS FOR EXECUTIVE SUMMARY GENERATION:
1. Provide a professional Executive Summary with EXACTLY ONE single slide:
   - "## 📊 Executive Slide Summary" containing the high-level context and structured Markdown Table (| Metric / Focus Area | FY Progress Status | Target Benchmark |) summarizing key data directly from the slides.
   - Beneath the table, provide concise bullet points / key takeaways synthesizing the core findings.
   - STRICT SINGLE-SLIDE CONSTRAINT: Do NOT generate multiple slides or a separate "Key Takeaways" slide.
2. STRICT CLOSED-BOOK GROUNDING CONTRACT:
   - You are operating in 100% STRICT CLOSED-BOOK MODE based SOLELY on the provided slide context.
   - You must synthesize information ONLY AND EXCLUSIVELY from the text and data present in the selected slides.
   - NEVER extrapolate, bring in external industry knowledge, or introduce topics, facts, or assumptions that do not appear in the selected slides.
   - If a concept, fact, or metric is not explicitly stated in the selected slides, DO NOT mention it.
3. Visual Chart: If the summary contains quantitative breakdown, comparisons, or metrics directly stated in the slides, output a structured JSON code block with the exact data metrics (chartType: "doughnut", title: "...", data: [...]) so our client presentation engine can render a crisp, high-resolution chart.
4. Content layout:
   - If the slides contain structured comparisons or metrics, provide a clean Markdown Table (| Category | Metric | Share / Value |) summarizing data directly from the slides.
   - Place all executive takeaway bullets (2 to 4 concise bullets with bold lead-ins) beneath the table.
5. DO NOT output conversational preamble or pleasantries.
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
  * Slide 1: Dedicated Title and Executive Summary Slide containing ONLY a clean Presentation Title, Subtitle, and ONE small Executive Summary paragraph (2-4 sentences max) explaining what the document is all about. NEVER include "Deck Scope & Outline", upcoming slide lists (Slide 2, Slide 3, etc.), operational bullets, metrics, charts, or tables on Slide 1.
  * Data & Metrics Slide: At least 1 slide MUST feature an Executive Markdown Table (3-4 columns) showing core metrics, dimensions, and strategic impacts.
  * Optional Chart Slide: At most 1 slide can feature a Data Chart (\`\`\`json { "chartType": "doughnut|bar|pie" ... } \`\`\`).
  * Remaining Slides: Strategic pillars, roadmap, or narrative bullets with bold lead-ins.
- NEVER OUTPUT PSEUDO-TEXT LABELS: NEVER output labels like "Metric Grid", "Comparison Card", or raw placeholder text in the slide content.
- EACH CHART MUST BE 100% UNIQUE: NEVER repeat or duplicate the same chart, metrics, or title across multiple slides. If a deck has 5 to 10 slides, at most 2 should contain a data chart, and each MUST cover a completely different topic and metric.
- CRITICAL CHART TITLE RULE: Every chart title MUST be short and punchy (maximum 2 to 5 words). NEVER create long, rambling titles or include parenthetical details in the title.
- NEVER COMBINE A CHART AND A TABLE ON THE SAME SLIDE: A slide must feature EITHER a table OR a chart, NEVER both.
- SLIDE 1 MUST BE A TITLE & SUMMARY SLIDE: Slide 1 must contain ONLY the presentation title, subtitle, and ONE single executive summary small paragraph. Save specific operational details, findings, metrics, data tables, and pillars for Slides 2 through 6.
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
