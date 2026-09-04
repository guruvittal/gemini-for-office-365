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
  // If the user is asking to create, generate, or make slides, append structured formatting rules.
  const lowerPrompt = userPrompt.toLowerCase();
  if (lowerPrompt.includes("slide") || lowerPrompt.includes("presentation") || lowerPrompt.includes("deck") || lowerPrompt.includes("table") || lowerPrompt.includes("pitch")) {
    const rules = `
IMPORTANT RULES FOR SLIDE GENERATION:
1. Do NOT use the canvas feature. Provide the complete content directly in your text response.
2. Structure your response clearly using Markdown Headings (e.g. ## Slide 1: [Emoji] [Title]) for each slide.
3. For each slide, provide:
   - A short, punchy **Title** of **MAXIMUM 3 TO 4 WORDS (under 40 characters)** prefixed with a relevant **Emoji / Icon** (e.g., "📊 Financial Highlights", "🚀 Growth Strategy", "💰 Capital & Resources", "📈 Outlook & Guidance"). Put extra details (like dates or quarters) into the Subtitle.
   - A **Subtitle** (if applicable, clearly labeled).
   - **Main Content**:
     * **MANDATORY MARKDOWN TABLES FOR DATA & METRICS**: Whenever presenting financial results (e.g., revenue, EPS, operating income, margins), key performance metrics, comparisons, or structured data, you MUST format the main content as a clean Markdown table (e.g. | Metric | Q1 2026 | YoY Change | Impact |). Tables are essential for executive financial and KPI slides.
     * **BULLETS FOR STRATEGY & NARRATIVE**: For strategy, vision, narrative, or qualitative discussion, use crisp, high-impact bullet points.
4. Do NOT output internal design metadata, font sizes (like "Title Size: 44"), hex colors (like "Color: #..."), or raw "Visual Concept:" labels. Keep the output clean, executive-ready presentation content.
5. Provide all slides in a single response, cleanly separated by headings.
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
