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
   - A short, punchy **Title** of **MAXIMUM 3 TO 4 WORDS (under 40 characters)** prefixed with a relevant **Emoji / Icon** (e.g., "📊 Executive Summary", "🚀 Growth Strategy", "🌍 Demographics", "📌 Financial Highlights"). Put extra details (like dates or quarters) into the Subtitle.
   - A **Subtitle** (if applicable, clearly labeled).
   - **Main Content** (use bullet points for readability, or a markdown table for structured comparison data).
   - **Visual Concept** (describe the recommended image/chart, labeled "Visual Concept:").
   - Recommended **Theme Color** (labeled "Color:").
   - Recommended **Title Font Size** and **Subtitle Font Size** (e.g., "Title Size: 44", "Subtitle Size: 24").
4. If presenting structured comparison data or metrics, format the main content as a clean Markdown table.
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
