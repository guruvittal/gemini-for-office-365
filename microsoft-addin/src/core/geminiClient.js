/**
 * Gemini Enterprise API Client for Microsoft 365
 * 
 * @author Sathya AG, Principal Architect, Google
 */

export async function askGeminiEnterprise(prompt, history = [], sessionId = null, enableGrounding = true) {
  // Configured via webpack DefinePlugin / .env / runtime config
  const functionUrl = (typeof process !== 'undefined' && process.env && process.env.GEMINI_PROXY_URL)
    ? process.env.GEMINI_PROXY_URL
    : (window.GEMINI_PROXY_URL || 'https://us-central1-genai-demo-catalog.cloudfunctions.net/askGemini');

  const payload = { 
    prompt: prompt,
    history: history,
    enableGrounding: enableGrounding
  };
  if (sessionId) {
    payload.sessionId = sessionId;
  }

  const response = await fetch(functionUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.details || errorData.error || `Server returned status ${response.status}`);
  }

  return await response.json();
}
