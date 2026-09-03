/**
 * Gemini Enterprise API Client for Microsoft 365
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { getOfficeAuthToken, getUserProfile, getGoogleAccessToken } from './authService.js';

// Default endpoint points directly to the secure auth-proxy gateway
const DEFAULT_AUTH_PROXY_URL = 'https://auth-proxy-16933400417.us-central1.run.app/askGeminiEnterprise';

export function getActiveProxyUrl() {
  if (typeof window !== 'undefined' && window.localStorage) {
    const override = window.localStorage.getItem('gemini_proxy_url');
    if (override) return override;
  }
  if (typeof process !== 'undefined' && process.env && process.env.GEMINI_PROXY_URL) {
    return process.env.GEMINI_PROXY_URL;
  }
  if (typeof window !== 'undefined' && window.location && window.location.hostname) {
    const host = window.location.hostname;
    if (host.includes('1062675944253') || host.includes('agentspace-wif')) {
      return 'https://auth-proxy-1062675944253.us-central1.run.app/askGeminiEnterprise';
    }
    if (host.includes('16933400417') || host.includes('agentspace-452714')) {
      return 'https://auth-proxy-16933400417.us-central1.run.app/askGeminiEnterprise';
    }
  }
  return DEFAULT_AUTH_PROXY_URL;
}

export function setProxyUrlOverride(url) {
  if (typeof window !== 'undefined' && window.localStorage) {
    if (url) {
      window.localStorage.setItem('gemini_proxy_url', url);
    } else {
      window.localStorage.removeItem('gemini_proxy_url');
    }
  }
}

export async function askGeminiEnterprise(prompt, history = [], sessionId = null, enableGrounding = true) {
  const functionUrl = getActiveProxyUrl();
  const userProfile = getUserProfile();

  // 1. Acquire Microsoft Entra ID SSO token
  let authToken = null;
  try {
    authToken = await getOfficeAuthToken();
  } catch (authErr) {
    console.warn('Proceeding without SSO token (server may reject if REQUIRE_ENTRA_AUTH=true):', authErr);
  }

  // 2. Check for 3-legged Google User OAuth token (for Google Drive grounding without DWD)
  const googleUserToken = getGoogleAccessToken();

  const payload = { 
    prompt: prompt,
    history: history,
    enableGrounding: enableGrounding,
    userPseudoId: userProfile.email || userProfile.user_id || 'office_365_user'
  };
  if (sessionId) {
    payload.sessionId = sessionId;
  }

  const headers = { 
    'Content-Type': 'application/json' 
  };
  if (authToken) {
    headers['Authorization'] = `Bearer ${authToken}`;
  }
  if (googleUserToken) {
    headers['X-End-User-Google-Token'] = googleUserToken;
    console.log('Attaching 3-legged Google User Token to request header.');
  }

  console.log(`Sending authenticated request to proxy endpoint: ${functionUrl}`);

  const response = await fetch(functionUrl, {
    method: 'POST',
    headers: headers,
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.detail || errorData.details || errorData.error || `Server returned status ${response.status}`);
  }

  return await response.json();
}

