/**
 * Gemini Enterprise API Client for Microsoft 365
 * 
 * @author Sathya AG, Principal Architect, Google
 */

export function getActiveProxyUrl() {
  if (typeof window !== 'undefined' && window.location) {
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('backend') === 'streamassist') {
      return 'https://gemini-enterprise-proxy-133594738129.us-central1.run.app/askGeminiEnterprise';
    }
  }
  return 'https://us-central1-genai-demo-catalog.cloudfunctions.net/askGemini';
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

function getOfficeUserId() {
  try {
    if (typeof Office !== 'undefined' && Office.context) {
      if (Office.context.user && Office.context.user.email) {
        return Office.context.user.email;
      }
      if (Office.context.mailbox && Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) {
        return Office.context.mailbox.userProfile.emailAddress;
      }
    }
  } catch (e) {
    console.warn('Could not read Office user context:', e);
  }
  if (typeof window !== 'undefined' && window.localStorage) {
    let localUserId = window.localStorage.getItem('gemini_user_pseudo_id');
    if (!localUserId) {
      localUserId = 'office_user_' + Math.random().toString(36).substring(2, 10);
      window.localStorage.setItem('gemini_user_pseudo_id', localUserId);
    }
    return localUserId;
  }
  return 'office_365_user';
}

export async function askGeminiEnterprise(prompt, history = [], sessionId = null, enableGrounding = true) {
  const functionUrl = getActiveProxyUrl();

  const payload = { 
    prompt: prompt,
    history: history,
    enableGrounding: enableGrounding,
    userPseudoId: getOfficeUserId()
  };
  if (sessionId) {
    payload.sessionId = sessionId;
  }

  console.log(`Sending request to proxy endpoint: ${functionUrl}`);

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

