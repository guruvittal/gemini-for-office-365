/**
 * Gemini Enterprise API Client for Microsoft 365
 * 
 * @author Sathya AG, Principal Architect, Google
 */

export function getActiveProxyUrl() {
  return 'https://gemini-proxy-j43mxpthfa-uc.a.run.app/askGeminiEnterprise';
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

export function getGoogleAccessToken() {
  if (typeof window !== 'undefined' && window.localStorage) {
    return window.localStorage.getItem('google_access_token') || '';
  }
  return '';
}

export function setGoogleAccessToken(token) {
  if (typeof window !== 'undefined' && window.localStorage) {
    if (token) {
      window.localStorage.setItem('google_access_token', token.trim());
    } else {
      window.localStorage.removeItem('google_access_token');
    }
  }
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

  const token = getGoogleAccessToken();
  const headers = { 'Content-Type': 'application/json' };
  if (token) {
    headers['Authorization'] = `Bearer ${token}`;
    payload.userAccessToken = token;
  }

  console.log(`Sending request to proxy endpoint: ${functionUrl} (Auth: ${token ? 'USER_TOKEN' : 'ANONYMOUS'})`);

  const response = await fetch(functionUrl, {
    method: 'POST',
    headers: headers,
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.details || errorData.error || `Server returned status ${response.status}`);
  }

  return await response.json();
}

