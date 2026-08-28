/**
 * Authentication Service for Microsoft Office 365 Add-in SSO
 * 
 * Handles silent token acquisition via Office.auth.getAccessToken(),
 * token caching, fallback consent dialogs, and JWT payload parsing.
 * 
 * @author Sathya AG & Antigravity Team
 */

let cachedToken = null;
let tokenExpiry = 0;
let cachedUserProfile = null;

/**
 * Decodes the base64url payload of a JWT without verifying signature
 * (signature verification is performed securely on the auth-proxy backend).
 */
export function decodeJwtPayload(token) {
  if (!token || typeof token !== 'string') return null;
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const base64Url = parts[1];
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload = decodeURIComponent(
      atob(base64)
        .split('')
        .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
        .join('')
    );
    return JSON.parse(jsonPayload);
  } catch (err) {
    console.warn('Could not decode JWT payload:', err);
    return null;
  }
}

/**
 * Retrieves the signed-in user's profile information from Office.js or decoded JWT claims.
 */
export function getUserProfile() {
  if (cachedUserProfile) return cachedUserProfile;

  // 1. If we have a cached JWT token, extract from claims
  if (cachedToken) {
    const claims = decodeJwtPayload(cachedToken);
    if (claims) {
      cachedUserProfile = {
        name: claims.name || claims.preferred_username || claims.upn || 'Corporate User',
        email: claims.preferred_username || claims.email || claims.upn || 'user@contoso.com',
        user_id: claims.oid || claims.sub || 'entra_user',
        tenant_id: claims.tid || null,
        roles: claims.roles || [],
        is_authenticated: true
      };
      return cachedUserProfile;
    }
  }

  // 2. Fallback to Office.context user if available
  try {
    if (typeof Office !== 'undefined' && Office.context) {
      if (Office.context.user && Office.context.user.displayName) {
        return {
          name: Office.context.user.displayName,
          email: Office.context.user.email || 'user@office.com',
          user_id: Office.context.user.accountId || 'office_user',
          is_authenticated: false
        };
      }
      if (Office.context.mailbox && Office.context.mailbox.userProfile) {
        const up = Office.context.mailbox.userProfile;
        return {
          name: up.displayName || 'Corporate User',
          email: up.emailAddress || 'user@office.com',
          user_id: up.emailAddress || 'office_user',
          is_authenticated: false
        };
      }
    }
  } catch (e) {
    console.debug('Office context user inspection skipped:', e);
  }

  return {
    name: 'Office 365 User',
    email: 'user@organization.com',
    user_id: 'dev_user',
    is_authenticated: false
  };
}

let lastAuthError = null;

export function getLastAuthError() {
  return lastAuthError;
}

/**
 * Retrieves the Microsoft Entra ID JWT access token via Office.js SSO.
 * Automatically handles caching and silent acquisition.
 * 
 * @param {boolean} forceRefresh - If true, bypasses in-memory cache
 * @returns {Promise<string|null>} Entra ID JWT access token or null in dev/fallback mode
 */
export async function getOfficeAuthToken(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && cachedToken && now < tokenExpiry) {
    return cachedToken;
  }

  if (typeof Office === 'undefined' || !Office.auth || !Office.auth.getAccessToken) {
    console.warn('Office.auth.getAccessToken is not available in this environment. Running in unauthenticated / dev mode.');
    lastAuthError = { code: 'NO_OFFICE_AUTH', message: 'Office.auth API not available in this host environment.' };
    return null;
  }

  try {
    console.log('Acquiring Microsoft Entra ID SSO token via Office.auth.getAccessToken()...');
    const token = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: false
    });

    if (token) {
      cachedToken = token;
      lastAuthError = null;
      // Cache for 50 minutes (Entra ID tokens typically valid for 60 minutes)
      tokenExpiry = now + (50 * 60 * 1000);
      
      // Update cached user profile
      const claims = decodeJwtPayload(token);
      if (claims) {
        cachedUserProfile = {
          name: claims.name || claims.preferred_username || claims.upn || 'Corporate User',
          email: claims.preferred_username || claims.email || claims.upn || 'user@contoso.com',
          user_id: claims.oid || claims.sub || 'entra_user',
          tenant_id: claims.tid || null,
          roles: claims.roles || [],
          is_authenticated: true
        };
      }
      
      console.log('Successfully acquired Entra ID SSO token for user:', cachedUserProfile ? cachedUserProfile.email : 'authenticated');
      return token;
    }
  } catch (error) {
    lastAuthError = error;
    console.error('Office SSO token acquisition failed:', error);

    // Specific Office SSO error code diagnostics
    if (error.code === 13001) {
      console.warn('SSO Error 13001: User is not signed into Office with a Microsoft Entra ID account.');
    } else if (error.code === 13002) {
      console.warn('SSO Error 13002: User cancelled the consent dialog.');
    } else if (error.code === 13003) {
      console.warn('SSO Error 13003: User type is not supported (e.g. personal Microsoft Account).');
    } else if (error.code === 13007) {
      console.warn('SSO Error 13007: Invalid Application ID URI or untrusted client application (Manifest Resource domain must match hosting domain).');
    } else if (error.code === 13012) {
      console.warn('SSO Error 13012: Precondition failed (e.g. platform API not supported).');
    }

    // Return null so the client can attempt calling the proxy in fallback mode if allowed
    return null;
  }

  return null;
}

// ---------------------------------------------------------------------------
// Google 3-Legged OAuth Management (No DWD Required)
// ---------------------------------------------------------------------------

let cachedGoogleToken = null;
let googleTokenExpiry = 0;
let cachedAppConfig = null;
let googleAuthDialog = null;

export function getProxyBaseUrl() {
  if (typeof window !== 'undefined' && window.localStorage) {
    const override = window.localStorage.getItem('gemini_proxy_url');
    if (override) return override.replace(/\/askGeminiEnterprise$/, '');
  }
  if (typeof process !== 'undefined' && process.env && process.env.GEMINI_PROXY_URL) {
    return process.env.GEMINI_PROXY_URL.replace(/\/askGeminiEnterprise$/, '');
  }
  if (typeof window !== 'undefined' && window.location && window.location.hostname) {
    const host = window.location.hostname;
    if (host.includes('1062675944253') || host.includes('agentspace-wif')) {
      return 'https://auth-proxy-1062675944253.us-central1.run.app';
    }
    if (host.includes('16933400417') || host.includes('agentspace-452714')) {
      return 'https://auth-proxy-16933400417.us-central1.run.app';
    }
  }
  return 'https://auth-proxy-16933400417.us-central1.run.app';
}

/**
 * Loads dynamic frontend configuration from the backend auth-proxy (/api/config)
 * so that Google OAuth Client IDs and settings are never hardcoded in client code.
 */
export async function fetchAppConfig() {
  if (cachedAppConfig) return cachedAppConfig;

  // Check localStorage override
  if (typeof window !== 'undefined' && window.localStorage) {
    const localClientId = window.localStorage.getItem('google_oauth_client_id');
    if (localClientId) {
      cachedAppConfig = { google_oauth_client_id: localClientId };
      return cachedAppConfig;
    }
  }

  try {
    const configUrl = `${getProxyBaseUrl()}/api/config`;
    const resp = await fetch(configUrl);
    if (resp.ok) {
      cachedAppConfig = await resp.json();
      console.log('Successfully fetched dynamic app config:', {
        google_client_id_configured: bool(cachedAppConfig?.google_oauth_client_id),
        user_auth_mode: cachedAppConfig?.user_auth_mode
      });
      return cachedAppConfig;
    }
  } catch (err) {
    console.warn('Could not fetch backend app config:', err);
  }

  return cachedAppConfig || {};
}

function bool(val) {
  return !!val;
}

/**
 * Retrieves the configured Google OAuth 2.0 Web Client ID dynamically.
 */
export async function getGoogleOAuthClientId() {
  if (typeof window !== 'undefined' && window.localStorage) {
    const stored = window.localStorage.getItem('google_oauth_client_id');
    if (stored) return stored;
  }
  const config = await fetchAppConfig();
  return config?.google_oauth_client_id || '';
}

/**
 * Stores Google OAuth access token with expiration in memory and sessionStorage.
 */
export function setGoogleAccessToken(token, expiresIn = 3600) {
  if (!token) {
    cachedGoogleToken = null;
    googleTokenExpiry = 0;
    if (typeof window !== 'undefined' && window.sessionStorage) {
      window.sessionStorage.removeItem('google_user_access_token');
      window.sessionStorage.removeItem('google_user_token_expiry');
    }
    return;
  }

  cachedGoogleToken = token;
  const now = Date.now();
  // Safe buffer: expire 5 minutes earlier
  const ttlMs = Math.max((expiresIn - 300) * 1000, 60000);
  googleTokenExpiry = now + ttlMs;

  if (typeof window !== 'undefined' && window.sessionStorage) {
    window.sessionStorage.setItem('google_user_access_token', token);
    window.sessionStorage.setItem('google_user_token_expiry', String(googleTokenExpiry));
  }
  console.log('Google user access token stored. Valid for ~' + Math.round(ttlMs / 60000) + ' minutes.');
}

/**
 * Retrieves the active Google user OAuth token if valid.
 */
export function getGoogleAccessToken() {
  const now = Date.now();
  if (cachedGoogleToken && now < googleTokenExpiry) {
    return cachedGoogleToken;
  }

  // Check sessionStorage
  if (typeof window !== 'undefined' && window.sessionStorage) {
    const stored = window.sessionStorage.getItem('google_user_access_token');
    const exp = parseInt(window.sessionStorage.getItem('google_user_token_expiry') || '0', 10);
    if (stored && now < exp) {
      cachedGoogleToken = stored;
      googleTokenExpiry = exp;
      return cachedGoogleToken;
    }
  }

  return null;
}

export function isGoogleTokenValid() {
  return !!getGoogleAccessToken();
}

/**
 * Initiates the 3-Legged Google OAuth Sign-In flow using Office Dialog API.
 * 
 * @param {string|null} loginHint - Email hint (e.g. scim@jeansson.demo.altostrat.com)
 * @param {string} prompt - OAuth prompt mode ('select_account' or 'none')
 * @returns {Promise<{status: string, token?: string, error?: string}>}
 */
export async function initiateGoogleSignIn(loginHint = null, prompt = 'select_account') {
  const clientId = await getGoogleOAuthClientId();
  if (!clientId) {
    const errMsg = 'Google OAuth Client ID is not configured. Please ensure GOOGLE_OAUTH_CLIENT_ID is set.';
    console.error(errMsg);
    return { status: 'error', error: errMsg };
  }

  // Pre-fill user email from profile if not provided
  if (!loginHint) {
    const profile = getUserProfile();
    if (profile && profile.email && profile.email !== 'user@organization.com') {
      loginHint = profile.email;
    }
  }

  return new Promise((resolve) => {
    const origin = typeof window !== 'undefined' && window.location ? window.location.origin : 'https://gemini-frontend-16933400417.us-central1.run.app';
    const authUrl = `${origin}/google-auth.html?client_id=${encodeURIComponent(clientId)}&login_hint=${encodeURIComponent(loginHint || '')}&prompt=${encodeURIComponent(prompt)}`;

    console.log('Opening Google OAuth dialog at:', authUrl);

    if (typeof Office !== 'undefined' && Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
      Office.context.ui.displayDialogAsync(
        authUrl,
        { height: 60, width: 40, promptBeforeOpen: false },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error('Failed to open Office Google Auth dialog:', asyncResult.error);
            resolve({ status: 'error', error: asyncResult.error.message });
            return;
          }

          googleAuthDialog = asyncResult.value;

          googleAuthDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            try {
              const data = JSON.parse(arg.message);
              if (data.google_token) {
                setGoogleAccessToken(data.google_token, data.expires_in || 3600);
                if (googleAuthDialog) {
                  googleAuthDialog.close();
                  googleAuthDialog = null;
                }
                resolve({ status: 'success', token: data.google_token });
              } else {
                if (googleAuthDialog) {
                  googleAuthDialog.close();
                  googleAuthDialog = null;
                }
                resolve({ status: 'error', error: data.error || 'Google authentication was not completed.' });
              }
            } catch (e) {
              console.error('Error handling dialog message:', e);
              if (googleAuthDialog) {
                googleAuthDialog.close();
                googleAuthDialog = null;
              }
              resolve({ status: 'error', error: e.message });
            }
          });

          googleAuthDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            console.warn('Dialog event / closed by user:', arg);
            googleAuthDialog = null;
            resolve({ status: 'error', error: 'Authentication window closed.' });
          });
        }
      );
    } else {
      // Standalone browser fallback popup
      const popup = window.open(authUrl, 'google_auth_popup', 'width=500,height=650');
      const messageHandler = (event) => {
        if (event.data && event.data.type === 'GOOGLE_AUTH_RESULT') {
          window.removeEventListener('message', messageHandler);
          const payload = event.data.payload;
          if (payload.google_token) {
            setGoogleAccessToken(payload.google_token, payload.expires_in || 3600);
            resolve({ status: 'success', token: payload.google_token });
          } else {
            resolve({ status: 'error', error: payload.error || 'Authentication failed' });
          }
        }
      };
      window.addEventListener('message', messageHandler);
    }
  });
}
