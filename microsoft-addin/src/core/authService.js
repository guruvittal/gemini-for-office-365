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
