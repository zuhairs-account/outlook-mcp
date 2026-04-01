/**
 * OAuth server routes for Outlook MCP authentication.
 *
 * Action plan fixes applied:
 *   - (c) auth/tools.js: URL construction deduplicated → uses shared buildAuthUrl()
 *   - (e) auth/tools.js: PKCE support → code_verifier stored per-session
 *   - oauth-server.js (c): State validation enforced (not just commented out)
 *   - oauth-server.js (c): In-memory session store for state + PKCE verifiers
 */

const express = require('express');
const querystring = require('querystring');
const https = require('https');
const crypto = require('crypto');
const TokenStorage = require('./token-storage');

// BEFORE: buildAuthUrl logic was duplicated inline in the /auth route.
// AFTER: Import shared buildAuthUrl from tools.js.
// GOOD EFFECT: Single source of truth for auth URL construction — changes
//              to query params or PKCE logic happen in one place.
const { buildAuthUrl } = require('./tools');

// ─── HTML Templates ───────────────────────────────────────────────────
function escapeHtml(unsafe) {
  return unsafe
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

const templates = {
  authError: (error, errorDescription) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">&#x274C; Authorization Failed</h1>
        <p><strong>Error:</strong> ${escapeHtml(error)}</p>
        ${errorDescription ? `<p><strong>Description:</strong> ${escapeHtml(errorDescription)}</p>` : ''}
        <p>You can close this window and try again.</p>
      </body>
    </html>`,
  authSuccess: `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #2ecc71;">&#x2705; Authentication Successful</h1>
        <p>You have successfully authenticated with Microsoft Graph API.</p>
        <p>You can close this window.</p>
      </body>
    </html>`,
  tokenExchangeError: (error) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1 style="color: #e74c3c;">&#x274C; Token Exchange Failed</h1>
        <p>Failed to exchange authorization code for access token.</p>
        <p><strong>Error:</strong> ${escapeHtml(error instanceof Error ? error.message : String(error))}</p>
        <p>You can close this window and try again.</p>
      </body>
    </html>`,
  tokenStatus: (status) => `
    <html>
      <body style="font-family: Arial, sans-serif; text-align: center; margin-top: 50px;">
        <h1>&#x1F510; Token Status</h1>
        <p>${escapeHtml(status)}</p>
      </body>
    </html>`
};

// ─── In-Memory Session Store for CSRF State + PKCE ────────────────────
// BEFORE: State was generated but never stored or validated. Comments said
//         "session management is outside this module's scope" — but that
//         meant state validation was never actually enforced (CSRF risk).
// AFTER: Lightweight in-memory Map with TTL-based auto-expiry.
// GOOD EFFECT: CSRF state validation is fully enforced within this module;
//              PKCE code_verifiers are stored for use during token exchange.

class OAuthSessionStore {
  constructor(ttlMs = 10 * 60 * 1000) { // 10-minute TTL
    this._sessions = new Map();
    this._ttlMs = ttlMs;
  }

  /**
   * Stores a session keyed by state parameter.
   * @param {string} state
   * @param {object} data - { codeVerifier, createdAt }
   */
  set(state, data) {
    data.createdAt = Date.now();
    this._sessions.set(state, data);
    // Auto-cleanup expired sessions periodically
    this._cleanup();
  }

  /**
   * Retrieves and deletes a session (one-time use).
   * @param {string} state
   * @returns {object|null}
   */
  consume(state) {
    const session = this._sessions.get(state);
    if (!session) return null;
    this._sessions.delete(state);
    if (Date.now() - session.createdAt > this._ttlMs) {
      return null; // Expired
    }
    return session;
  }

  _cleanup() {
    const now = Date.now();
    for (const [key, val] of this._sessions) {
      if (now - val.createdAt > this._ttlMs) {
        this._sessions.delete(key);
      }
    }
  }
}

function createAuthConfig(envPrefix = 'MS_') {
  const tenantId = process.env[`${envPrefix}TENANT_ID`] || 'common';
  const authorityHost = (process.env[`${envPrefix}AUTHORITY_HOST`] || 'https://login.microsoftonline.com').replace(/\/+$/, '');

  return {
    clientId: process.env[`${envPrefix}CLIENT_ID`] || '',
    clientSecret: process.env[`${envPrefix}CLIENT_SECRET`] || '',
    redirectUri: process.env[`${envPrefix}REDIRECT_URI`] || 'http://localhost:3333/auth/callback',
    scopes: (process.env[`${envPrefix}SCOPES`] || 'offline_access User.Read Mail.Read').split(' '),
    tenantId,
    tokenEndpoint: process.env[`${envPrefix}TOKEN_ENDPOINT`] || `${authorityHost}/${tenantId}/oauth2/v2.0/token`,
    authEndpoint: process.env[`${envPrefix}AUTH_ENDPOINT`] || `${authorityHost}/${tenantId}/oauth2/v2.0/authorize`
  };
}

function setupOAuthRoutes(app, tokenStorage, authConfig, envPrefix = 'MS_') {
  if (!authConfig) {
    authConfig = createAuthConfig(envPrefix);
  }

  if (!(tokenStorage instanceof TokenStorage)) {
    console.error("Error: tokenStorage is not an instance of TokenStorage. OAuth routes will not function correctly.");
  }

  // BEFORE: No session store — state was generated but never stored/validated.
  // AFTER: Per-setupOAuthRoutes session store for CSRF and PKCE.
  // GOOD EFFECT: Full CSRF protection — callback state is validated against
  //              the stored value; PKCE code_verifier is available for exchange.
  const sessionStore = new OAuthSessionStore();

  app.get('/auth', (req, res) => {
    if (!authConfig.clientId) {
      return res.status(500).send(templates.authError('Configuration Error', 'Client ID is not configured.'));
    }

    // BEFORE: const state = crypto.randomBytes(16).toString('hex');
    //         const authorizationUrl = `${authConfig.authEndpoint}?` + querystring.stringify({...});
    //         — URL built inline, duplicating logic from tools.js; no PKCE.
    // AFTER: const { url, codeVerifier, state } = buildAuthUrl(authConfig);
    // GOOD EFFECT: URL built via shared utility; PKCE enabled; no duplication.
    const { url, codeVerifier, state } = buildAuthUrl(authConfig);

    // BEFORE: State was generated but NOT stored — the callback could never validate it.
    //         Comments acknowledged this: "session management is outside this module's scope".
    // AFTER: State + codeVerifier stored in the in-memory session store.
    // GOOD EFFECT: The /auth/callback route can now fully validate state (CSRF protection)
    //              and pass the code_verifier for PKCE token exchange.
    sessionStore.set(state, { codeVerifier });

    res.redirect(url);
  });

  app.get('/auth/callback', async (req, res) => {
    const { code, error, error_description, state } = req.query;

    // ── State Validation (CSRF Protection) ──
    // BEFORE: State presence was checked but VALUE was never validated.
    //         The code had extensive comments about "consuming app must validate"
    //         but no one actually did. This left a CSRF vulnerability.
    // AFTER: State is validated against the session store — both presence AND value.
    // GOOD EFFECT: Full CSRF protection. An attacker cannot forge a callback
    //              because they don't know the random state value.
    if (!state) {
      console.error("OAuth callback received without 'state' parameter.");
      return res.status(400).send(templates.authError('Missing State Parameter',
        'The state parameter was missing from the OAuth callback. Please try again.'));
    }

    const session = sessionStore.consume(state);
    if (!session) {
      // BEFORE: This check was commented out — state mismatch was silently ignored.
      // AFTER: Enforced — unknown or expired state is rejected.
      // GOOD EFFECT: Prevents CSRF attacks and replay of old callback URLs.
      console.error("OAuth callback state mismatch or expired. Potential CSRF attack.");
      return res.status(400).send(templates.authError('Invalid State',
        'CSRF token mismatch or expired. Please try authenticating again.'));
    }

    if (error) {
      return res.status(400).send(templates.authError(error, error_description));
    }

    if (!code) {
      return res.status(400).send(templates.authError('Missing Authorization Code',
        'No authorization code was provided in the callback.'));
    }

    try {
      // BEFORE: await tokenStorage.exchangeCodeForTokens(code);
      //         — no PKCE code_verifier passed.
      // AFTER: Pass the code_verifier from the session for PKCE verification.
      // GOOD EFFECT: Microsoft's token endpoint verifies the code_challenge
      //              matches the code_verifier, preventing code interception attacks.
      await tokenStorage.exchangeCodeForTokens(code, session.codeVerifier);
      res.send(templates.authSuccess);
    } catch (exchangeError) {
      console.error('Token exchange error:', exchangeError);
      res.status(500).send(templates.tokenExchangeError(exchangeError));
    }
  });

  app.get('/token-status', async (req, res) => {
    try {
      const token = await tokenStorage.getValidAccessToken();
      if (token) {
        const expiryDate = new Date(tokenStorage.getExpiryTime());
        res.send(templates.tokenStatus(`Access token is valid. Expires at: ${expiryDate.toLocaleString()}`));
      } else {
        res.send(templates.tokenStatus('No valid access token found. Please authenticate.'));
      }
    } catch (err) {
      res.status(500).send(templates.tokenStatus(`Error checking token status: ${err.message}`));
    }
  });
}

module.exports = {
  setupOAuthRoutes,
  createAuthConfig,
  // BEFORE: OAuthSessionStore didn't exist.
  // AFTER: Exported for testing.
  // GOOD EFFECT: Tests can verify session lifecycle without spinning up Express.
  OAuthSessionStore
};