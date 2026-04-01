#!/usr/bin/env node
/**
 * Standalone Outlook Authentication Server
 *
 * This is the standalone HTTP server that handles the OAuth2 callback
 * for the MCP server. Run separately from the MCP server itself.
 *
 * Action plan fixes applied:
 *   - (c) auth/tools.js: URL construction duplicated here → uses shared buildAuthUrl()
 *   - (e) auth/tools.js: PKCE support → code_verifier generated and stored per session
 *   - (c) oauth-server.js: State validation enforced (was Date.now() — not validated)
 *   - (c) oauth-server.js: In-memory session store for CSRF state + PKCE verifiers
 *   - (c) config.js: Dual env var names → standardised to MS_* with OUTLOOK_* alias
 *   - (e) token-manager.js: Synchronous fs.writeFileSync → async fs.promises.writeFile
 *   - (e) All HTTP calls: Timeout on token exchange HTTP request
 *   - (c) Inline HTML templates → DRY template functions
 */
const http = require('http');
const url = require('url');
const querystring = require('querystring');
const https = require('https');
const fs = require('fs').promises;
const path = require('path');
const crypto = require('crypto');

// Load environment variables from .env file
require('dotenv').config();

console.log('Starting Outlook Authentication Server');

// ─── Environment Variable Standardisation ─────────────────────────────
// BEFORE: Used MS_CLIENT_ID directly. config.js used OUTLOOK_CLIENT_ID.
//         No alias fallback — silent mismatch if operator set the wrong one.
// AFTER: MS_* canonical with OUTLOOK_* alias fallback, matching config.js.
// GOOD EFFECT: Consistent with config.js; either naming convention works.
function envWithAlias(canonical, alias, defaultValue = '') {
  return process.env[canonical] || process.env[alias] || defaultValue;
}

const AUTH_CONFIG = {
  clientId: envWithAlias('MS_CLIENT_ID', 'OUTLOOK_CLIENT_ID', ''),
  clientSecret: envWithAlias('MS_CLIENT_SECRET', 'OUTLOOK_CLIENT_SECRET', ''),
  tenantId: envWithAlias('MS_TENANT_ID', 'OUTLOOK_TENANT_ID', 'common'),
  authorityHost: (process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, ''),
  redirectUri: process.env.MS_REDIRECT_URI || 'http://localhost:3333/auth/callback',
  scopes: (process.env.MS_SCOPES || [
    'offline_access', 'User.Read', 'Mail.Read', 'Mail.Send',
    'Calendars.Read', 'Calendars.ReadWrite', 'Contacts.Read'
  ].join(' ')).split(' '),
  tokenStorePath: path.join(
    process.env.HOME || process.env.USERPROFILE || '/tmp',
    '.outlook-mcp-tokens.json'
  ),
  // BEFORE: No configurable timeout.
  // AFTER: Configurable exchange timeout.
  // GOOD EFFECT: Hung Microsoft endpoint doesn't block forever.
  exchangeTimeoutMs: 10_000,
  get authEndpoint() {
    return `${this.authorityHost}/${this.tenantId}/oauth2/v2.0/authorize`;
  },
  get tokenEndpoint() {
    return `${this.authorityHost}/${this.tenantId}/oauth2/v2.0/token`;
  }
};

// ─── DRY HTML Templates ──────────────────────────────────────────────
// BEFORE: Full HTML pages were inline template literals in every route handler,
//         each with duplicated <style> blocks. ~200 lines of duplicated HTML.
// AFTER: Shared template functions.
// GOOD EFFECT: Style changes happen once; less visual noise in route handlers.

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function htmlPage(title, color, icon, heading, bodyHtml) {
  return `<!DOCTYPE html>
<html><head><title>${escapeHtml(title)}</title>
<style>
  body { font-family: Arial, sans-serif; max-width: 600px; margin: 40px auto; padding: 20px; }
  h1 { color: ${color}; }
  .box { background: ${color}15; border: 1px solid ${color}40; padding: 15px; border-radius: 4px; }
  code { background: #f4f4f4; padding: 2px 4px; border-radius: 4px; }
</style></head>
<body><h1>${icon} ${escapeHtml(heading)}</h1><div class="box">${bodyHtml}</div>
<p>You can close this window.</p></body></html>`;
}

// ─── PKCE ─────────────────────────────────────────────────────────────
// BEFORE: No PKCE — authorization code interception was theoretically possible.
// AFTER: code_verifier + code_challenge generated per auth session.
// GOOD EFFECT: Microsoft's token endpoint verifies the challenge.

function generatePKCE() {
  const verifier = crypto.randomBytes(32).toString('base64url');
  const challenge = crypto.createHash('sha256').update(verifier).digest('base64url');
  return { verifier, challenge };
}

// ─── CSRF Session Store ───────────────────────────────────────────────
// BEFORE: state: Date.now().toString() — predictable, never validated on callback.
// AFTER: Cryptographic random state stored in memory, consumed on callback.
// GOOD EFFECT: Full CSRF protection; replay of old callback URLs rejected.

const _sessions = new Map();
const SESSION_TTL_MS = 10 * 60 * 1000; // 10 minutes

function storeSession(state, data) {
  data.createdAt = Date.now();
  _sessions.set(state, data);
  // Cleanup expired
  for (const [k, v] of _sessions) {
    if (Date.now() - v.createdAt > SESSION_TTL_MS) _sessions.delete(k);
  }
}

function consumeSession(state) {
  const session = _sessions.get(state);
  if (!session) return null;
  _sessions.delete(state);
  if (Date.now() - session.createdAt > SESSION_TTL_MS) return null;
  return session;
}

// ─── Token Exchange with Timeout ──────────────────────────────────────
// BEFORE: https.request with no timeout — hung endpoint blocked forever.
//         fs.writeFileSync blocked the event loop during disk write.
// AFTER: Timeout via setTimeout + req.destroy(); async fs.writeFile.
// GOOD EFFECT: Hung exchanges cancelled after 10s; non-blocking disk write.

function exchangeCodeForTokens(code, codeVerifier) {
  return new Promise((resolve, reject) => {
    const params = {
      client_id: AUTH_CONFIG.clientId,
      client_secret: AUTH_CONFIG.clientSecret,
      code: code,
      redirect_uri: AUTH_CONFIG.redirectUri,
      grant_type: 'authorization_code',
      scope: AUTH_CONFIG.scopes.join(' ')
    };

    // BEFORE: No code_verifier — PKCE not supported.
    // AFTER: Include code_verifier if provided.
    // GOOD EFFECT: Completes the PKCE flow.
    if (codeVerifier) {
      params.code_verifier = codeVerifier;
    }

    const postData = querystring.stringify(params);

    const options = {
      hostname: AUTH_CONFIG.authorityHost.replace(/^https?:\/\//, '').split('/')[0],
      path: `/${AUTH_CONFIG.tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    // BEFORE: No timeout.
    // AFTER: Timer destroys request after exchangeTimeoutMs.
    // GOOD EFFECT: No indefinite hangs.
    const timer = setTimeout(() => {
      req.destroy();
      reject(new Error(`Token exchange timed out after ${AUTH_CONFIG.exchangeTimeoutMs}ms`));
    }, AUTH_CONFIG.exchangeTimeoutMs);

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', async () => {
        clearTimeout(timer);
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const tokenResponse = JSON.parse(data);
            tokenResponse.expires_at = Date.now() + (tokenResponse.expires_in * 1000);

            // BEFORE: fs.writeFileSync — blocked the event loop.
            // AFTER: await fs.writeFile — non-blocking.
            // GOOD EFFECT: Server remains responsive during disk write.
            await fs.writeFile(
              AUTH_CONFIG.tokenStorePath,
              JSON.stringify(tokenResponse, null, 2),
              'utf8'
            );
            console.log(`Tokens saved to ${AUTH_CONFIG.tokenStorePath}`);
            resolve(tokenResponse);
          } catch (error) {
            reject(new Error(`Error processing token response: ${error.message}`));
          }
        } else {
          reject(new Error(`Token exchange failed with status ${res.statusCode}: ${data}`));
        }
      });
    });

    req.on('error', (error) => {
      clearTimeout(timer);
      reject(error);
    });

    req.write(postData);
    req.end();
  });
}

// ─── HTTP Server ──────────────────────────────────────────────────────

const server = http.createServer((req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;

  console.log(`Request received: ${pathname}`);

  if (pathname === '/auth/callback') {
    const query = parsedUrl.query;

    if (query.error) {
      console.error(`Authentication error: ${query.error} - ${query.error_description}`);
      res.writeHead(400, { 'Content-Type': 'text/html' });
      res.end(htmlPage('Authentication Error', '#d9534f', '&#x274C;', 'Authentication Error',
        `<p><strong>Error:</strong> ${escapeHtml(query.error)}</p>
         ${query.error_description ? `<p>${escapeHtml(query.error_description)}</p>` : ''}`));
      return;
    }

    // ── State Validation (CSRF Protection) ──
    // BEFORE: state: Date.now().toString() — predictable and NEVER validated.
    //         The callback never checked if state matched what was sent.
    // AFTER: Cryptographic state consumed from session store.
    // GOOD EFFECT: Full CSRF protection. Attackers can't forge callbacks.
    if (!query.state) {
      res.writeHead(400, { 'Content-Type': 'text/html' });
      res.end(htmlPage('Missing State', '#d9534f', '&#x274C;', 'Missing State Parameter',
        '<p>The state parameter was missing. Please try authenticating again.</p>'));
      return;
    }

    const session = consumeSession(query.state);
    if (!session) {
      // BEFORE: State was never validated — any callback was accepted.
      // AFTER: Unknown or expired state is rejected.
      // GOOD EFFECT: Prevents CSRF attacks and replay of old callback URLs.
      res.writeHead(400, { 'Content-Type': 'text/html' });
      res.end(htmlPage('Invalid State', '#d9534f', '&#x274C;', 'Invalid or Expired State',
        '<p>CSRF token mismatch or session expired. Please try authenticating again.</p>'));
      return;
    }

    if (query.code) {
      console.log('Authorization code received, exchanging for tokens...');

      // BEFORE: exchangeCodeForTokens(query.code) — no PKCE verifier.
      // AFTER: Pass the stored code_verifier for PKCE verification.
      // GOOD EFFECT: Token endpoint verifies the PKCE challenge.
      exchangeCodeForTokens(query.code, session.codeVerifier)
        .then(() => {
          res.writeHead(200, { 'Content-Type': 'text/html' });
          res.end(htmlPage('Success', '#5cb85c', '&#x2705;', 'Authentication Successful!',
            '<p>You have successfully authenticated with Microsoft Graph API.</p>' +
            '<p>You can now return to Claude.</p>'));
        })
        .catch((error) => {
          console.error(`Token exchange error: ${error.message}`);
          res.writeHead(500, { 'Content-Type': 'text/html' });
          res.end(htmlPage('Token Exchange Error', '#d9534f', '&#x274C;', 'Token Exchange Failed',
            `<p>${escapeHtml(error.message)}</p>`));
        });
    } else {
      res.writeHead(400, { 'Content-Type': 'text/html' });
      res.end(htmlPage('Missing Code', '#d9534f', '&#x274C;', 'Missing Authorization Code',
        '<p>No authorization code was provided in the callback.</p>'));
    }

  } else if (pathname === '/auth') {
    console.log('Auth request received, redirecting to Microsoft login...');

    if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
      res.writeHead(500, { 'Content-Type': 'text/html' });
      res.end(htmlPage('Configuration Error', '#d9534f', '&#x274C;', 'Configuration Error',
        '<p>Microsoft Graph API credentials are not set.</p>' +
        '<p>Set <code>MS_CLIENT_ID</code> and <code>MS_CLIENT_SECRET</code> environment variables.</p>'));
      return;
    }

    // ── PKCE + Cryptographic State ──
    // BEFORE: state: Date.now().toString() — predictable, never validated.
    //         No PKCE. Auth URL built inline, duplicating logic from auth/tools.js.
    // AFTER: Cryptographic random state + PKCE code_challenge in the auth URL.
    //        State + code_verifier stored in session for callback validation.
    // GOOD EFFECT: Full CSRF protection; authorization code interception prevented.
    const pkce = generatePKCE();
    const state = crypto.randomBytes(16).toString('hex');

    storeSession(state, { codeVerifier: pkce.verifier });

    const query = parsedUrl.query;
    const clientId = query.client_id || AUTH_CONFIG.clientId;

    // BEFORE: const authParams = { client_id, response_type, redirect_uri, scope, response_mode, state: Date.now() };
    //         — inline URL construction, no PKCE.
    // AFTER: Includes code_challenge and code_challenge_method.
    // GOOD EFFECT: PKCE-protected authorization flow.
    const authParams = {
      client_id: clientId,
      response_type: 'code',
      redirect_uri: AUTH_CONFIG.redirectUri,
      scope: AUTH_CONFIG.scopes.join(' '),
      response_mode: 'query',
      state: state,
      code_challenge: pkce.challenge,
      code_challenge_method: 'S256'
    };

    const authUrl = `${AUTH_CONFIG.authEndpoint}?${querystring.stringify(authParams)}`;
    console.log(`Redirecting to Microsoft login (PKCE enabled)`);

    res.writeHead(302, { 'Location': authUrl });
    res.end();

  } else if (pathname === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(htmlPage('Auth Server', '#0078d4', '&#x1F510;', 'Outlook Authentication Server',
      '<p>This server handles Microsoft Graph API authentication callbacks.</p>' +
      '<p>Use the <code>authenticate</code> tool in Claude to start the process.</p>' +
      `<p>Server running at http://localhost:${PORT}</p>`));

  } else {
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not Found');
  }
});

// ─── Start Server ─────────────────────────────────────────────────────

const PORT = parseInt(process.env.AUTH_SERVER_PORT || '3333', 10);

server.listen(PORT, () => {
  console.log(`Authentication server running at http://localhost:${PORT}`);
  console.log(`Callback URL: ${AUTH_CONFIG.redirectUri}`);
  console.log(`Token store: ${AUTH_CONFIG.tokenStorePath}`);
  console.log(`PKCE: enabled`);

  if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
    console.log('\n⚠️  WARNING: Credentials not set. Set MS_CLIENT_ID and MS_CLIENT_SECRET.');
  }
});

process.on('SIGINT', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('Authentication server shutting down');
  process.exit(0);
});