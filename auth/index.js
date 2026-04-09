/**
 * Authentication module for Outlook MCP server.
 *
 * Canva-style resolution order:
 * 1) explicit bearer_token arg
 * 2) per-request context (AsyncLocalStorage)
 * 3) Authorization header from raw HTTP request (if available)
 * 4) MS_ACCESS_TOKEN env var
 * 5) stored OAuth token manager
 */
const { AsyncLocalStorage } = require('async_hooks');
const TokenManager = require('./token-manager');
const { authTools } = require('./tools');
const config = require('../config');

const bearerTokenStorage = new AsyncLocalStorage();
const _defaultTokenManager = new TokenManager(config);

class OutlookClient {
  constructor(token) {
    if (!token || typeof token !== 'string') {
      throw new Error('OutlookClient requires a non-empty token string.');
    }
    this._token = token;
  }

  authHeaders(contentType = null) {
    const headers = { Authorization: `Bearer ${this._token}` };
    if (contentType) headers['Content-Type'] = contentType;
    return headers;
  }

  get rawToken() {
    return this._token;
  }
}

function createTokenManager(cfg) {
  return new TokenManager(cfg);
}

async function resolveToken(bearerToken = null, httpRequest = null, forceNew = false, tokenManagerInstance = null) {
  if (bearerToken && typeof bearerToken === 'string' && bearerToken.trim()) {
    return bearerToken.trim();
  }

  const storedToken = bearerTokenStorage.getStore();
  if (storedToken && typeof storedToken === 'string' && storedToken.trim()) {
    return storedToken.trim();
  }

  if (httpRequest) {
    const raw = (httpRequest.headers?.authorization || httpRequest.headers?.Authorization || '').trim();
    if (raw.startsWith('Bearer ')) {
      const token = raw.slice(7).trim();
      if (token) return token;
    }
  }

  const envToken = (process.env.MS_ACCESS_TOKEN || '').trim();
  if (envToken) return envToken;

  if (!forceNew) {
    const tm = tokenManagerInstance || _defaultTokenManager;
    const accessToken = await tm.getValidAccessToken();
    if (accessToken && accessToken.trim()) return accessToken.trim();
  }

  throw new Error('Authentication required. No valid token found from header, token manager, or env var.');
}

async function getClient(bearerToken = null, httpRequest = null, forceNew = false, tokenManagerInstance = null) {
  const token = await resolveToken(bearerToken, httpRequest, forceNew, tokenManagerInstance);
  return new OutlookClient(token);
}

// Backward-compatible shim used by existing handlers.
async function ensureAuthenticated(tokenManagerInstance = null, forceNew = false, bearerToken = null, httpRequest = null) {
  return resolveToken(bearerToken, httpRequest, forceNew, tokenManagerInstance);
}

module.exports = {
  bearerTokenStorage,
  OutlookClient,
  getClient,
  resolveToken,
  createTokenManager,
  authTools,
  ensureAuthenticated
};