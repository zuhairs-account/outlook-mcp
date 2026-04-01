/**
 * Token management for Microsoft Graph API authentication
 *
 * Addresses action plan items for auth/token-manager.js:
 *   - P0: Refresh race condition → module-level Promise lock
 *   - P1: Synchronous fs I/O → replaced with fs.promises
 *   - P2: No in-memory token cache → cache with expiry; disk read only on miss
 *   - P0: No timeout on refresh HTTP call → AbortController-based timeout
 *   - (c): God Method decomposed into loadTokens/isExpired/refreshTokens/persistTokens
 *   - (c): File path injected via config, not hardcoded
 *   - (c): Dual token management separated (Flow tokens handled distinctly)
 */

// BEFORE: const fs = require('fs');
// AFTER: const fs = require('fs').promises;
// GOOD EFFECT: All file I/O is now async, preventing event-loop blocking
//              that stalls other MCP requests during disk reads/writes.
const fs = require('fs').promises;
const fsSync = require('fs');
const path = require('path');
const https = require('https');
const querystring = require('querystring');

// BEFORE: const config = require('../config');
//         — hardcoded dependency on a relative config file
// AFTER: Config is injected via constructor parameter.
// GOOD EFFECT: Tests can pass mock config (temp file paths, fake credentials)
//              without monkey-patching require().

// BEFORE: let cachedTokens = null;  (module-level mutable singleton)
// AFTER: Instance-level cache inside the class.
// GOOD EFFECT: Multiple TokenManager instances (Graph vs Flow) don't
//              share a single mutable cache — eliminates cross-contamination.

/**
 * @class TokenManager
 * Manages OAuth2 token lifecycle: loading, caching, expiry detection,
 * refresh, and persistence.
 *
 * BEFORE: Module-level functions with a global `cachedTokens` variable.
 * AFTER: Encapsulated class with injected config and instance-level state.
 * GOOD EFFECT: Dependency Inversion — callers inject config/storage path;
 *              testable without touching the filesystem; no global mutable state.
 */
class TokenManager {
  /**
   * @param {object} config - Configuration object
   * @param {string} config.tokenStorePath - Path to the token JSON file
   * @param {string} config.clientId - OAuth client ID
   * @param {string} config.clientSecret - OAuth client secret
   * @param {string} config.tokenEndpoint - Microsoft token endpoint URL
   * @param {string[]} config.scopes - OAuth scopes
   * @param {number} [config.refreshTimeoutMs=10000] - Timeout for refresh HTTP call
   * @param {number} [config.refreshBufferMs=300000] - Buffer before expiry to trigger refresh (5 min)
   */
  constructor(config = {}) {
    const tenantId = config.tenantId || process.env.MS_TENANT_ID || 'common';
    const authorityHost = (config.authorityHost || process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, '');

    // BEFORE: File path was hardcoded inside loadTokenCache() as config.AUTH_CONFIG.tokenStorePath
    // AFTER: Injected and stored as an instance property.
    // GOOD EFFECT: Tests can point to temp files; no implicit coupling to a global config module.
    this.tokenStorePath = config.tokenStorePath ||
      path.join(process.env.HOME || process.env.USERPROFILE || '/tmp', '.outlook-mcp-tokens.json');
    this.clientId = config.clientId || process.env.MS_CLIENT_ID || '';
    this.clientSecret = config.clientSecret || process.env.MS_CLIENT_SECRET || '';
    this.redirectUri = config.redirectUri || process.env.MS_REDIRECT_URI || 'http://localhost:3333/auth/callback';
    this.scopes = config.scopes || (process.env.MS_SCOPES || 'offline_access User.Read Mail.Read').split(' ');
    this.tokenEndpoint = config.tokenEndpoint || `${authorityHost}/${tenantId}/oauth2/v2.0/token`;

    // BEFORE: No timeout on refresh HTTP call (P0 issue).
    // AFTER: Configurable timeout, defaults to 10 seconds.
    // GOOD EFFECT: A hung Microsoft token endpoint no longer blocks the
    //              MCP server indefinitely.
    this.refreshTimeoutMs = config.refreshTimeoutMs || 10_000;

    // BEFORE: No buffer — tokens could be used right up to expiry.
    // AFTER: 5-minute buffer before expiry triggers proactive refresh.
    // GOOD EFFECT: Prevents "token expired mid-request" failures.
    this.refreshBufferMs = config.refreshBufferMs || 5 * 60 * 1000;

    // BEFORE: let cachedTokens = null; (module-level global)
    // AFTER: Instance-level cache.
    // GOOD EFFECT: No cross-contamination between Graph and Flow token managers.
    this._cachedTokens = null;

    // BEFORE: No refresh lock — concurrent refreshes caused race conditions (P0).
    // AFTER: Module-level (now instance-level) Promise lock.
    // GOOD EFFECT: If two tool calls arrive simultaneously with an expired token,
    //              only one refresh fires; the second awaits the first's result.
    this._refreshPromise = null;
  }

  // ─── Decomposed Methods ─────────────────────────────────────────────
  // BEFORE: Everything was in one God Method `getAccessToken()` which
  //         read disk, parsed JSON, checked expiry, called Microsoft,
  //         wrote disk, and returned a token — all inline.
  // AFTER: Decomposed into loadTokens(), isExpired(), refreshTokens(),
  //        persistTokens(), each independently testable.
  // GOOD EFFECT: Single Responsibility per method; unit tests can verify
  //              each step in isolation.

  /**
   * Loads tokens from disk (async).
   *
   * BEFORE: fs.readFileSync — blocked the event loop (P1 issue).
   * AFTER: await fs.readFile — non-blocking.
   * GOOD EFFECT: Other MCP requests are processed while disk I/O completes.
   *
   * @returns {Promise<object|null>} Parsed token object or null
   */
  async loadTokens() {
    try {
      // BEFORE: const tokenData = fs.readFileSync(tokenPath, 'utf8');
      // AFTER:  const tokenData = await fs.readFile(this.tokenStorePath, 'utf8');
      // GOOD EFFECT: Non-blocking I/O — the event loop remains free for
      //              concurrent MCP requests.
      const tokenData = await fs.readFile(this.tokenStorePath, 'utf8');
      const tokens = JSON.parse(tokenData);

      if (!tokens.access_token) {
        console.error('[TokenManager] No access_token found in stored tokens');
        return null;
      }

      this._cachedTokens = tokens;
      return tokens;
    } catch (error) {
      if (error.code === 'ENOENT') {
        console.log('[TokenManager] Token file not found. No tokens loaded.');
      } else {
        console.error('[TokenManager] Error loading token cache:', error.message);
      }
      this._cachedTokens = null;
      return null;
    }
  }

  /**
   * Checks if the current cached token is expired (or near-expired).
   *
   * BEFORE: Expiry check was inline inside the God Method.
   * AFTER: Standalone pure method.
   * GOOD EFFECT: Independently testable; refresh buffer logic is explicit.
   *
   * @returns {boolean}
   */
  isExpired() {
    if (!this._cachedTokens || !this._cachedTokens.expires_at) {
      return true;
    }
    // BEFORE: if (now > expiresAt) — no buffer, tokens used right up to expiry.
    // AFTER: Subtract refreshBufferMs to trigger proactive refresh.
    // GOOD EFFECT: Prevents "token expired mid-Graph-API-call" failures.
    return Date.now() >= (this._cachedTokens.expires_at - this.refreshBufferMs);
  }

  /**
   * Persists tokens to disk (async).
   *
   * BEFORE: fs.writeFileSync(tokenPath, ...) — blocked the event loop.
   * AFTER: await fs.writeFile — non-blocking.
   * GOOD EFFECT: Disk writes no longer stall concurrent MCP request processing.
   *
   * @param {object} tokens
   * @returns {Promise<void>}
   */
  async persistTokens(tokens) {
    try {
      // BEFORE: fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
      // AFTER:  await fs.writeFile(...)
      // GOOD EFFECT: Non-blocking disk write.
      await fs.writeFile(this.tokenStorePath, JSON.stringify(tokens, null, 2));
      this._cachedTokens = tokens;
      console.log('[TokenManager] Tokens persisted successfully.');
    } catch (error) {
      console.error('[TokenManager] Error persisting tokens:', error.message);
      throw error;
    }
  }

  /**
   * Refreshes the access token using the refresh_token grant.
   *
   * BEFORE: No race protection — concurrent callers each fired a separate
   *         refresh HTTP call, causing race conditions on disk writes (P0).
   * AFTER: Instance-level _refreshPromise lock — second caller awaits first.
   * GOOD EFFECT: Only one refresh fires; no duplicate token endpoint calls;
   *              no race on writing the token file.
   *
   * BEFORE: No timeout — a hung Microsoft endpoint blocked forever (P0).
   * AFTER: AbortController-based timeout (default 10s).
   * GOOD EFFECT: Stalled refresh calls are cancelled after the timeout,
   *              and the caller receives a clear timeout error.
   *
   * @returns {Promise<string>} Fresh access token
   */
  async refreshTokens() {
    // ── Refresh lock (P0 fix) ──
    // BEFORE: (no lock) — every concurrent caller hit the Microsoft endpoint.
    // AFTER: If a refresh is already in flight, queue this caller to await it.
    // GOOD EFFECT: Eliminates duplicate refresh calls and disk-write races.
    if (this._refreshPromise) {
      console.log('[TokenManager] Refresh already in progress, awaiting existing promise.');
      return this._refreshPromise;
    }

    if (!this._cachedTokens || !this._cachedTokens.refresh_token) {
      throw new Error('No refresh token available.');
    }

    this._refreshPromise = this._doRefresh();

    try {
      const accessToken = await this._refreshPromise;
      return accessToken;
    } finally {
      // BEFORE: _refreshPromise was never cleared (or cleared inconsistently).
      // AFTER: Always cleared in finally block.
      // GOOD EFFECT: Subsequent refresh attempts aren't permanently blocked.
      this._refreshPromise = null;
    }
  }

  /**
   * Internal: performs the actual HTTP refresh call.
   * Separated from refreshTokens() so the lock logic stays clean.
   * @private
   */
  async _doRefresh() {
    console.log('[TokenManager] Refreshing access token...');
    const postData = querystring.stringify({
      client_id: this.clientId,
      client_secret: this.clientSecret,
      grant_type: 'refresh_token',
      refresh_token: this._cachedTokens.refresh_token,
      scope: this.scopes.join(' ')
    });

    // BEFORE: No timeout on the HTTPS request (P0).
    // AFTER: Promise.race with a rejection timer.
    // GOOD EFFECT: Hung Microsoft endpoint is cancelled after refreshTimeoutMs.
    const timeoutMs = this.refreshTimeoutMs;

    return new Promise((resolve, reject) => {
      // ── Timeout guard ──
      const timer = setTimeout(() => {
        req.destroy();
        reject(new Error(`Token refresh timed out after ${timeoutMs}ms`));
      }, timeoutMs);

      const requestOptions = {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(postData)
        }
      };

      const req = https.request(this.tokenEndpoint, requestOptions, (res) => {
        let data = '';
        res.on('data', (chunk) => data += chunk);
        res.on('end', async () => {
          clearTimeout(timer);
          try {
            const body = JSON.parse(data);
            if (res.statusCode >= 200 && res.statusCode < 300) {
              this._cachedTokens.access_token = body.access_token;
              if (body.refresh_token) {
                this._cachedTokens.refresh_token = body.refresh_token;
              }
              this._cachedTokens.expires_in = body.expires_in;
              this._cachedTokens.expires_at = Date.now() + (body.expires_in * 1000);

              await this.persistTokens(this._cachedTokens);
              console.log('[TokenManager] Access token refreshed successfully.');
              resolve(this._cachedTokens.access_token);
            } else {
              reject(new Error(body.error_description || `Token refresh failed with status ${res.statusCode}`));
            }
          } catch (e) {
            reject(e);
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

  /**
   * Returns a valid access token, refreshing if necessary.
   *
   * BEFORE: getAccessToken() — synchronous, read disk every time, no refresh.
   * AFTER: getValidAccessToken() — async, uses in-memory cache with TTL,
   *        auto-refreshes on expiry.
   * GOOD EFFECT: (P2) Disk is only read on cache miss; (P0) expired tokens
   *              are auto-refreshed; callers always get a usable token.
   *
   * @returns {Promise<string|null>}
   */
  async getValidAccessToken() {
    // ── In-memory cache check (P2 fix) ──
    // BEFORE: Every call to getAccessToken() hit the disk via loadTokenCache().
    // AFTER: Return cached token immediately if still valid.
    // GOOD EFFECT: Eliminates redundant disk reads for every Graph API call.
    if (this._cachedTokens && this._cachedTokens.access_token && !this.isExpired()) {
      return this._cachedTokens.access_token;
    }

    // Cache miss or expired — load from disk
    if (!this._cachedTokens) {
      await this.loadTokens();
    }

    if (!this._cachedTokens || !this._cachedTokens.access_token) {
      return null;
    }

    if (this.isExpired()) {
      if (this._cachedTokens.refresh_token) {
        try {
          return await this.refreshTokens();
        } catch (err) {
          console.error('[TokenManager] Refresh failed:', err.message);
          this._cachedTokens = null;
          return null;
        }
      } else {
        console.warn('[TokenManager] Token expired, no refresh token available.');
        this._cachedTokens = null;
        return null;
      }
    }

    return this._cachedTokens.access_token;
  }

  // ─── Legacy compat: synchronous getAccessToken (deprecated) ─────────
  // BEFORE: function getAccessToken() — the only public method.
  // AFTER: Kept for backward compat but marked deprecated; delegates to cache.
  // GOOD EFFECT: Existing callers don't break immediately, but are guided
  //              toward the async getValidAccessToken().
  /**
   * @deprecated Use getValidAccessToken() instead.
   * @returns {string|null}
   */
  getAccessToken() {
    if (this._cachedTokens && this._cachedTokens.access_token && !this.isExpired()) {
      return this._cachedTokens.access_token;
    }
    return null;
  }

  /**
   * Gets the token expiry timestamp.
   * @returns {number}
   */
  getExpiryTime() {
    return this._cachedTokens && this._cachedTokens.expires_at ? this._cachedTokens.expires_at : 0;
  }

  // ─── Flow Token Management (SRP separation) ────────────────────────
  // BEFORE: getFlowAccessToken() and saveFlowTokens() were in the same
  //         module as Graph token management — SRP violation.
  // AFTER: Still present for backward compat, but clearly separated
  //        into their own section. The action plan recommends eventually
  //        splitting into GraphTokenManager and FlowTokenManager sharing
  //        a base OAuthTokenManager.
  // GOOD EFFECT: Clear boundary between Graph and Flow token concerns;
  //              easier to extract into separate classes later.

  /**
   * Gets the current Flow API access token (if stored alongside Graph tokens).
   * @returns {string|null}
   */
  getFlowAccessToken() {
    if (!this._cachedTokens) return null;
    if (this._cachedTokens.flow_access_token && this._cachedTokens.flow_expires_at) {
      if (Date.now() < this._cachedTokens.flow_expires_at) {
        return this._cachedTokens.flow_access_token;
      }
    }
    return null;
  }

  /**
   * Saves Flow API tokens alongside existing Graph tokens.
   * @param {object} flowTokens
   * @returns {Promise<boolean>}
   */
  async saveFlowTokens(flowTokens) {
    try {
      // BEFORE: Synchronous fs.readFileSync + fs.writeFileSync
      // AFTER: Async fs.readFile + fs.writeFile
      // GOOD EFFECT: Non-blocking I/O.
      let existingTokens = {};
      try {
        const data = await fs.readFile(this.tokenStorePath, 'utf8');
        existingTokens = JSON.parse(data);
      } catch (e) {
        // File doesn't exist yet — start fresh
      }

      const mergedTokens = {
        ...existingTokens,
        flow_access_token: flowTokens.access_token,
        flow_refresh_token: flowTokens.refresh_token,
        flow_expires_at: flowTokens.expires_at || (Date.now() + (flowTokens.expires_in || 3600) * 1000)
      };

      await this.persistTokens(mergedTokens);
      console.log('[TokenManager] Flow tokens saved successfully.');
      return true;
    } catch (error) {
      console.error('[TokenManager] Error saving Flow tokens:', error.message);
      return false;
    }
  }

  /**
   * Creates test tokens for USE_TEST_MODE.
   * @returns {Promise<object>}
   */
  async createTestTokens() {
    const testTokens = {
      access_token: "test_access_token_" + Date.now(),
      refresh_token: "test_refresh_token_" + Date.now(),
      expires_at: Date.now() + (3600 * 1000)
    };
    await this.persistTokens(testTokens);
    return testTokens;
  }

  /**
   * Clears all tokens (for logout or forced re-auth).
   * @returns {Promise<void>}
   */
  async clearTokens() {
    this._cachedTokens = null;
    try {
      await fs.unlink(this.tokenStorePath);
      console.log('[TokenManager] Token file deleted.');
    } catch (error) {
      if (error.code !== 'ENOENT') {
        console.error('[TokenManager] Error deleting token file:', error.message);
      }
    }
  }
}

module.exports = TokenManager;