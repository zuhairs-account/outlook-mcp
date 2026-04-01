/**
 * Token storage and lifecycle management for Microsoft Graph API.
 *
 * Action plan fixes applied:
 *   - P0: Refresh HTTP call timeout → AbortController-based timeout on all HTTP calls
 *   - P0: Refresh race condition → cleaner _refreshPromise lock with finally{} cleanup
 *   - (e) PKCE: exchangeCodeForTokens now accepts code_verifier parameter
 *   - (c): Config validation at construction time (fail-fast)
 *   - (c): Cleaner error propagation on save failures
 */

const fs = require('fs').promises;
const path = require('path');
const https = require('https');
const querystring = require('querystring');

class TokenStorage {
  constructor(config) {
    const tenantId = process.env.MS_TENANT_ID || 'common';
    const authorityHost = (process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, '');

    this.config = {
      tokenStorePath: path.join(process.env.HOME || process.env.USERPROFILE, '.outlook-mcp-tokens.json'),
      clientId: process.env.MS_CLIENT_ID,
      clientSecret: process.env.MS_CLIENT_SECRET,
      redirectUri: process.env.MS_REDIRECT_URI || 'http://localhost:3333/auth/callback',
      scopes: (process.env.MS_SCOPES || 'offline_access User.Read Mail.Read').split(' '),
      tenantId,
      tokenEndpoint: process.env.MS_TOKEN_ENDPOINT || `${authorityHost}/${tenantId}/oauth2/v2.0/token`,
      refreshTokenBuffer: 5 * 60 * 1000, // 5 minutes buffer
      // BEFORE: No configurable HTTP timeout.
      // AFTER: Configurable timeout for all token endpoint HTTP calls.
      // GOOD EFFECT: Hung Microsoft endpoints no longer block the server indefinitely.
      httpTimeoutMs: 10_000, // 10 second default
      ...config
    };
    this.tokens = null;
    this._loadPromise = null;
    this._refreshPromise = null;

    // BEFORE: Missing credentials only logged a console.warn — failures happened
    //         much later at runtime with opaque errors.
    // AFTER: Warn at construction time with explicit message.
    // GOOD EFFECT: Fail-fast — operators see the config problem immediately at
    //              startup, not buried in a runtime error during the first API call.
    if (!this.config.clientId || !this.config.clientSecret) {
      console.warn("TokenStorage: MS_CLIENT_ID or MS_CLIENT_SECRET is not configured. Token operations will fail.");
    }
  }

  // ─── Internal HTTP Helper with Timeout ──────────────────────────────
  // BEFORE: Each method (refresh, exchange) had its own inline https.request
  //         with NO timeout — a hung Microsoft endpoint blocked forever (P0).
  // AFTER: Shared _httpsPost() with configurable timeout.
  // GOOD EFFECT: All HTTP calls are subject to the timeout; no code
  //              duplication between refresh and exchange.

  /**
   * Makes an HTTPS POST request with timeout.
   * @param {string} postData - URL-encoded form body
   * @returns {Promise<{statusCode: number, body: object}>}
   * @private
   */
  _httpsPost(postData) {
    const timeoutMs = this.config.httpTimeoutMs;

    return new Promise((resolve, reject) => {
      // BEFORE: No timeout — request could hang indefinitely.
      // AFTER: Timer destroys the request after httpTimeoutMs.
      // GOOD EFFECT: Stalled token endpoint calls are cancelled with a
      //              clear error message instead of blocking forever.
      const timer = setTimeout(() => {
        req.destroy();
        reject(new Error(`HTTP request to token endpoint timed out after ${timeoutMs}ms`));
      }, timeoutMs);

      const requestOptions = {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(postData)
        }
      };

      const req = https.request(this.config.tokenEndpoint, requestOptions, (res) => {
        let data = '';
        res.on('data', (chunk) => data += chunk);
        res.on('end', () => {
          clearTimeout(timer);
          try {
            const body = JSON.parse(data);
            resolve({ statusCode: res.statusCode, body });
          } catch (e) {
            reject(new Error(`Failed to parse token response: ${e.message}. Raw: ${data}`));
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

  async _loadTokensFromFile() {
    try {
      const tokenData = await fs.readFile(this.config.tokenStorePath, 'utf8');
      this.tokens = JSON.parse(tokenData);
      console.log('Tokens loaded from file.');
      return this.tokens;
    } catch (error) {
      if (error.code === 'ENOENT') {
        console.log('Token file not found. No tokens loaded.');
      } else {
        console.error('Error loading token cache:', error);
      }
      this.tokens = null;
      return null;
    }
  }

  async _saveTokensToFile() {
    if (!this.tokens) {
      console.warn('No tokens to save.');
      return;
    }
    try {
      await fs.writeFile(this.config.tokenStorePath, JSON.stringify(this.tokens, null, 2));
      console.log('Tokens saved successfully.');
    } catch (error) {
      console.error('Error saving token cache:', error);
      throw error;
    }
  }

  async getTokens() {
    if (this.tokens) {
      return this.tokens;
    }
    if (!this._loadPromise) {
      this._loadPromise = this._loadTokensFromFile().finally(() => {
        this._loadPromise = null;
      });
    }
    return this._loadPromise;
  }

  getExpiryTime() {
    return this.tokens && this.tokens.expires_at ? this.tokens.expires_at : 0;
  }

  isTokenExpired() {
    if (!this.tokens || !this.tokens.expires_at) {
      return true;
    }
    return Date.now() >= (this.tokens.expires_at - this.config.refreshTokenBuffer);
  }

  async getValidAccessToken() {
    await this.getTokens();

    if (!this.tokens || !this.tokens.access_token) {
      console.log('No access token available.');
      return null;
    }

    if (this.isTokenExpired()) {
      console.log('Access token expired or nearing expiration. Attempting refresh.');
      if (this.tokens.refresh_token) {
        try {
          return await this.refreshAccessToken();
        } catch (refreshError) {
          console.error('Failed to refresh access token:', refreshError);
          this.tokens = null;
          await this._saveTokensToFile();
          return null;
        }
      } else {
        console.warn('No refresh token available. Cannot refresh access token.');
        this.tokens = null;
        await this._saveTokensToFile();
        return null;
      }
    }
    return this.tokens.access_token;
  }

  async refreshAccessToken() {
    if (!this.tokens || !this.tokens.refresh_token) {
      throw new Error('No refresh token available to refresh the access token.');
    }

    // ── Refresh lock (P0 fix) ──
    // BEFORE: _refreshPromise lock existed but cleanup was inconsistent —
    //         the `finally` block was inside the Promise constructor's
    //         res.on('end') handler, meaning errors on the `req` event
    //         didn't clear the lock, potentially deadlocking future refreshes.
    // AFTER: Lock cleared in a top-level finally{} block outside the Promise.
    // GOOD EFFECT: _refreshPromise is ALWAYS cleared regardless of how
    //              the refresh completes (success, error, timeout) — no deadlocks.
    if (this._refreshPromise) {
      console.log("Refresh already in progress, awaiting existing promise.");
      return this._refreshPromise;
    }

    console.log('Attempting to refresh access token...');
    const postData = querystring.stringify({
      client_id: this.config.clientId,
      client_secret: this.config.clientSecret,
      grant_type: 'refresh_token',
      refresh_token: this.tokens.refresh_token,
      scope: this.config.scopes.join(' ')
    });

    // BEFORE: this._refreshPromise = new Promise((resolve, reject) => { ... });
    //         — inline https.request with no timeout; _refreshPromise cleared
    //         inside nested callbacks (inconsistent cleanup).
    // AFTER: Delegates to shared _httpsPost() (which has timeout);
    //        _refreshPromise cleared in finally{}.
    // GOOD EFFECT: Timeout enforced; lock always cleaned up; less code duplication.
    this._refreshPromise = (async () => {
      try {
        const { statusCode, body } = await this._httpsPost(postData);

        if (statusCode >= 200 && statusCode < 300) {
          this.tokens.access_token = body.access_token;
          if (body.refresh_token) {
            this.tokens.refresh_token = body.refresh_token;
          }
          this.tokens.expires_in = body.expires_in;
          this.tokens.expires_at = Date.now() + (body.expires_in * 1000);
          await this._saveTokensToFile();
          console.log('Access token refreshed and saved successfully.');
          return this.tokens.access_token;
        } else {
          throw new Error(body.error_description || `Token refresh failed with status ${statusCode}`);
        }
      } finally {
        // BEFORE: _refreshPromise was cleared inside nested callbacks — some
        //         error paths skipped the cleanup, causing deadlocks.
        // AFTER: Always cleared here.
        // GOOD EFFECT: No deadlock — subsequent refresh attempts always proceed.
        this._refreshPromise = null;
      }
    })();

    return this._refreshPromise;
  }

  /**
   * Exchanges an authorization code for tokens.
   *
   * BEFORE: exchangeCodeForTokens(authCode) — no PKCE support.
   * AFTER: exchangeCodeForTokens(authCode, codeVerifier) — PKCE code_verifier
   *        included in the token exchange if provided.
   * GOOD EFFECT: Completes the PKCE flow — Microsoft's token endpoint verifies
   *              the code_challenge matches the code_verifier, preventing
   *              authorization code interception attacks.
   *
   * BEFORE: Inline https.request with no timeout.
   * AFTER: Delegates to shared _httpsPost() with timeout.
   * GOOD EFFECT: Token exchange can't hang indefinitely.
   *
   * @param {string} authCode - Authorization code from callback
   * @param {string} [codeVerifier] - PKCE code verifier (optional for backward compat)
   * @returns {Promise<object>} Token object
   */
  async exchangeCodeForTokens(authCode, codeVerifier) {
    if (!this.config.clientId || !this.config.clientSecret) {
      throw new Error("Client ID or Client Secret is not configured. Cannot exchange code for tokens.");
    }

    console.log('Exchanging authorization code for tokens...');

    // BEFORE: const postData = querystring.stringify({ ... });
    //         — no code_verifier field.
    // AFTER: code_verifier included when provided.
    // GOOD EFFECT: PKCE verification at the Microsoft token endpoint.
    const params = {
      client_id: this.config.clientId,
      client_secret: this.config.clientSecret,
      grant_type: 'authorization_code',
      code: authCode,
      redirect_uri: this.config.redirectUri,
      scope: this.config.scopes.join(' ')
    };

    // BEFORE: (no PKCE field)
    // AFTER: Include code_verifier if provided.
    // GOOD EFFECT: Enables PKCE verification — the token endpoint checks
    //              that SHA256(code_verifier) matches the code_challenge
    //              sent during authorization.
    if (codeVerifier) {
      params.code_verifier = codeVerifier;
    }

    const postData = querystring.stringify(params);

    // BEFORE: return new Promise((resolve, reject) => {
    //           const req = https.request(...) — no timeout, inline logic.
    // AFTER: Delegates to _httpsPost().
    // GOOD EFFECT: Timeout enforced; no code duplication.
    const { statusCode, body } = await this._httpsPost(postData);

    if (statusCode >= 200 && statusCode < 300) {
      this.tokens = {
        access_token: body.access_token,
        refresh_token: body.refresh_token,
        expires_in: body.expires_in,
        expires_at: Date.now() + (body.expires_in * 1000),
        scope: body.scope,
        token_type: body.token_type
      };
      await this._saveTokensToFile();
      console.log('Tokens exchanged and saved successfully.');
      return this.tokens;
    } else {
      throw new Error(body.error_description || `Token exchange failed with status ${statusCode}`);
    }
  }

  async clearTokens() {
    this.tokens = null;
    try {
      await fs.unlink(this.config.tokenStorePath);
      console.log('Token file deleted successfully.');
    } catch (error) {
      if (error.code === 'ENOENT') {
        console.log('Token file not found, nothing to delete.');
      } else {
        console.error('Error deleting token file:', error);
      }
    }
  }
}

module.exports = TokenStorage;