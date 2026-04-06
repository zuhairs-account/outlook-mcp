/**
 * Authentication module for Outlook MCP server
 * @module auth
 */

// BEFORE: const tokenManager = require('./token-manager');
// AFTER: Import the class, not a singleton instance.
// GOOD EFFECT: Enables factory-based DI — consumers are decoupled from
//              the concrete token-manager.js implementation.
const TokenManager = require('./token-manager');
const { authTools } = require('./tools');

/**
 * Factory function to create a TokenManager instance with the given config.
 * Consumers should use this instead of importing token-manager directly.
 * @param {object} [config] - Optional config overrides
 * @returns {TokenManager} - A configured TokenManager instance
 *
 * GOOD EFFECT: Enables dependency injection — tests can pass mock config,
 * production code passes real config, and no one is coupled to the concrete
 * file path of token-manager.js.
 */
function createTokenManager(config) {
  return new TokenManager(config);
}

/**
 * Ensures the user is authenticated and returns an access token.
 *
 * @param {TokenManager} tokenManagerInstance - Injected token manager
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<string>} - Access token
 * @throws {Error} - If authentication fails
 */
// AFTER
const config = require('../config');

// Module-level singleton — created once, reused on every call
const _defaultTokenManager = new TokenManager(config);

async function ensureAuthenticated(tokenManagerInstance, forceNew = false) {
  // If no instance injected (the common case across all 28 tool files),
  // fall back to the module-level singleton. Injected instance still works
  // for tests or callers that want to provide their own.
  const tm = tokenManagerInstance || _defaultTokenManager;

  try {
    if (forceNew) {
      throw new Error('Authentication required. Please re-authenticate.');
    }
    const accessToken = await tm.getValidAccessToken();
    if (!accessToken) {
      throw new Error('Authentication required. No valid token found.');
    }
    return accessToken;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Authentication failed: ${message}`);
  }
}

module.exports = {
  // BEFORE: tokenManager (concrete singleton export)
  // AFTER: createTokenManager (factory function)
  // GOOD EFFECT: Barrel enforces the intended interface — consumers go
  //              through the factory and never import internals directly.
  createTokenManager,
  authTools,
  ensureAuthenticated
};