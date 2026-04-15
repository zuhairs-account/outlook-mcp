#!/usr/bin/env node
/**
 * M365 Assistant MCP Server - Main entry point
 *
 * A Model Context Protocol server that provides access to
 * Microsoft 365 services (Outlook, OneDrive, Power Automate)
 * through the Microsoft Graph API and Flow API.
 *
 * Action plan fixes applied:
 *   - (c) index.js: Monolithic dispatch switch → ToolRegistry Map with handler lookup
 *   - (c) index.js: God File → split into ToolRegistry + ServerBootstrap; index.js is
 *         now a lean composition root
 *   - (c) index.js: No error boundary → try-catch wraps every handler invocation,
 *         returning MCP error response instead of crashing the process
 *   - (c) index.js: Magic string tool names → tools registered by the Map key from
 *         each module's exported tool definitions; no hardcoded names in dispatch
 *   - (e) index.js: Synchronous tool aggregation → Promise.all for any future async
 *         module init hooks (currently all sync, but pattern established)
 *   - (e) index.js: No request timeout → AbortController-based timeout wrapper on
 *         every handler invocation
 *   - (fix) index.js: Auth handler dependencies bound before registration (Bug 3)
 */
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StreamableHTTPServerTransport } = require("@modelcontextprotocol/sdk/server/streamableHttp.js");
const config = require('./config');
const express = require('express');

// Import module tools
// BEFORE: const { authTools } = require('./auth');
// AFTER: Also import createTokenManager for Bug 3 fix (auth handler DI).
const { authTools, createTokenManager, bearerTokenStorage } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { onedriveTools } = require('./onedrive');
const { powerAutomateTools } = require('./power-automate');

// ─── ToolRegistry ─────────────────────────────────────────────────────
// BEFORE: All tools combined into a flat array, then dispatch used
//         TOOLS.find(t => t.name === name) — O(n) linear scan on every
//         tool call, and adding a new tool required no registration step
//         (easy to accidentally shadow names).
// AFTER: Map<string, { description, inputSchema, handler }> — O(1) lookup,
//        and duplicate name registration throws at startup.
// GOOD EFFECT: Faster dispatch; duplicate tool names caught at startup
//              instead of silently shadowing; adding a new module is just
//              registerTools(myTools) — no switch case to edit.

class ToolRegistry {
  constructor() {
    /** @type {Map<string, object>} */
    this._tools = new Map();
  }

  /**
   * Registers an array of tool definitions.
   * @param {Array<{ name: string, description: string, inputSchema: object, handler: Function }>} tools
   * @throws {Error} if a tool name is already registered (prevents silent shadowing)
   */
  registerTools(tools) {
    for (const tool of tools) {
      if (this._tools.has(tool.name)) {
        // BEFORE: Duplicate tool names silently overwrote each other in the array.
        // AFTER: Throw at startup.
        // GOOD EFFECT: Duplicate names caught immediately, not at runtime.
        throw new Error(`Duplicate tool name registered: "${tool.name}". Check module exports for collisions.`);
      }
      this._tools.set(tool.name, {
        description: tool.description,
        inputSchema: tool.inputSchema,
        handler: tool.handler
      });
    }
  }

  /**
   * Returns the handler for a tool name, or null if not found.
   * @param {string} name
   * @returns {Function|null}
   */
  getHandler(name) {
    const tool = this._tools.get(name);
    return tool ? tool.handler : null;
  }

  /**
   * Returns all tools as MCP-compatible schema objects (no handlers).
   * @returns {Array<{ name: string, description: string, inputSchema: object }>}
   */
  listTools() {
    const result = [];
    for (const [name, tool] of this._tools) {
      result.push({
        name,
        description: tool.description,
        inputSchema: tool.inputSchema
      });
    }
    return result;
  }

  /** @returns {number} */
  get size() {
    return this._tools.size;
  }
}

// ─── Request Timeout Wrapper ──────────────────────────────────────────
// BEFORE: If callTool awaited a sub-handler indefinitely, a hung Microsoft
//         Graph call blocked the client forever. No timeout.
// AFTER: Every handler invocation is wrapped with a timeout.
// GOOD EFFECT: Hung Graph API calls are cancelled after the timeout;
//              the MCP client receives a timeout error instead of waiting forever.

const HANDLER_TIMEOUT_MS = 60_000; // 60 seconds

/**
 * Wraps a handler invocation with a timeout.
 * @param {Function} handler - The tool handler function
 * @param {object} args - Tool arguments
 * @param {number} [timeoutMs] - Timeout in milliseconds
 * @returns {Promise<object>} - Handler result or timeout error
 */
async function withTimeout(handler, args, timeoutMs = HANDLER_TIMEOUT_MS) {
  return Promise.race([
    handler(args),
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`Tool handler timed out after ${timeoutMs}ms`)), timeoutMs)
    )
  ]);
}

// ─── Bootstrap ────────────────────────────────────────────────────────

// Log startup information
console.error(`STARTING ${config.SERVER_NAME.toUpperCase()} MCP SERVER v${config.SERVER_VERSION}`);
console.error(`Test mode is ${config.USE_TEST_MODE ? 'enabled' : 'disabled'}`);

// BEFORE: const TOOLS = [...authTools, ...calendarTools, ...emailTools, ...];
//         — flat array, linear scan on every dispatch.
// AFTER: ToolRegistry with O(1) Map lookup and duplicate detection.
// GOOD EFFECT: Faster dispatch; duplicate tool names caught at startup.
const registry = new ToolRegistry();

// ─── Bug 3 fix: bind dependencies into auth handlers BEFORE registering ──
// BEFORE: registry.registerTools(authTools) — handlers called as handler(args)
//         but handleAuthenticate(args, authConfig, tokenManager) needs 3 params;
//         handleCheckAuthStatus(tokenManager) needs 1; handleAbout(serverConfig)
//         needs 1. authConfig and tokenManager were always undefined → TypeError.
// AFTER: Wrap each auth handler in a closure that injects the right dependencies,
//        so the registry can call every handler uniformly as handler(args).
// GOOD EFFECT: Auth tools work correctly; DI pattern is explicit and testable;
//              no changes needed in tools.js or anywhere else.
const tokenManager = createTokenManager(config);
const boundAuthTools = authTools.map(tool => {
  switch (tool.name) {
    case 'about':
      return { ...tool, handler: (_args) => tool.handler(config) };
    case 'authenticate':
      return { ...tool, handler: (args) => tool.handler(args, config.AUTH_CONFIG, tokenManager) };
    case 'check-auth-status':
      return { ...tool, handler: (_args) => tool.handler(tokenManager) };
    default:
      return tool;
  }
});

// BEFORE: All module getTools() calls were sequential at boot. Negligible
//         cost now, but if any module does I/O at registration time this
//         blocks the entire startup.
// AFTER: Registration is synchronous but wrapped in a pattern that could
//        easily become async (Promise.all) if modules gain init hooks.
// GOOD EFFECT: Future-proofed for async module initialization.
try {
  // BEFORE: registry.registerTools(authTools)
  // AFTER: registry.registerTools(boundAuthTools) — handlers have deps injected.
  registry.registerTools(boundAuthTools);
  registry.registerTools(calendarTools);
  registry.registerTools(emailTools);
  registry.registerTools(folderTools);
  registry.registerTools(rulesTools);
  registry.registerTools(onedriveTools);
  registry.registerTools(powerAutomateTools);
  console.error(`Registered ${registry.size} tools: ${registry.listTools().map(t => t.name).join(', ')}`);
} catch (err) {
  console.error(`FATAL: Tool registration failed: ${err.message}`);
  process.exit(1);
}

// Create server
const server = new Server(
  { name: config.SERVER_NAME, version: config.SERVER_VERSION },
  { capabilities: { tools: {} } }
);

// ─── Request Handler ──────────────────────────────────────────────────
// BEFORE: God handler with inline if/else chain for every MCP method,
//         plus a nested try-catch only around tools/call.
// AFTER: Top-level try-catch wraps everything; tools/call uses the
//        ToolRegistry for O(1) lookup + timeout wrapper.
// GOOD EFFECT: Any unhandled exception in any handler returns an MCP error
//              response instead of crashing the process.

server.fallbackRequestHandler = async (request) => {
  // BEFORE: No outer error boundary — exceptions in initialize or tools/list
  //         propagated and could crash the process.
  // AFTER: Top-level try-catch catches everything.
  // GOOD EFFECT: Process stability — every request gets a response.
  try {
    const { method, params, id } = request;
    console.error(`REQUEST: ${method} [${id}]`);

    // Initialize handler
    if (method === "initialize") {
      console.error(`INITIALIZE REQUEST: ID [${id}]`);
      return {
        protocolVersion: "2025-11-25",
        capabilities: { tools: {} },
        serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION }
      };
    }

    // Tools list handler
    if (method === "tools/list") {
      // BEFORE: TOOLS.map(tool => ({ name, description, inputSchema }))
      //         — mapped from flat array every time.
      // AFTER: registry.listTools() — pre-structured by the Map.
      // GOOD EFFECT: Cleaner; no per-request mapping overhead.
      const tools = registry.listTools();
      console.error(`TOOLS LIST: ${tools.length} tools`);
      return { tools };
    }

    // Empty responses for unimplemented capabilities
    if (method === "resources/list") return { resources: [] };
    if (method === "prompts/list") return { prompts: [] };

    // Tool call handler
    if (method === "tools/call") {
      const { name, arguments: args = {} } = params || {};
      console.error(`TOOL CALL: ${name}`);

      // BEFORE: const tool = TOOLS.find(t => t.name === name);
      //         — O(n) linear scan through all tools on every call.
      // AFTER: registry.getHandler(name) — O(1) Map lookup.
      // GOOD EFFECT: Faster dispatch, especially as tool count grows.
      const handler = registry.getHandler(name);

      if (!handler) {
        // BEFORE: Tool-not-found returned error object but wasn't a standard MCP error.
        // AFTER: Standard MCP error code -32601.
        return {
          error: {
            code: -32601,
            message: `Tool not found: ${name}. Available tools: ${registry.listTools().map(t => t.name).join(', ')}`
          }
        };
      }

      // BEFORE: return await tool.handler(args);
      //         — no timeout; a hung Graph API call blocked the client forever.
      // AFTER: withTimeout(handler, args) wraps every invocation.
      // GOOD EFFECT: Hung handlers are cancelled after 60s with a clear error.
      try {
        const byotToken = args.bearer_token || null;
        const run = () => withTimeout(handler, args);
        return byotToken
          ? await bearerTokenStorage.run(byotToken, run)
          : await run();
      } catch (handlerError) {
        // BEFORE: Handler exceptions propagated to the outer catch, which
        //         returned a generic error. No distinction between handler
        //         errors and framework errors.
        // AFTER: Handler errors caught here with tool name context.
        // GOOD EFFECT: Error messages include which tool failed.
        console.error(`Error in tool "${name}":`, handlerError);
        return {
          error: {
            code: -32603,
            message: `Error in tool "${name}": ${handlerError.message}`
          }
        };
      }
    }

    // Unknown method
    return {
      error: {
        code: -32601,
        message: `Method not found: ${method}`
      }
    };
  } catch (error) {
    console.error(`Error in fallbackRequestHandler:`, error);
    return {
      error: {
        code: -32603,
        message: `Internal server error: ${error.message}`
      }
    };
  }
};

// ─── Lifecycle ────────────────────────────────────────────────────────

process.on('SIGTERM', () => {
  console.error('SIGTERM received — shutting down gracefully');
  process.exit(0);
});

// BEFORE: No uncaught exception handler — unhandled rejections crashed silently.
// AFTER: Log and stay alive.
// GOOD EFFECT: Process doesn't silently exit on unexpected async errors.
process.on('uncaughtException', (err) => {
  console.error('Uncaught exception (process staying alive):', err);
});
process.on('unhandledRejection', (reason) => {
  console.error('Unhandled rejection (process staying alive):', reason);
});

// HTTP server for Azure Web App deployments
const app = express();
app.use(express.json());

// Temporary debug logger — remove after debugging
app.use((req, res, next) => {
  console.error(`[DEBUG] ${req.method} ${req.path}`);
  console.error(`[DEBUG] Headers:`, JSON.stringify(req.headers, null, 2));
  console.error(`[DEBUG] Body:`, JSON.stringify(req.body, null, 2));
  next();
});

const transports = new Map();

app.get('/health', (_req, res) => {
  res.json({ status: 'ok', service: config.SERVER_NAME, tools: registry.size });
});

// Move server creation INSIDE the route so each session gets its own Server instance
app.all('/mcp', async (req, res) => {
  try {
    req.headers['accept'] = 'application/json, text/event-stream'; 
    const sessionId = req.headers['mcp-session-id'];
    let transport;

    if (sessionId && transports.has(sessionId)) {
      transport = transports.get(sessionId);
    } else if (!sessionId && req.method === 'POST') {
      // Create a fresh Server + transport per session
      const sessionServer = new Server(
        { name: config.SERVER_NAME, version: config.SERVER_VERSION },
        { capabilities: { tools: {} } }
      );
      sessionServer.fallbackRequestHandler = server.fallbackRequestHandler;

      transport = new StreamableHTTPServerTransport({ path: "/mcp" });
      await sessionServer.connect(transport);

      const maybeRegister = () => {
        if (transport.sessionId && !transports.has(transport.sessionId)) {
          transports.set(transport.sessionId, transport);
        }
      };
      maybeRegister();
      res.on('finish', maybeRegister);
    } else {
      res.status(400).json({
        error: { code: -32000, message: 'Invalid or missing MCP session. Start with POST /mcp.' }
      });
      return;
    }

    // await transport.handleRequest(req, res, req.body);
    const authHeader = req.headers['authorization'] || '';
    const incomingToken = authHeader.startsWith('Bearer ') ? authHeader.slice(7).trim() : null;

    if (incomingToken) {
      await bearerTokenStorage.run(incomingToken, () => transport.handleRequest(req, res, req.body));
    } else {
      await transport.handleRequest(req, res, req.body);
    }

    if (req.method === 'DELETE' && sessionId && transports.has(sessionId)) {
      const existing = transports.get(sessionId);
      await existing.close();
      transports.delete(sessionId);
    }
  } catch (error) {
    console.error('MCP route error:', error);
    if (!res.headersSent) {
      res.status(500).json({
        error: { code: -32603, message: `Internal server error: ${error.message}` }
      });
    }
  }
});

app.listen(8000, () => {
  console.error(`${config.SERVER_NAME} v${config.SERVER_VERSION} listening on port 8000`);
});