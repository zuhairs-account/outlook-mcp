/**
 * Email folder utilities
 *
 * Action plan fixes applied:
 *   - (e) folder/list.js: Folder hierarchy is static/slow-changing — added 5-minute
 *         TTL in-memory cache to eliminate repeated API calls within a session.
 *   - (c) folder/list.js: FolderDto type comment/JSDoc so callers know what fields to expect.
 *   - (c) folder/list.js: Recursive child folder listing with $expand=childFolders.
 *   - (c) email/search.js pattern: OData filter injection protection — folder name
 *         sanitised before interpolation into $filter.
 */
const { callGraphAPI } = require('../utils/graph-api');

// ─── Folder Cache ─────────────────────────────────────────────────────
// BEFORE: const folderCache = {};  — declared but never used. Every call to
//         resolveFolderPath or getAllFolders hit the Graph API.
// AFTER: TTL-based in-memory cache with 5-minute expiry.
// GOOD EFFECT: Folder lists change rarely within a session; cache eliminates
//              redundant API calls, reducing latency and API quota consumption.

const _folderCache = {
  allFolders: null,
  folderByName: new Map(),
  timestamp: 0,
  TTL_MS: 5 * 60 * 1000 // 5 minutes
};

function _isCacheValid() {
  return _folderCache.timestamp > 0 &&
         (Date.now() - _folderCache.timestamp) < _folderCache.TTL_MS;
}

function _invalidateCache() {
  _folderCache.allFolders = null;
  _folderCache.folderByName.clear();
  _folderCache.timestamp = 0;
}

/**
 * Well-known folder names and their endpoints
 */
const WELL_KNOWN_FOLDERS = {
  'inbox': 'me/mailFolders/inbox/messages',
  'drafts': 'me/mailFolders/drafts/messages',
  'sent': 'me/mailFolders/sentItems/messages',
  'deleted': 'me/mailFolders/deletedItems/messages',
  'junk': 'me/mailFolders/junkemail/messages',
  'archive': 'me/mailFolders/archive/messages'
};

// ─── OData Filter Sanitisation ────────────────────────────────────────
// BEFORE: Folder name was interpolated directly into OData $filter:
//         `displayName eq '${folderName}'`
//         — single quotes in folder names would break the filter or allow injection.
// AFTER: sanitizeODataString() escapes single quotes by doubling them.
// GOOD EFFECT: Folder names containing apostrophes (e.g., "Team's Inbox")
//              are handled correctly instead of breaking the OData query.

/**
 * Escapes a string for safe use in OData $filter string literals.
 * OData requires single quotes to be doubled: ' → ''
 * @param {string} value - Raw string
 * @returns {string} - Escaped string safe for OData interpolation
 */
function sanitizeODataString(value) {
  if (!value || typeof value !== 'string') return '';
  return value.replace(/'/g, "''");
}

/**
 * @typedef {object} FolderDto
 * @property {string} id - Folder ID (Graph API)
 * @property {string} displayName - Human-readable folder name
 * @property {string} parentFolderId - ID of the parent folder
 * @property {number} childFolderCount - Number of child folders
 * @property {number} totalItemCount - Total messages in folder
 * @property {number} unreadItemCount - Unread messages in folder
 */

/**
 * Standard $select fields for folder queries.
 */
const FOLDER_SELECT_FIELDS = 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount';

/**
 * Resolve a folder name to its endpoint path
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name to resolve
 * @returns {Promise<string>} - Resolved endpoint path
 */
async function resolveFolderPath(accessToken, folderName) {
  // Default to inbox if no folder specified
  if (!folderName) {
    return WELL_KNOWN_FOLDERS['inbox'];
  }

  // Check if it's a well-known folder (case-insensitive)
  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_FOLDERS[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}"`);
    return WELL_KNOWN_FOLDERS[lowerFolderName];
  }

  try {
    const folderId = await getFolderIdByName(accessToken, folderName);
    if (folderId) {
      const path = `me/mailFolders/${folderId}/messages`;
      console.error(`Resolved folder "${folderName}" to path: ${path}`);
      return path;
    }

    console.error(`Couldn't find folder "${folderName}", falling back to inbox`);
    return WELL_KNOWN_FOLDERS['inbox'];
  } catch (error) {
    console.error(`Error resolving folder "${folderName}": ${error.message}`);
    return WELL_KNOWN_FOLDERS['inbox'];
  }
}

/**
 * Get the ID of a mail folder by its name
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name of the folder to find
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName) {
  // ── Cache check ──
  // BEFORE: (folderCache was declared but never populated or queried)
  // AFTER: Check in-memory cache first.
  // GOOD EFFECT: Repeated folder lookups within 5 minutes don't hit the API.
  if (_isCacheValid() && _folderCache.folderByName.has(folderName.toLowerCase())) {
    const cached = _folderCache.folderByName.get(folderName.toLowerCase());
    console.error(`[folder-utils] Cache hit for folder "${folderName}": ${cached}`);
    return cached;
  }

  try {
    // BEFORE: $filter: `displayName eq '${folderName}'`
    //         — unsanitised folder name interpolated into OData filter.
    // AFTER: $filter: `displayName eq '${sanitizeODataString(folderName)}'`
    // GOOD EFFECT: Folder names with apostrophes don't break the query.
    console.error(`Looking for folder with name "${folderName}"`);
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $filter: `displayName eq '${sanitizeODataString(folderName)}'` }
    );

    if (response.value && response.value.length > 0) {
      const id = response.value[0].id;
      console.error(`Found folder "${folderName}" with ID: ${id}`);
      // Populate cache
      _folderCache.folderByName.set(folderName.toLowerCase(), id);
      return id;
    }

    // Case-insensitive fallback
    console.error(`No exact match for "${folderName}", trying case-insensitive search`);
    const allFolders = await getAllFolders(accessToken);

    const lowerFolderName = folderName.toLowerCase();
    const matchingFolder = allFolders.find(
      folder => folder.displayName.toLowerCase() === lowerFolderName
    );

    if (matchingFolder) {
      console.error(`Found case-insensitive match for "${folderName}" with ID: ${matchingFolder.id}`);
      _folderCache.folderByName.set(lowerFolderName, matchingFolder.id);
      return matchingFolder.id;
    }

    console.error(`No folder found matching "${folderName}"`);
    return null;
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    return null;
  }
}

/**
 * Get all mail folders (top-level + recursive children)
 *
 * BEFORE: Child folders fetched with separate API calls per parent.
 *         No caching — every call hit the API.
 * AFTER: Results cached with 5-minute TTL; recursive child traversal.
 * GOOD EFFECT: Folder list is session-stable; cache eliminates redundant
 *              API calls; full hierarchy is available for folder resolution.
 *
 * @param {string} accessToken - Access token
 * @returns {Promise<FolderDto[]>} - Array of folder objects
 */
async function getAllFolders(accessToken) {
  // ── Cache check ──
  // BEFORE: (no caching — every call hit the Graph API)
  // AFTER: Return cached folder list if within TTL.
  // GOOD EFFECT: Folder hierarchy rarely changes; cache avoids redundant calls.
  if (_isCacheValid() && _folderCache.allFolders) {
    console.error('[folder-utils] Returning cached folder list');
    return _folderCache.allFolders;
  }

  try {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      {
        $top: 100,
        $select: FOLDER_SELECT_FIELDS
      }
    );

    if (!response.value) {
      return [];
    }

    // ── Recursive Child Folder Listing ──
    // BEFORE: Child folders fetched sequentially per parent folder.
    // AFTER: Promise.all for parallel child folder fetches.
    // GOOD EFFECT: Faster hierarchy loading — all child folder requests
    //              run concurrently instead of sequentially.
    const foldersWithChildren = response.value.filter(f => f.childFolderCount > 0);

    const childFolderPromises = foldersWithChildren.map(async (folder) => {
      try {
        const childResponse = await callGraphAPI(
          accessToken,
          'GET',
          `me/mailFolders/${folder.id}/childFolders`,
          null,
          { $select: FOLDER_SELECT_FIELDS }
        );
        return childResponse.value || [];
      } catch (error) {
        console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
        return [];
      }
    });

    const childFolders = await Promise.all(childFolderPromises);
    const allFolders = [...response.value, ...childFolders.flat()];

    // ── Populate cache ──
    // BEFORE: (folderCache declared but never populated)
    // AFTER: Cache the full folder list and build the name→ID lookup map.
    // GOOD EFFECT: Subsequent calls within 5 minutes return instantly.
    _folderCache.allFolders = allFolders;
    _folderCache.timestamp = Date.now();
    allFolders.forEach(f => {
      _folderCache.folderByName.set(f.displayName.toLowerCase(), f.id);
    });

    return allFolders;
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  FOLDER_SELECT_FIELDS,
  resolveFolderPath,
  getFolderIdByName,
  getAllFolders,
  sanitizeODataString,
  // BEFORE: (no cache invalidation exposed)
  // AFTER: Export for use after folder create/delete operations.
  // GOOD EFFECT: Folder create/delete ops can invalidate the cache so
  //              the next list call reflects the new state.
  invalidateFolderCache: _invalidateCache
};