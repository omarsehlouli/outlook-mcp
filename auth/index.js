/**
 * Authentication module for Outlook MCP server
 * Supports multi-account: each account has its own token file.
 */
const tokenManager = require('./token-manager');
const TokenStorage = require('./token-storage');
const { authTools } = require('./tools');
const { getAccountDB } = require('./account-db');

// Singleton TokenStorage for legacy single-account mode
const legacyTokenStorage = new TokenStorage();

// Cache of TokenStorage instances per primary account email
const tokenStorageCache = new Map();

/**
 * Get or create a TokenStorage instance for a specific account.
 */
function getTokenStorageForAccount(accountEmail, tokenFile) {
  if (tokenStorageCache.has(accountEmail)) {
    return tokenStorageCache.get(accountEmail);
  }
  const ts = new TokenStorage({ tokenStorePath: tokenFile });
  tokenStorageCache.set(accountEmail, ts);
  return ts;
}

/**
 * Ensures the user is authenticated and returns an access token.
 * Supports multi-account: pass an email (primary or alias) to authenticate as that account.
 * Falls back to legacy single-token mode if no accounts.json exists.
 *
 * @param {string|null} accountOrAlias - Email address (primary or alias) to authenticate as
 * @param {boolean} forceNew - Whether to force a new authentication
 * @returns {Promise<{accessToken: string, accountEmail: string|null}>}
 * @throws {Error} - If authentication fails
 */
async function ensureAuthenticated(accountOrAlias = null, forceNew = false) {
  if (forceNew) {
    throw new Error('Authentication required');
  }

  const accountDB = getAccountDB();

  // If accounts.json exists, use multi-account mode
  if (accountDB.exists()) {
    let accountEmail, tokenFile;

    if (accountOrAlias) {
      const resolved = await accountDB.resolveAddress(accountOrAlias);
      if (!resolved) {
        throw new Error(`Unknown email address: ${accountOrAlias}. Use list-accounts to see configured accounts.`);
      }
      accountEmail = resolved.accountEmail;
      tokenFile = resolved.tokenFile;
    } else {
      // Use default account
      const defaultEmail = await accountDB.getDefaultAccount();
      if (!defaultEmail) {
        throw new Error('No default account configured. Authenticate an account first.');
      }
      const acct = await accountDB.getAccount(defaultEmail);
      if (!acct) {
        throw new Error(`Default account ${defaultEmail} not found in accounts database.`);
      }
      accountEmail = defaultEmail;
      tokenFile = acct.tokenFile;
    }

    const ts = getTokenStorageForAccount(accountEmail, tokenFile);
    const accessToken = await ts.getValidAccessToken();
    if (!accessToken) {
      throw new Error(`Authentication required for account ${accountEmail}. Use the authenticate tool.`);
    }

    return { accessToken, accountEmail };
  }

  // Legacy single-account mode (no accounts.json)
  const accessToken = await legacyTokenStorage.getValidAccessToken();
  if (!accessToken) {
    throw new Error('Authentication required');
  }

  return { accessToken, accountEmail: null };
}

module.exports = {
  tokenManager,
  authTools,
  ensureAuthenticated
};
