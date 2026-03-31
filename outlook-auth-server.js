#!/usr/bin/env node
const http = require('http');
const url = require('url');
const querystring = require('querystring');
const https = require('https');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

// Load environment variables from .env file
require('dotenv').config();

const { AccountDB } = require('./auth/account-db');

// Log to console
console.log('Starting Outlook Authentication Server');

// HTML escaping to prevent reflected XSS
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// CSRF state store (state -> { timestamp, account })
const pendingStates = new Map();
const TEN_MINUTES = 10 * 60 * 1000;

// Periodically clean up expired CSRF states
setInterval(() => {
  const now = Date.now();
  for (const [key, val] of pendingStates.entries()) {
    if (now - val.timestamp > TEN_MINUTES) pendingStates.delete(key);
  }
}, 5 * 60 * 1000).unref();

// Account database
const accountDB = new AccountDB();

// Authentication configuration
const AUTH_CONFIG = {
  clientId: process.env.MS_CLIENT_ID || '',
  clientSecret: process.env.MS_CLIENT_SECRET || '',
  tenantId: process.env.MS_TENANT_ID || 'common',
  authorityHost: (process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, ''),
  redirectUri: 'http://localhost:3333/auth/callback',
  scopes: [
    'offline_access',
    'User.Read',
    'Mail.Read',
    'Mail.ReadWrite',
    'Mail.Send',
    'MailboxSettings.Read',
    'Calendars.Read',
    'Calendars.ReadWrite',
    'Files.Read',
    'Files.ReadWrite'
  ]
};

/**
 * Fetch the authenticated user's primary email via Graph API.
 */
function getUserEmail(accessToken) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'graph.microsoft.com',
      path: '/v1.0/me?$select=mail,userPrincipalName',
      method: 'GET',
      headers: { 'Authorization': `Bearer ${accessToken}` }
    };
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const body = JSON.parse(data);
          resolve(body.mail || body.userPrincipalName || null);
        } catch (e) {
          reject(e);
        }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

/**
 * Exchange authorization code for tokens.
 */
function exchangeCodeForTokens(code) {
  return new Promise((resolve, reject) => {
    const postData = querystring.stringify({
      client_id: AUTH_CONFIG.clientId,
      client_secret: AUTH_CONFIG.clientSecret,
      code: code,
      redirect_uri: AUTH_CONFIG.redirectUri,
      grant_type: 'authorization_code',
      scope: AUTH_CONFIG.scopes.join(' ')
    });

    const options = {
      hostname: AUTH_CONFIG.authorityHost.replace(/^https?:\/\//, '').split('/')[0],
      path: `/${AUTH_CONFIG.tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const tokenResponse = JSON.parse(data);
            tokenResponse.expires_at = Date.now() + (tokenResponse.expires_in * 1000);
            resolve(tokenResponse);
          } catch (error) {
            reject(new Error(`Error parsing token response: ${error.message}`));
          }
        } else {
          reject(new Error(`Token exchange failed with status ${res.statusCode}: ${data}`));
        }
      });
    });
    req.on('error', reject);
    req.write(postData);
    req.end();
  });
}

/**
 * Parse JSON body from request.
 */
function parseBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', () => {
      try { resolve(JSON.parse(body)); }
      catch (e) { reject(new Error('Invalid JSON body')); }
    });
    req.on('error', reject);
  });
}

/**
 * Send JSON response.
 */
function jsonResponse(res, statusCode, data) {
  res.writeHead(statusCode, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify(data, null, 2));
}

// Create HTTP server
const server = http.createServer(async (req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;

  console.log(`Request received: ${req.method} ${pathname}`);

  try {
    // ===== AUTH ROUTES =====

    if (pathname === '/auth' && req.method === 'GET') {
      if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
        res.writeHead(500, { 'Content-Type': 'text/html' });
        res.end(`<html><body><h1>Configuration Error</h1><p>MS_CLIENT_ID and MS_CLIENT_SECRET are not set.</p></body></html>`);
        return;
      }

      // Optional: ?account=email to tag which account is being authenticated
      const accountHint = parsedUrl.query.account || '';

      const state = crypto.randomBytes(32).toString('hex');
      pendingStates.set(state, { timestamp: Date.now(), account: accountHint });

      const authParams = {
        client_id: AUTH_CONFIG.clientId,
        response_type: 'code',
        redirect_uri: AUTH_CONFIG.redirectUri,
        scope: AUTH_CONFIG.scopes.join(' '),
        response_mode: 'query',
        state
      };

      // If account hint provided, add login_hint for smoother UX
      if (accountHint) {
        authParams.login_hint = accountHint;
      }

      const authUrl = `${AUTH_CONFIG.authorityHost}/${AUTH_CONFIG.tenantId}/oauth2/v2.0/authorize?${querystring.stringify(authParams)}`;
      console.log(`Redirecting to Microsoft login (account hint: ${accountHint || 'none'})`);

      res.writeHead(302, { 'Location': authUrl });
      res.end();

    } else if (pathname === '/auth/callback' && req.method === 'GET') {
      const query = parsedUrl.query;

      if (!query.state || !pendingStates.has(query.state)) {
        res.writeHead(403, { 'Content-Type': 'text/html' });
        res.end(`<html><body><h1>Invalid State</h1><p>Invalid or expired OAuth state. Please try again.</p></body></html>`);
        return;
      }

      const stateData = pendingStates.get(query.state);
      pendingStates.delete(query.state);

      if (query.error) {
        res.writeHead(400, { 'Content-Type': 'text/html' });
        res.end(`<html><body><h1>Auth Error</h1><p>${escapeHtml(query.error)}: ${escapeHtml(query.error_description)}</p></body></html>`);
        return;
      }

      if (!query.code) {
        res.writeHead(400, { 'Content-Type': 'text/html' });
        res.end(`<html><body><h1>Missing Code</h1><p>No authorization code provided.</p></body></html>`);
        return;
      }

      console.log('Authorization code received, exchanging for tokens...');
      const tokens = await exchangeCodeForTokens(query.code);

      // Determine the user's primary email
      let primaryEmail = stateData.account;
      if (!primaryEmail) {
        primaryEmail = await getUserEmail(tokens.access_token);
      }

      if (primaryEmail) {
        primaryEmail = primaryEmail.toLowerCase();
        // Save to account-specific token file
        const tokenFile = AccountDB.tokenFileForEmail(primaryEmail);
        fs.writeFileSync(tokenFile, JSON.stringify(tokens, null, 2), { encoding: 'utf8', mode: 0o600 });
        console.log(`Tokens saved to ${tokenFile} for account ${primaryEmail}`);

        // Auto-register in AccountDB
        await accountDB.addAccount(primaryEmail, { tokenFile });
        console.log(`Account ${primaryEmail} registered in accounts database`);
      }

      // Also save to legacy token path for backward compatibility
      const legacyPath = path.join(process.env.HOME || process.env.USERPROFILE, '.outlook-mcp-tokens.json');
      fs.writeFileSync(legacyPath, JSON.stringify(tokens, null, 2), { encoding: 'utf8', mode: 0o600 });

      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(`
        <html><body>
          <h1 style="color: #5cb85c;">Authentication Successful!</h1>
          <p>Account: <strong>${escapeHtml(primaryEmail || 'unknown')}</strong></p>
          <p>Tokens saved. You can close this window.</p>
        </body></html>
      `);

    // ===== ACCOUNT API ROUTES =====

    } else if (pathname === '/accounts' && req.method === 'GET') {
      await accountDB.load();
      jsonResponse(res, 200, accountDB.data);

    } else if (pathname === '/accounts/aliases' && req.method === 'POST') {
      const body = await parseBody(req);
      if (!body.account || !body.alias) {
        jsonResponse(res, 400, { error: 'Both "account" and "alias" fields are required' });
        return;
      }
      const result = await accountDB.addAlias(body.account, body.alias);
      jsonResponse(res, 200, { message: `Alias ${body.alias} added to ${body.account}`, account: result });

    } else if (pathname === '/accounts/aliases' && req.method === 'DELETE') {
      const body = await parseBody(req);
      if (!body.account || !body.alias) {
        jsonResponse(res, 400, { error: 'Both "account" and "alias" fields are required' });
        return;
      }
      const result = await accountDB.removeAlias(body.account, body.alias);
      jsonResponse(res, 200, { message: `Alias ${body.alias} removed from ${body.account}`, account: result });

    } else if (pathname === '/accounts/default' && req.method === 'POST') {
      const body = await parseBody(req);
      if (!body.account) {
        jsonResponse(res, 400, { error: '"account" field is required' });
        return;
      }
      await accountDB.setDefaultAccount(body.account);
      jsonResponse(res, 200, { message: `Default account set to ${body.account}` });

    } else if (pathname === '/') {
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(`
        <html><body>
          <h1 style="color: #0078d4;">M365 Authentication Server</h1>
          <p>Use the <code>authenticate</code> tool in Claude to start auth.</p>
          <p>API endpoints:</p>
          <ul>
            <li><code>GET /auth?account=email</code> — Start OAuth for an account</li>
            <li><code>GET /accounts</code> — List all accounts and aliases</li>
            <li><code>POST /accounts/aliases</code> — Add alias {"account":"...", "alias":"..."}</li>
            <li><code>DELETE /accounts/aliases</code> — Remove alias</li>
            <li><code>POST /accounts/default</code> — Set default {"account":"..."}</li>
          </ul>
        </body></html>
      `);
    } else {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not Found');
    }
  } catch (error) {
    console.error(`Error handling ${pathname}:`, error.message);
    jsonResponse(res, 500, { error: error.message });
  }
});

// Start server
const PORT = 3333;
server.listen(PORT, () => {
  console.log(`Authentication server running at http://localhost:${PORT}`);
  console.log(`Callback URI: ${AUTH_CONFIG.redirectUri}`);

  if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
    console.log('\nWARNING: MS_CLIENT_ID and MS_CLIENT_SECRET are not set.');
  }
});

process.on('SIGINT', () => { console.log('Shutting down'); process.exit(0); });
process.on('SIGTERM', () => { console.log('Shutting down'); process.exit(0); });
