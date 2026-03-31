/**
 * Account management MCP tools for multi-account support.
 */
const { getAccountDB } = require('./account-db');
const { callGraphAPI } = require('../utils/graph-api');

async function handleListAccounts() {
  const db = getAccountDB();
  if (!db.exists()) {
    return {
      content: [{
        type: "text",
        text: "No accounts configured yet. Use the 'authenticate' tool to add your first account."
      }]
    };
  }

  const addresses = await db.listAllAddresses();
  const defaultAccount = await db.getDefaultAccount();

  if (addresses.length === 0) {
    return {
      content: [{
        type: "text",
        text: "No accounts configured yet. Use the 'authenticate' tool to add your first account."
      }]
    };
  }

  const lines = [];
  let currentAccount = null;
  for (const addr of addresses) {
    if (addr.account !== currentAccount) {
      currentAccount = addr.account;
      const isDefault = currentAccount === defaultAccount ? ' (default)' : '';
      lines.push(`\nAccount: ${currentAccount}${isDefault} [${addr.label}]`);
    }
    if (addr.isAlias) {
      lines.push(`  Alias: ${addr.email}`);
    }
  }

  return {
    content: [{
      type: "text",
      text: `Configured accounts:${lines.join('\n')}`
    }]
  };
}

async function handleAddAlias(args) {
  const { account, alias } = args;
  if (!account || !alias) {
    return { content: [{ type: "text", text: "Both 'account' and 'alias' are required." }] };
  }

  const db = getAccountDB();
  try {
    await db.addAlias(account, alias);
    return { content: [{ type: "text", text: `Alias ${alias} added to account ${account}.` }] };
  } catch (err) {
    return { content: [{ type: "text", text: `Error: ${err.message}` }] };
  }
}

async function handleRemoveAlias(args) {
  const { account, alias } = args;
  if (!account || !alias) {
    return { content: [{ type: "text", text: "Both 'account' and 'alias' are required." }] };
  }

  const db = getAccountDB();
  try {
    await db.removeAlias(account, alias);
    return { content: [{ type: "text", text: `Alias ${alias} removed from account ${account}.` }] };
  } catch (err) {
    return { content: [{ type: "text", text: `Error: ${err.message}` }] };
  }
}

async function handleSetDefaultAccount(args) {
  const { account } = args;
  if (!account) {
    return { content: [{ type: "text", text: "'account' is required." }] };
  }

  const db = getAccountDB();
  try {
    await db.setDefaultAccount(account);
    return { content: [{ type: "text", text: `Default account set to ${account}.` }] };
  } catch (err) {
    return { content: [{ type: "text", text: `Error: ${err.message}` }] };
  }
}

async function handleDiscoverAliases(args) {
  // Lazy require to avoid circular dependency
  const { ensureAuthenticated } = require('./index');
  const account = args && args.account ? args.account : null;

  try {
    const { accessToken, accountEmail } = await ensureAuthenticated(account);

    // Fetch proxyAddresses from Graph API — these are all email aliases
    const response = await callGraphAPI(accessToken, 'GET', 'me', null, {
      $select: 'mail,userPrincipalName,proxyAddresses'
    });

    const primaryEmail = (response.mail || response.userPrincipalName || '').toLowerCase();
    const proxyAddresses = response.proxyAddresses || [];

    // proxyAddresses format: ["SMTP:primary@domain.com", "smtp:alias1@domain.com", "smtp:alias2@domain.com"]
    // Uppercase SMTP = primary, lowercase smtp = alias
    const aliases = [];
    for (const addr of proxyAddresses) {
      if (addr.startsWith('smtp:')) {
        // Lowercase smtp prefix = alias/secondary address
        const email = addr.substring(5).toLowerCase();
        if (email !== primaryEmail) {
          aliases.push(email);
        }
      }
    }

    // Auto-register discovered aliases in account DB
    const db = getAccountDB();
    if (db.exists()) {
      const resolvedAccount = accountEmail || primaryEmail;
      let added = 0;
      for (const alias of aliases) {
        try {
          await db.addAlias(resolvedAccount, alias);
          added++;
        } catch (e) {
          // Account might not exist in DB yet
        }
      }

      const allAddresses = await db.listAllAddresses();
      const accountAliases = allAddresses.filter(a => a.account === resolvedAccount && a.isAlias);

      return {
        content: [{
          type: "text",
          text: `Account: ${resolvedAccount}\nDiscovered ${aliases.length} alias(es) from Microsoft 365:\n${aliases.map(a => `  - ${a}`).join('\n') || '  (none)'}\n\nTotal aliases registered: ${accountAliases.length}\n${accountAliases.map(a => `  - ${a.email}`).join('\n') || '  (none)'}`
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Account: ${primaryEmail}\nDiscovered aliases from Microsoft 365:\n${aliases.map(a => `  - ${a}`).join('\n') || '  (none)'}\n\nNote: No accounts.json found. Authenticate first to auto-register these aliases.`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return { content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }] };
    }
    return { content: [{ type: "text", text: `Error discovering aliases: ${error.message}` }] };
  }
}

const accountTools = [
  {
    name: "list-accounts",
    description: "Lists all configured email accounts and their aliases",
    inputSchema: { type: "object", properties: {}, required: [] },
    handler: handleListAccounts
  },
  {
    name: "add-alias",
    description: "Adds an email alias to an account",
    inputSchema: {
      type: "object",
      properties: {
        account: { type: "string", description: "Primary email address of the account" },
        alias: { type: "string", description: "Alias email address to add" }
      },
      required: ["account", "alias"]
    },
    handler: handleAddAlias
  },
  {
    name: "remove-alias",
    description: "Removes an email alias from an account",
    inputSchema: {
      type: "object",
      properties: {
        account: { type: "string", description: "Primary email address of the account" },
        alias: { type: "string", description: "Alias email address to remove" }
      },
      required: ["account", "alias"]
    },
    handler: handleRemoveAlias
  },
  {
    name: "discover-aliases",
    description: "Fetches email aliases from Microsoft 365 and auto-registers them in the account database",
    inputSchema: {
      type: "object",
      properties: {
        account: { type: "string", description: "Primary email address of the account to discover aliases for. Omit for default account." }
      },
      required: []
    },
    handler: handleDiscoverAliases
  },
  {
    name: "set-default-account",
    description: "Sets the default email account used when no account is specified",
    inputSchema: {
      type: "object",
      properties: {
        account: { type: "string", description: "Primary email address to set as default" }
      },
      required: ["account"]
    },
    handler: handleSetDefaultAccount
  }
];

module.exports = { accountTools };
