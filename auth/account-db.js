/**
 * Account database for multi-account support.
 * Manages ~/.outlook-mcp-accounts.json with account→token mappings and aliases.
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const config = require('../config');

const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

class AccountDB {
  constructor(dbPath) {
    this.dbPath = dbPath || config.AUTH_CONFIG.accountDbPath;
    this.data = null;
  }

  async load() {
    try {
      const raw = fs.readFileSync(this.dbPath, 'utf8');
      this.data = JSON.parse(raw);
    } catch (err) {
      if (err.code === 'ENOENT') {
        this.data = { accounts: {}, defaultAccount: null };
      } else {
        throw err;
      }
    }
    return this.data;
  }

  async save() {
    await this._ensureLoaded();
    fs.writeFileSync(this.dbPath, JSON.stringify(this.data, null, 2), {
      encoding: 'utf8',
      mode: 0o600
    });
  }

  async _ensureLoaded() {
    if (!this.data) await this.load();
  }

  /**
   * Derive a token file path from an email address.
   */
  static tokenFileForEmail(email) {
    const safe = email.toLowerCase().replace(/@/g, '-').replace(/\./g, '-');
    return path.join(homeDir, `.outlook-mcp-tokens-${safe}.json`);
  }

  async getAccount(primaryEmail) {
    await this._ensureLoaded();
    return this.data.accounts[primaryEmail.toLowerCase()] || null;
  }

  async addAccount(primaryEmail, opts = {}) {
    await this._ensureLoaded();
    const key = primaryEmail.toLowerCase();
    if (this.data.accounts[key]) return this.data.accounts[key];

    this.data.accounts[key] = {
      label: opts.label || primaryEmail.split('@')[0],
      tokenFile: opts.tokenFile || AccountDB.tokenFileForEmail(primaryEmail),
      aliases: opts.aliases || []
    };

    // First account becomes default
    if (!this.data.defaultAccount) {
      this.data.defaultAccount = key;
    }

    await this.save();
    return this.data.accounts[key];
  }

  async removeAccount(primaryEmail) {
    await this._ensureLoaded();
    const key = primaryEmail.toLowerCase();
    delete this.data.accounts[key];
    if (this.data.defaultAccount === key) {
      const remaining = Object.keys(this.data.accounts);
      this.data.defaultAccount = remaining.length > 0 ? remaining[0] : null;
    }
    await this.save();
  }

  async addAlias(primaryEmail, alias) {
    await this._ensureLoaded();
    const key = primaryEmail.toLowerCase();
    const account = this.data.accounts[key];
    if (!account) throw new Error(`Account ${primaryEmail} not found`);

    const aliasLower = alias.toLowerCase();
    if (!account.aliases.includes(aliasLower)) {
      account.aliases.push(aliasLower);
      await this.save();
    }
    return account;
  }

  async removeAlias(primaryEmail, alias) {
    await this._ensureLoaded();
    const key = primaryEmail.toLowerCase();
    const account = this.data.accounts[key];
    if (!account) throw new Error(`Account ${primaryEmail} not found`);

    account.aliases = account.aliases.filter(a => a !== alias.toLowerCase());
    await this.save();
    return account;
  }

  /**
   * Given any email address (primary or alias), resolve which account owns it.
   * Returns { accountEmail, tokenFile, isAlias } or null.
   */
  async resolveAddress(email) {
    await this._ensureLoaded();
    const emailLower = email.toLowerCase();

    // Check primaries first
    if (this.data.accounts[emailLower]) {
      const acct = this.data.accounts[emailLower];
      return { accountEmail: emailLower, tokenFile: acct.tokenFile, isAlias: false };
    }

    // Check aliases
    for (const [primary, acct] of Object.entries(this.data.accounts)) {
      if (acct.aliases.includes(emailLower)) {
        return { accountEmail: primary, tokenFile: acct.tokenFile, isAlias: true };
      }
    }

    return null;
  }

  async listAllAddresses() {
    await this._ensureLoaded();
    const result = [];
    for (const [primary, acct] of Object.entries(this.data.accounts)) {
      result.push({ email: primary, account: primary, isAlias: false, label: acct.label });
      for (const alias of acct.aliases) {
        result.push({ email: alias, account: primary, isAlias: true, label: acct.label });
      }
    }
    return result;
  }

  async getDefaultAccount() {
    await this._ensureLoaded();
    return this.data.defaultAccount;
  }

  async setDefaultAccount(email) {
    await this._ensureLoaded();
    const key = email.toLowerCase();
    if (!this.data.accounts[key]) throw new Error(`Account ${email} not found`);
    this.data.defaultAccount = key;
    await this.save();
  }

  exists() {
    return fs.existsSync(this.dbPath);
  }
}

// Singleton instance
let _instance = null;
function getAccountDB() {
  if (!_instance) _instance = new AccountDB();
  return _instance;
}

module.exports = { AccountDB, getAccountDB };
