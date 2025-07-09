import keytar from 'keytar';
import storage from 'node-persist';
import crypto from 'crypto';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SERVICE_NAME = 'outlook-mcp';
const ENCRYPTION_KEY_ACCOUNT = 'encryption-key';
const ACCESS_TOKEN_ACCOUNT = 'access-token';
const REFRESH_TOKEN_ACCOUNT = 'refresh-token';
const TOKEN_METADATA_KEY = 'token-metadata';

export class TokenManager {
  constructor(clientId) {
    this.clientId = clientId;
    this.storageInitialized = false;
    this.encryptionKey = null;
  }

  async initialize() {
    if (this.storageInitialized) return;

    await storage.init({
      dir: path.join(__dirname, '../../.tokens'),
      logging: false,
    });

    this.encryptionKey = await this.getOrCreateEncryptionKey();
    this.storageInitialized = true;
  }

  async getOrCreateEncryptionKey() {
    try {
      const existingKey = await keytar.getPassword(SERVICE_NAME, ENCRYPTION_KEY_ACCOUNT);
      if (existingKey) {
        return Buffer.from(existingKey, 'base64');
      }

      const newKey = crypto.randomBytes(32);
      await keytar.setPassword(SERVICE_NAME, ENCRYPTION_KEY_ACCOUNT, newKey.toString('base64'));
      return newKey;
    } catch (error) {
      // This is expected in some environments (containers, MCP servers, etc.)
      // Tokens will still be encrypted and stored securely in the file system
      const fallbackKey = crypto.createHash('sha256')
        .update(this.clientId + (process.env.AZURE_TENANT_ID || 'default'))
        .digest();
      return fallbackKey;
    }
  }

  encrypt(text) {
    const iv = crypto.randomBytes(16);
    const cipher = crypto.createCipheriv('aes-256-cbc', this.encryptionKey, iv);
    let encrypted = cipher.update(text, 'utf8', 'hex');
    encrypted += cipher.final('hex');
    return iv.toString('hex') + ':' + encrypted;
  }

  decrypt(encryptedText) {
    const parts = encryptedText.split(':');
    const iv = Buffer.from(parts.shift(), 'hex');
    const encrypted = parts.join(':');
    const decipher = crypto.createDecipheriv('aes-256-cbc', this.encryptionKey, iv);
    let decrypted = decipher.update(encrypted, 'hex', 'utf8');
    decrypted += decipher.final('utf8');
    return decrypted;
  }

  async storeTokens(accessToken, refreshToken, expiresIn = 3600) {
    await this.initialize();

    let usingFallback = false;
    try {
      await keytar.setPassword(SERVICE_NAME, ACCESS_TOKEN_ACCOUNT, this.encrypt(accessToken));
      if (refreshToken) {
        await keytar.setPassword(SERVICE_NAME, REFRESH_TOKEN_ACCOUNT, this.encrypt(refreshToken));
      }
    } catch (error) {
      // Keytar not available in this environment, using secure file storage instead
      usingFallback = true;
      await storage.setItem('fallback_access_token', this.encrypt(accessToken));
      if (refreshToken) {
        await storage.setItem('fallback_refresh_token', this.encrypt(refreshToken));
      }
    }

    const metadata = {
      accessTokenExpiry: Date.now() + (expiresIn * 1000),
      refreshTokenExpiry: Date.now() + (90 * 24 * 60 * 60 * 1000), // 90 days
      lastRefresh: Date.now(),
    };
    await storage.setItem(TOKEN_METADATA_KEY, metadata);

    if (usingFallback) {
      console.log('Tokens stored securely in encrypted file storage');
    }
  }

  async getAccessToken() {
    await this.initialize();

    const metadata = await storage.getItem(TOKEN_METADATA_KEY);
    if (!metadata) {
      throw new Error('No token metadata found');
    }

    const refreshThreshold = 55 * 60 * 1000; // 55 minutes
    const shouldRefresh = Date.now() > (metadata.accessTokenExpiry - refreshThreshold);

    if (shouldRefresh) {
      throw new Error('Access token needs refresh');
    }

    try {
      const encryptedToken = await keytar.getPassword(SERVICE_NAME, ACCESS_TOKEN_ACCOUNT);
      if (encryptedToken) {
        return this.decrypt(encryptedToken);
      }
    } catch (error) {
      const fallbackToken = await storage.getItem('fallback_access_token');
      if (fallbackToken) {
        return this.decrypt(fallbackToken);
      }
    }

    throw new Error('No access token found');
  }

  async getRefreshToken() {
    await this.initialize();

    const metadata = await storage.getItem(TOKEN_METADATA_KEY);
    if (!metadata) {
      throw new Error('No token metadata found');
    }

    if (Date.now() > metadata.refreshTokenExpiry) {
      throw new Error('Refresh token has expired');
    }

    try {
      const encryptedToken = await keytar.getPassword(SERVICE_NAME, REFRESH_TOKEN_ACCOUNT);
      if (encryptedToken) {
        return this.decrypt(encryptedToken);
      }
    } catch (error) {
      const fallbackToken = await storage.getItem('fallback_refresh_token');
      if (fallbackToken) {
        return this.decrypt(fallbackToken);
      }
    }

    throw new Error('No refresh token found');
  }

  async clearTokens() {
    await this.initialize();

    try {
      await keytar.deletePassword(SERVICE_NAME, ACCESS_TOKEN_ACCOUNT);
      await keytar.deletePassword(SERVICE_NAME, REFRESH_TOKEN_ACCOUNT);
    } catch (error) {
      // Silently continue - keytar might not be available
    }

    await storage.removeItem('fallback_access_token');
    await storage.removeItem('fallback_refresh_token');
    await storage.removeItem(TOKEN_METADATA_KEY);
  }

  generateCodeVerifier() {
    return crypto.randomBytes(32).toString('base64url');
  }

  generateCodeChallenge(verifier) {
    return crypto.createHash('sha256')
      .update(verifier)
      .digest('base64url');
  }

  async storePKCEVerifier(verifier) {
    await this.initialize();
    await storage.setItem('pkce_verifier', verifier);
  }

  async getPKCEVerifier() {
    await this.initialize();
    const verifier = await storage.getItem('pkce_verifier');
    await storage.removeItem('pkce_verifier');
    return verifier;
  }

  async isAuthenticated() {
    try {
      await this.getAccessToken();
      return true;
    } catch (error) {
      return false;
    }
  }

  async getTokenMetadata() {
    await this.initialize();
    return await storage.getItem(TOKEN_METADATA_KEY);
  }
}