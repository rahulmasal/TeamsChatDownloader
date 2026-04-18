/**
 * Authentication module using MSAL (Microsoft Authentication Library).
 * Uses device code flow for CLI — user opens a browser, enters a code, and signs in.
 */

import { PublicClientApplication, CryptoProvider } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TOKEN_CACHE_PATH = path.join(__dirname, '.token-cache.json');

let msalClient = null;

/**
 * Initialize MSAL client with config.
 */
export function initAuth(config) {
    const msalConfig = {
        auth: {
            clientId: config.auth.clientId,
            authority: config.auth.authority,
        },
        cache: {
            // Use a custom cache plugin to persist tokens to disk
        },
    };

    msalClient = new PublicClientApplication(msalConfig);

    // Load cached tokens if available
    if (fs.existsSync(TOKEN_CACHE_PATH)) {
        try {
            const cacheData = fs.readFileSync(TOKEN_CACHE_PATH, 'utf-8');
            msalClient.getTokenCache().deserialize(cacheData);
        } catch (e) {
            // Cache corrupt — ignore, will re-authenticate
        }
    }
}

/**
 * Save the MSAL token cache to disk for reuse.
 */
function saveCache() {
    try {
        const cacheData = msalClient.getTokenCache().serialize();
        fs.writeFileSync(TOKEN_CACHE_PATH, cacheData, 'utf-8');
    } catch (e) {
        // Non-critical
    }
}

/**
 * Get an access token. Tries silent (cached) first, falls back to device code flow.
 */
export async function getAccessToken(scopes) {
    if (!msalClient) {
        throw new Error('Auth not initialized. Call initAuth(config) first.');
    }

    // Try silent acquisition from cache
    const accounts = await msalClient.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
        try {
            const result = await msalClient.acquireTokenSilent({
                account: accounts[0],
                scopes,
            });
            saveCache();
            return result.accessToken;
        } catch (e) {
            // Silent failed — need interactive auth
        }
    }

    // Device code flow
    console.log('\n🔐 Authentication required. Follow the instructions below:\n');
    const result = await msalClient.acquireTokenByDeviceCode({
        scopes,
        deviceCodeCallback: (response) => {
            console.log('━'.repeat(60));
            console.log(`\n  1. Open: ${response.verificationUri}`);
            console.log(`  2. Enter code: ${response.userCode}\n`);
            console.log('━'.repeat(60));
            console.log('\n⏳ Waiting for you to sign in...\n');
        },
    });

    saveCache();
    console.log(`✅ Signed in as: ${result.account.name} (${result.account.username})\n`);
    return result.accessToken;
}

/**
 * Clear the token cache (sign out).
 */
export function clearAuth() {
    if (fs.existsSync(TOKEN_CACHE_PATH)) {
        fs.unlinkSync(TOKEN_CACHE_PATH);
    }
}

/**
 * Get the currently signed-in account info, or null.
 */
export async function getCurrentAccount() {
    if (!msalClient) return null;
    const accounts = await msalClient.getTokenCache().getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
}
