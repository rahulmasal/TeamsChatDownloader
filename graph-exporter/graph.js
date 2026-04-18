/**
 * Microsoft Graph API client for Teams chat operations.
 * Handles pagination, rate limiting, and media downloads.
 */

import 'isomorphic-fetch';
import { Client } from '@microsoft/microsoft-graph-client';
import { getAccessToken } from './auth.js';
import config from './config.js';

let graphClient = null;

/**
 * Initialize the Graph client with token-based auth.
 */
export function initGraphClient() {
    graphClient = Client.init({
        authProvider: async (done) => {
            try {
                const token = await getAccessToken(config.scopes);
                done(null, token);
            } catch (err) {
                done(err, null);
            }
        },
    });
    return graphClient;
}

/**
 * Get the current user's profile.
 */
export async function getMe() {
    return graphClient.api('/me').select('displayName,mail,userPrincipalName').get();
}

/**
 * List all chats for the current user with pagination.
 */
export async function listChats() {
    const allChats = [];
    let response = await graphClient
        .api('/me/chats')
        .expand('members')
        .select('id,topic,chatType,createdDateTime,lastUpdatedDateTime')
        .top(50)
        .orderby('lastUpdatedDateTime desc')
        .get();

    allChats.push(...(response.value || []));

    // Handle pagination
    while (response['@odata.nextLink']) {
        response = await graphClient.api(response['@odata.nextLink']).get();
        allChats.push(...(response.value || []));
    }

    return allChats;
}

/**
 * Get messages from a specific chat with pagination.
 * @param {string} chatId - The chat ID
 * @param {number|null} maxMessages - Max messages to fetch (null = all)
 * @param {function} onProgress - Progress callback (fetched, total)
 */
export async function getChatMessages(chatId, maxMessages = null, onProgress = null) {
    const allMessages = [];
    let response = await graphClient
        .api(`/me/chats/${chatId}/messages`)
        .top(50)
        .orderby('createdDateTime asc')
        .get();

    allMessages.push(...(response.value || []));
    if (onProgress) onProgress(allMessages.length);

    // Handle pagination
    while (response['@odata.nextLink']) {
        if (maxMessages && allMessages.length >= maxMessages) break;

        // Respect rate limits
        await sleep(200);

        try {
            response = await graphClient.api(response['@odata.nextLink']).get();
            allMessages.push(...(response.value || []));
            if (onProgress) onProgress(allMessages.length);
        } catch (err) {
            if (err.statusCode === 429) {
                // Rate limited — wait and retry
                const retryAfter = parseInt(err.headers?.['retry-after'] || '5', 10);
                console.log(`  ⏳ Rate limited, waiting ${retryAfter}s...`);
                await sleep(retryAfter * 1000);
                response = await graphClient.api(response['@odata.nextLink']).get();
                allMessages.push(...(response.value || []));
            } else {
                throw err;
            }
        }
    }

    if (maxMessages) {
        return allMessages.slice(0, maxMessages);
    }
    return allMessages;
}

/**
 * Download hosted content (inline images) from a message.
 * @param {string} chatId
 * @param {string} messageId
 * @param {string} hostedContentId
 * @returns {Buffer} - The binary content
 */
export async function downloadHostedContent(chatId, messageId, hostedContentId) {
    const token = await getAccessToken(config.scopes);
    const url = `https://graph.microsoft.com/v1.0/chats/${chatId}/messages/${messageId}/hostedContents/${hostedContentId}/$value`;

    const response = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
        throw new Error(`Failed to download hosted content: ${response.status} ${response.statusText}`);
    }

    const buffer = await response.arrayBuffer();
    return Buffer.from(buffer);
}

/**
 * Download a file attachment from OneDrive/SharePoint.
 * @param {string} driveItemUrl - The driveItem download URL
 * @returns {Buffer}
 */
export async function downloadDriveItem(driveItemUrl) {
    const token = await getAccessToken(config.scopes);

    const response = await fetch(driveItemUrl, {
        headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
        throw new Error(`Failed to download file: ${response.status}`);
    }

    const buffer = await response.arrayBuffer();
    return Buffer.from(buffer);
}

/**
 * Get a DriveItem download URL from Graph.
 * @param {string} driveId
 * @param {string} itemId
 * @returns {string} - The @microsoft.graph.downloadUrl
 */
export async function getDriveItemDownloadUrl(driveId, itemId) {
    try {
        const item = await graphClient
            .api(`/drives/${driveId}/items/${itemId}`)
            .select('@microsoft.graph.downloadUrl,name,size')
            .get();
        return item;
    } catch (err) {
        return null;
    }
}

/**
 * Get user profile photo as a Buffer.
 */
export async function getUserPhoto(userId) {
    try {
        const token = await getAccessToken(config.scopes);
        const response = await fetch(
            `https://graph.microsoft.com/v1.0/users/${userId}/photo/$value`,
            { headers: { Authorization: `Bearer ${token}` } }
        );
        if (!response.ok) return null;
        return Buffer.from(await response.arrayBuffer());
    } catch {
        return null;
    }
}

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}
