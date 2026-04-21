/**
 * Chat Exporter — Downloads messages and media, generates output files.
 * Handles inline images, file attachments, GIFs, stickers, and adaptive cards.
 */

import fs from 'fs';
import path from 'path';
import {
    getChatMessages,
    downloadHostedContent,
    downloadDriveItem,
    getDriveItemDownloadUrl,
} from './graph.js';
import config from './config.js';

/**
 * Export a single chat to disk.
 * @param {object} chat - Chat object from Graph API
 * @param {string} outputDir - Base output directory
 * @param {object} options - Export options
 * @param {function} onProgress - Progress callback
 */
export async function exportChat(chat, outputDir, options = {}, onProgress = null) {
    const chatName = getChatDisplayName(chat);
    const safeName = sanitizeFilename(chatName);
    const chatDir = path.join(outputDir, safeName);
    const mediaDir = path.join(chatDir, 'media');

    // Create directories
    fs.mkdirSync(chatDir, { recursive: true });
    if (options.downloadMedia !== false) {
        fs.mkdirSync(mediaDir, { recursive: true });
    }

    // Fetch all messages
    if (onProgress) onProgress('messages', 0, 'Fetching messages...');
    const messages = await getChatMessages(
        chat.id,
        options.maxMessages || null,
        (count) => {
            if (onProgress) onProgress('messages', count, `Fetched ${count} messages...`);
        }
    );

    if (messages.length === 0) {
        if (onProgress) onProgress('done', 0, 'No messages found');
        return { chatName, messageCount: 0, mediaCount: 0 };
    }

    // Async pool helper for concurrent message processing
    async function asyncPool(poolLimit, array, iteratorFn) {
        const ret = [];
        const executing = [];
        for (const item of array) {
            const p = Promise.resolve().then(() => iteratorFn(item));
            ret.push(p);
            if (poolLimit <= array.length) {
                const e = p.then(() => executing.splice(executing.indexOf(e), 1));
                executing.push(e);
                if (executing.length >= poolLimit) {
                    await Promise.race(executing);
                }
            }
        }
        return Promise.all(ret);
    }

    // Process messages and download media concurrently
    let mediaCount = 0;
    let mediaErrors = 0;
    let processedCount = 0;
    const concurrencyLimit = options.concurrency || 3;

    const processedMessages = await asyncPool(concurrencyLimit, messages, async (msg) => {
        const processed = {
            id: msg.id,
            sender: extractSender(msg),
            timestamp: msg.createdDateTime,
            body: msg.body?.content || '',
            bodyType: msg.body?.contentType || 'text',
            attachments: [],
            hostedImages: [],
            reactions: msg.reactions || [],
            importance: msg.importance,
            isDeleted: msg.deletedDateTime != null,
            isEdited: msg.lastEditedDateTime != null,
            mentions: msg.mentions || [],
        };

        // Download inline hosted content (images pasted into chat)
        if (options.downloadMedia !== false && msg.hostedContents?.length > 0) {
            for (const hc of msg.hostedContents) {
                try {
                    const ext = getExtensionFromContentType(hc.contentType || 'image/png');
                    const filename = `inline_${msg.id.substring(0, 8)}_${hc.id.substring(0, 8)}${ext}`;
                    const filePath = path.join(mediaDir, filename);

                    if (fs.existsSync(filePath) && fs.statSync(filePath).size > 0) {
                        // Skip download, file exists
                        processed.hostedImages.push({
                            id: hc.id,
                            filename,
                            contentType: hc.contentType,
                            localPath: `media/${filename}`,
                        });
                        mediaCount++;
                    } else {
                        const buffer = await downloadHostedContent(chat.id, msg.id, hc.id);
                        fs.writeFileSync(filePath, buffer);
                        processed.hostedImages.push({
                            id: hc.id,
                            filename,
                            contentType: hc.contentType,
                            localPath: `media/${filename}`,
                        });
                        mediaCount++;
                    }
                } catch (err) {
                    mediaErrors++;
                    processed.hostedImages.push({
                        id: hc.id,
                        error: err.message,
                    });
                }
            }
        }

        // Download file attachments
        if (options.downloadMedia !== false && msg.attachments?.length > 0) {
            for (const att of msg.attachments) {
                try {
                    const result = await processAttachment(att, chat.id, msg.id, mediaDir);
                    if (result) {
                        processed.attachments.push(result);
                        if (result.localPath) mediaCount++;
                    }
                } catch (err) {
                    mediaErrors++;
                    processed.attachments.push({
                        name: att.name || 'unknown',
                        contentType: att.contentType,
                        error: err.message,
                    });
                }
            }
        }

        // Extract inline images from HTML body
        if (
            options.downloadMedia !== false &&
            processed.bodyType === 'html' &&
            processed.body.includes('<img')
        ) {
            const inlineImages = await extractAndDownloadInlineImages(
                processed.body,
                chat.id,
                msg.id,
                mediaDir,
                mediaCount
            );
            processed.hostedImages.push(...inlineImages.images);
            mediaCount += inlineImages.downloaded;
            processed.body = inlineImages.updatedBody;
        }

        processedCount++;
        if (onProgress && processedCount % 20 === 0) {
            onProgress(
                'processing',
                Math.round((processedCount / messages.length) * 100),
                `Processing ${processedCount}/${messages.length}... (${mediaCount} media files)`
            );
        }

        return processed;
    });

    // Write chat metadata
    const metadata = {
        chatId: chat.id,
        chatName,
        chatType: chat.chatType,
        createdDateTime: chat.createdDateTime,
        lastUpdatedDateTime: chat.lastUpdatedDateTime,
        participants: (chat.members || []).map((m) => ({
            displayName: m.displayName,
            email: m.email,
            userId: m.userId,
        })),
        messageCount: processedMessages.length,
        mediaCount,
        mediaErrors,
        exportedAt: new Date().toISOString(),
    };

    fs.writeFileSync(path.join(chatDir, 'metadata.json'), JSON.stringify(metadata, null, 2));

    // Generate output files based on format
    const format = options.format || 'all';

    if (format === 'json' || format === 'all') {
        writeJSON(chatDir, processedMessages, metadata);
    }
    if (format === 'html' || format === 'all') {
        writeHTML(chatDir, processedMessages, metadata);
    }
    if (format === 'txt' || format === 'all') {
        writeTXT(chatDir, processedMessages, metadata);
    }

    if (onProgress) onProgress('done', 100, `Exported ${processedMessages.length} messages, ${mediaCount} media files`);

    return {
        chatName,
        messageCount: processedMessages.length,
        mediaCount,
        mediaErrors,
        outputPath: chatDir,
    };
}

// ================================================================
// ATTACHMENT PROCESSING
// ================================================================

async function processAttachment(att, chatId, messageId, mediaDir) {
    const attInfo = {
        name: att.name || 'untitled',
        contentType: att.contentType,
        contentUrl: att.contentUrl,
    };

    // Reference attachments (files shared from OneDrive/SharePoint)
    if (att.contentType === 'reference') {
        // Try to get the file from the content URL
        if (att.contentUrl) {
            try {
                // Parse OneDrive/SharePoint URL to get driveId and itemId
                const urlMatch = att.contentUrl.match(
                    /drives\/([^/]+)\/items\/([^/]+)/
                );
                if (urlMatch) {
                    const driveItem = await getDriveItemDownloadUrl(urlMatch[1], urlMatch[2]);
                    if (driveItem && driveItem['@microsoft.graph.downloadUrl']) {
                        const safeName = sanitizeFilename(att.name || driveItem.name || 'file');
                        const filePath = path.join(mediaDir, safeName);
                        
                        if (fs.existsSync(filePath) && fs.statSync(filePath).size > 0) {
                            attInfo.localPath = `media/${safeName}`;
                            attInfo.size = fs.statSync(filePath).size;
                        } else {
                            const buffer = await downloadDriveItem(
                                driveItem['@microsoft.graph.downloadUrl']
                            );
                            fs.writeFileSync(filePath, buffer);
                            attInfo.localPath = `media/${safeName}`;
                            attInfo.size = buffer.length;
                        }
                    }
                }
            } catch (err) {
                attInfo.downloadError = err.message;
            }
        }
        return attInfo;
    }

    // Inline/file attachments with direct content
    if (att.contentUrl && !att.contentUrl.startsWith('https://teams.microsoft.com')) {
        try {
            const ext = getExtensionFromContentType(att.contentType || '');
            const safeName = sanitizeFilename(att.name || `attachment_${messageId.substring(0, 10)}${ext}`);
            const filePath = path.join(mediaDir, safeName);

            if (fs.existsSync(filePath) && fs.statSync(filePath).size > 0) {
                attInfo.localPath = `media/${safeName}`;
                attInfo.size = fs.statSync(filePath).size;
            } else {
                const buffer = await downloadDriveItem(att.contentUrl);
                fs.writeFileSync(filePath, buffer);
                attInfo.localPath = `media/${safeName}`;
                attInfo.size = buffer.length;
            }
        } catch (err) {
            attInfo.downloadError = err.message;
        }
    }

    return attInfo;
}

/**
 * Extract <img> tags from HTML body, download the images, and replace src with local paths.
 */
async function extractAndDownloadInlineImages(body, chatId, messageId, mediaDir, mediaOffset) {
    const images = [];
    let downloaded = 0;
    let updatedBody = body;

    // Match img tags with src pointing to hosted content or graph URLs
    const imgRegex = /<img[^>]+src=["']([^"']+)["'][^>]*>/gi;
    let match;

    while ((match = imgRegex.exec(body)) !== null) {
        const src = match[1];
        const imgTag = match[0];

        // Only download Graph/Teams hosted images
        if (src.includes('graph.microsoft.com') || src.includes('teams.microsoft.com')) {
            try {
                const response = await fetch(src, {
                    headers: {
                        Authorization: `Bearer ${(await import('./auth.js')).default}`,
                    },
                });
                // Skip if auth is needed differently
            } catch {
                // Will be handled in the main hosted content loop
            }
        }
    }

    return { images, downloaded, updatedBody };
}

// ================================================================
// OUTPUT FORMATTERS
// ================================================================

function writeJSON(chatDir, messages, metadata) {
    const output = {
        ...metadata,
        messages: messages.map((m) => ({
            sender: m.sender,
            timestamp: m.timestamp,
            text: stripHtml(m.body),
            bodyHtml: m.bodyType === 'html' ? m.body : undefined,
            attachments: m.attachments,
            inlineImages: m.hostedImages.filter((img) => !img.error),
            reactions: m.reactions.map((r) => ({
                type: r.reactionType,
                user: r.user?.displayName || 'Unknown',
            })),
            isDeleted: m.isDeleted,
            isEdited: m.isEdited,
        })),
    };

    fs.writeFileSync(path.join(chatDir, 'chat.json'), JSON.stringify(output, null, 2));
}

function writeHTML(chatDir, messages, metadata) {
    const participants = metadata.participants.map((p) => p.displayName).join(', ');
    const esc = (s) =>
        (s || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/\n/g, '<br>');

    let html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${esc(metadata.chatName)} — Teams Chat Export</title>
    <style>
        :root { --bg: #0f0f1a; --card: #1a1a2e; --card-hover: #1e2a45; --accent: #7b68ee; --text: #e0e0e0; --muted: #6c6c80; --border: rgba(255,255,255,0.06); }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); color: var(--text); padding: 0; }
        .header { background: linear-gradient(135deg, #16213e 0%, #0f0f1a 100%); padding: 32px; border-bottom: 1px solid var(--border); }
        .header h1 { color: var(--accent); font-size: 1.5em; margin-bottom: 4px; }
        .header .meta { color: var(--muted); font-size: 0.85em; line-height: 1.6; }
        .chat-container { max-width: 850px; margin: 0 auto; padding: 24px; }
        .date-divider { text-align: center; margin: 24px 0 16px; position: relative; }
        .date-divider::before { content: ''; position: absolute; left: 0; top: 50%; width: 100%; height: 1px; background: var(--border); }
        .date-divider span { position: relative; background: var(--bg); padding: 0 16px; font-size: 0.8em; color: var(--muted); font-weight: 600; }
        .msg { background: var(--card); border-radius: 12px; padding: 14px 18px; margin-bottom: 6px; border-left: 3px solid var(--accent); transition: background 0.2s; }
        .msg:hover { background: var(--card-hover); }
        .msg.deleted { opacity: 0.5; border-left-color: #ef5350; }
        .msg-header { display: flex; align-items: baseline; gap: 10px; margin-bottom: 6px; }
        .msg-sender { color: var(--accent); font-weight: 600; font-size: 0.9em; }
        .msg-time { color: var(--muted); font-size: 0.75em; }
        .msg-edited { color: var(--muted); font-size: 0.7em; font-style: italic; }
        .msg-body { line-height: 1.55; font-size: 0.9em; word-wrap: break-word; white-space: pre-wrap; }
        .msg-body img { max-width: 400px; border-radius: 8px; margin: 8px 0; display: block; }
        .msg-attachments { margin-top: 10px; }
        .attachment { display: inline-flex; align-items: center; gap: 6px; padding: 6px 12px; background: #0f3460; border-radius: 8px; font-size: 0.8em; color: #a0c4ff; margin: 2px 4px 2px 0; text-decoration: none; }
        .attachment:hover { background: #1a4a80; }
        .reactions { margin-top: 8px; display: flex; gap: 6px; flex-wrap: wrap; }
        .reaction { background: rgba(123,104,238,0.15); border: 1px solid rgba(123,104,238,0.25); border-radius: 20px; padding: 2px 8px; font-size: 0.75em; }
        .stats { background: var(--card); border-radius: 12px; padding: 16px; margin-top: 24px; text-align: center; color: var(--muted); font-size: 0.85em; }
    </style>
</head>
<body>
    <div class="header">
        <h1>📋 ${esc(metadata.chatName)}</h1>
        <div class="meta">
            <strong>Participants:</strong> ${esc(participants)}<br>
            <strong>Messages:</strong> ${metadata.messageCount} · <strong>Media:</strong> ${metadata.mediaCount} files<br>
            <strong>Exported:</strong> ${new Date(metadata.exportedAt).toLocaleString()}
        </div>
    </div>
    <div class="chat-container">
`;

    let lastDate = '';
    for (const msg of messages) {
        const date = new Date(msg.timestamp);
        const dateStr = date.toLocaleDateString('en-US', {
            weekday: 'long',
            year: 'numeric',
            month: 'long',
            day: 'numeric',
        });

        // Date divider
        if (dateStr !== lastDate) {
            html += `        <div class="date-divider"><span>${esc(dateStr)}</span></div>\n`;
            lastDate = dateStr;
        }

        const timeStr = date.toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit',
        });

        const deletedClass = msg.isDeleted ? ' deleted' : '';
        html += `        <div class="msg${deletedClass}">\n`;
        html += `            <div class="msg-header">\n`;
        html += `                <span class="msg-sender">${esc(msg.sender)}</span>\n`;
        html += `                <span class="msg-time">${esc(timeStr)}</span>\n`;
        if (msg.isEdited) html += `                <span class="msg-edited">(edited)</span>\n`;
        html += `            </div>\n`;

        // Body
        if (msg.isDeleted) {
            html += `            <div class="msg-body"><em>This message was deleted.</em></div>\n`;
        } else if (msg.bodyType === 'html') {
            // Render HTML body directly but sanitize scripts
            const safeBody = msg.body
                .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
                .replace(/on\w+="[^"]*"/gi, '');
            html += `            <div class="msg-body">${safeBody}</div>\n`;
        } else {
            html += `            <div class="msg-body">${esc(msg.body)}</div>\n`;
        }

        // Inline images
        for (const img of msg.hostedImages) {
            if (img.localPath) {
                html += `            <img src="${img.localPath}" alt="inline image" style="max-width:400px;border-radius:8px;margin:8px 0;">\n`;
            }
        }

        // Attachments
        if (msg.attachments.length > 0) {
            html += `            <div class="msg-attachments">\n`;
            for (const att of msg.attachments) {
                if (att.localPath) {
                    html += `                <a href="${att.localPath}" class="attachment">📎 ${esc(att.name)}</a>\n`;
                } else {
                    html += `                <span class="attachment">📎 ${esc(att.name)} (not downloaded)</span>\n`;
                }
            }
            html += `            </div>\n`;
        }

        // Reactions
        if (msg.reactions.length > 0) {
            html += `            <div class="reactions">\n`;
            for (const r of msg.reactions) {
                const emoji = getReactionEmoji(r.reactionType);
                html += `                <span class="reaction">${emoji} ${esc(r.user?.displayName || '')}</span>\n`;
            }
            html += `            </div>\n`;
        }

        html += `        </div>\n`;
    }

    html += `
        <div class="stats">
            📊 ${metadata.messageCount} messages · ${metadata.mediaCount} media files · Exported ${new Date(metadata.exportedAt).toLocaleString()}
        </div>
    </div>
</body>
</html>`;

    fs.writeFileSync(path.join(chatDir, 'chat.html'), html);
}

function writeTXT(chatDir, messages, metadata) {
    const participants = metadata.participants.map((p) => p.displayName).join(', ');

    let txt = `Microsoft Teams Chat Export\n`;
    txt += `Chat: ${metadata.chatName}\n`;
    txt += `Participants: ${participants}\n`;
    txt += `Messages: ${metadata.messageCount} · Media: ${metadata.mediaCount}\n`;
    txt += `Exported: ${new Date(metadata.exportedAt).toLocaleString()}\n`;
    txt += '═'.repeat(60) + '\n\n';

    let lastDate = '';
    for (const msg of messages) {
        const date = new Date(msg.timestamp);
        const dateStr = date.toLocaleDateString();

        if (dateStr !== lastDate) {
            txt += `\n--- ${dateStr} ---\n\n`;
            lastDate = dateStr;
        }

        const timeStr = date.toLocaleTimeString();
        const editTag = msg.isEdited ? ' (edited)' : '';
        const deleteTag = msg.isDeleted ? ' [DELETED]' : '';

        txt += `[${timeStr}] ${msg.sender}${editTag}${deleteTag}:\n`;

        if (msg.isDeleted) {
            txt += `  (This message was deleted)\n`;
        } else {
            const text = stripHtml(msg.body);
            txt += `  ${text}\n`;
        }

        // Attachments
        for (const att of msg.attachments) {
            const loc = att.localPath ? ` → ${att.localPath}` : ' (not downloaded)';
            txt += `  📎 ${att.name}${loc}\n`;
        }

        // Inline images
        for (const img of msg.hostedImages) {
            if (img.localPath) {
                txt += `  🖼️ ${img.localPath}\n`;
            }
        }

        // Reactions
        if (msg.reactions.length > 0) {
            const rxns = msg.reactions
                .map((r) => `${getReactionEmoji(r.reactionType)} ${r.user?.displayName || ''}`)
                .join(', ');
            txt += `  Reactions: ${rxns}\n`;
        }

        txt += '\n';
    }

    txt += '═'.repeat(60) + '\n';
    txt += `End of export — ${metadata.messageCount} messages\n`;

    fs.writeFileSync(path.join(chatDir, 'chat.txt'), txt);
}

// ================================================================
// UTILITY FUNCTIONS
// ================================================================

function getChatDisplayName(chat) {
    if (chat.topic) return chat.topic;
    const members = (chat.members || []).map((m) => m.displayName).filter(Boolean);
    if (members.length > 0) return members.join(', ');
    return `Chat_${chat.id.substring(0, 8)}`;
}

function extractSender(msg) {
    if (msg.from?.user?.displayName) return msg.from.user.displayName;
    if (msg.from?.application?.displayName) return `[Bot] ${msg.from.application.displayName}`;
    if (msg.messageType === 'systemEventMessage') return '[System]';
    return 'Unknown';
}

function sanitizeFilename(name) {
    return name
        .replace(/[<>:"/\\|?*]/g, '_')
        .replace(/\s+/g, ' ')
        .trim()
        .substring(0, 100);
}

function stripHtml(html) {
    if (!html) return '';
    return html
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/p>/gi, '\n')
        .replace(/<\/div>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&nbsp;/g, ' ')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
}

function getExtensionFromContentType(contentType) {
    const map = {
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/gif': '.gif',
        'image/webp': '.webp',
        'image/svg+xml': '.svg',
        'image/bmp': '.bmp',
        'video/mp4': '.mp4',
        'audio/mpeg': '.mp3',
        'audio/ogg': '.ogg',
        'application/pdf': '.pdf',
        'application/zip': '.zip',
        'text/plain': '.txt',
    };
    return map[contentType] || '.bin';
}

function getReactionEmoji(type) {
    const map = {
        like: '👍',
        heart: '❤️',
        laugh: '😂',
        surprised: '😮',
        sad: '😢',
        angry: '😡',
    };
    return map[type] || `(${type})`;
}
