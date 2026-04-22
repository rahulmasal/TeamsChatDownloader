/**
 * Teams Chat Downloader - Content Script
 * Extracts chat messages from the Microsoft Teams web interface.
 *
 * Uses multiple selector strategies to handle Teams UI changes.
 * Supports auto-scrolling to load full chat history.
 */

// Import shared utilities
const {
    SELECTOR_STRATEGIES,
    querySelectorFallback,
    querySelectorAllFallback,
    detectStrategy,
    loadFullChat,
    extractTimestamp,
    sanitizeText
} = window.TeamsChatUtils || {};

/**
 * Scan all visible chat messages using the detected strategy.
 */
function extractMessages(container, strategy) {
    const messages = [];
    const messageElements = querySelectorAllFallback(container, strategy.message);
    const seen = new Set(); // For deduplication

    messageElements.forEach(element => {
        const senderEl = querySelectorFallback(element, strategy.sender);
        const contentEl = querySelectorFallback(element, strategy.content);
        const timestampEl = querySelectorFallback(element, strategy.timestamp);

        const sender = sanitizeText(senderEl?.textContent?.trim() || 'Unknown');
        const text = sanitizeText(contentEl?.textContent?.trim() || '');
        const timestamp = extractTimestamp(timestampEl, strategy);

        // Skip empty messages
        if (!text && !senderEl) return;

        // Deduplication key
        const dedupKey = `${sender}|${timestamp}|${text.substring(0, 50)}`;
        if (seen.has(dedupKey)) return;
        seen.add(dedupKey);

        // Extract attachments
        const attachments = [];
        const attachmentElements = querySelectorAllFallback(element, strategy.attachment);
        attachmentElements.forEach(attachment => {
            const nameEl = querySelectorFallback(attachment, strategy.attachmentName);
            const name = sanitizeText(nameEl?.textContent?.trim());
            if (name) attachments.push(name);
        });

        messages.push({
            sender,
            text,
            timestamp: timestamp || Date.now(),
            attachments
        });
    });

    return messages;
}

/**
 * Main scan function — called by the popup via message passing.
 * @param {boolean} fullHistory - If true, auto-scrolls to load all messages.
 * @param {function} onProgress - Progress callback (0-100).
 */
async function scanChat(fullHistory = false, onProgress = null) {
    // Use the shared detectStrategy function
    const detected = detectStrategy();

    if (!detected) {
        return {
            success: false,
            error: 'Chat container not found. Make sure you have a chat open in Teams.',
            diagnostics: {
                url: window.location.href,
                strategiesTried: SELECTOR_STRATEGIES.map(s => s.name)
            }
        };
    }

    const { strategy, container } = detected;

    // Optionally load full history by auto-scrolling
    if (fullHistory) {
        if (onProgress) onProgress(5);
        await loadFullChat(container, onProgress);
    }

    if (onProgress) onProgress(95);

    const messages = extractMessages(container, strategy);

    if (messages.length === 0) {
        return {
            success: false,
            error: 'No messages found in chat. The chat may be empty or the DOM structure has changed.',
            diagnostics: {
                strategy: strategy.name,
                containerFound: true,
                url: window.location.href
            }
        };
    }

    // Sort by timestamp
    messages.sort((a, b) => a.timestamp - b.timestamp);

    return {
        success: true,
        chatData: messages,
        messageCount: messages.length,
        strategy: strategy.name,
        // Preview: last 10 messages, truncated
        preview: messages.slice(-10).map(msg => ({
            sender: msg.sender,
            text: msg.text.substring(0, 100) + (msg.text.length > 100 ? '...' : ''),
            timestamp: msg.timestamp
        }))
    };
}

// ---- Message listener ----
chrome.runtime.onMessage.addListener(function (request, sender, sendResponse) {
    if (request.action === 'scanChat') {
        const fullHistory = request.fullHistory || false;

        // scanChat is async, so we must return true and call sendResponse later
        scanChat(fullHistory, (progress) => {
            try {
                chrome.runtime.sendMessage({ action: 'scanProgress', progress });
            } catch (e) {
                // Popup may have closed — ignore
            }
        }).then(result => {
            sendResponse(result);
        }).catch(error => {
            sendResponse({
                success: false,
                error: `Scan failed: ${error.message}`
            });
        });

        return true; // Keep the message channel open for async response
    }
});
