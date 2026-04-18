/**
 * Teams Chat Downloader - Content Script
 * Extracts chat messages from the Microsoft Teams web interface.
 *
 * Uses multiple selector strategies to handle Teams UI changes.
 * Supports auto-scrolling to load full chat history.
 */

// Selector strategies — ordered by priority.
// Teams frequently changes its DOM; having fallbacks makes the extension more resilient.
const SELECTOR_STRATEGIES = [
    {
        name: 'data-tid',
        chatContainer: '[data-tid="chat-pane-list"], [data-tid="chat-list"], [data-tid="message-pane-list"]',
        message: '[data-tid="chat-pane-message"], [data-tid="message"], [data-tid="message-body"]',
        sender: '[data-tid="message-author-name"], [data-tid="message-sender"]',
        content: '[data-tid="message-text"], [data-tid="message-content"]',
        timestamp: '[data-tid="message-timestamp"], [data-tid="ts-message-timestamp"]',
        attachment: '[data-tid="file-card"], [data-tid="message-attachment"]',
        attachmentName: '[data-tid="file-card-name"], [data-tid="attachment-name"]'
    },
    {
        name: 'aria-role',
        chatContainer: '[role="main"] [role="list"], [role="log"]',
        message: '[role="listitem"], [data-is-focusable="true"]',
        sender: '[data-testid="message-author"], .ui-chat__message__author',
        content: '[data-testid="message-body"], .ui-chat__message__content',
        timestamp: 'time[datetime], [data-testid="message-timestamp"]',
        attachment: '.ui-attachment, [data-testid="file-attachment"]',
        attachmentName: '.ui-attachment__header, [data-testid="file-name"]'
    },
    {
        name: 'class-based',
        chatContainer: '.ts-message-list-container, .message-list',
        message: '.ts-message, .message-body-content',
        sender: '.ts-message-sender, .message-sender-name',
        content: '.ts-message-text, .message-body-text',
        timestamp: '.ts-message-timestamp, time',
        attachment: '.ts-attachment, .file-attachment',
        attachmentName: '.ts-attachment-name, .file-name'
    }
];

/**
 * Find the first matching element using multiple selectors from a strategy.
 */
function querySelectorFallback(root, selectorString) {
    const selectors = selectorString.split(',').map(s => s.trim());
    for (const selector of selectors) {
        const el = root.querySelector(selector);
        if (el) return el;
    }
    return null;
}

function querySelectorAllFallback(root, selectorString) {
    const selectors = selectorString.split(',').map(s => s.trim());
    for (const selector of selectors) {
        const elements = root.querySelectorAll(selector);
        if (elements.length > 0) return elements;
    }
    return [];
}

/**
 * Detect which selector strategy works for the current Teams DOM.
 */
function detectStrategy() {
    for (const strategy of SELECTOR_STRATEGIES) {
        const container = querySelectorFallback(document, strategy.chatContainer);
        if (container) {
            const messages = querySelectorAllFallback(container, strategy.message);
            if (messages.length > 0) {
                console.log(`[TeamsChatDownloader] Using selector strategy: ${strategy.name}`);
                return { strategy, container };
            }
        }
    }
    return null;
}

/**
 * Auto-scroll the chat container to load older messages.
 * Teams uses virtual scrolling, so only visible messages are in the DOM.
 */
async function loadFullChat(container, onProgress) {
    let previousHeight = 0;
    let attempts = 0;
    const maxAttempts = 50; // Safety limit to prevent infinite scrolling

    while (attempts < maxAttempts) {
        const currentHeight = container.scrollHeight;

        if (currentHeight === previousHeight && attempts > 2) {
            // No new content loaded after scrolling — we've reached the top
            break;
        }

        previousHeight = currentHeight;
        container.scrollTop = 0; // Scroll to top to trigger lazy loading

        const progress = Math.min(90, Math.round((attempts / maxAttempts) * 90));
        if (onProgress) onProgress(progress);

        await new Promise(resolve => setTimeout(resolve, 1200));
        attempts++;
    }

    // Scroll back to bottom so the user sees the latest messages
    container.scrollTop = container.scrollHeight;
}

/**
 * Extract a timestamp value from a timestamp element.
 * Handles both data-timestamp attributes (Unix ms or s) and datetime attributes.
 */
function extractTimestamp(element, strategy) {
    if (!element) return null;

    // Try data-timestamp attribute first
    const dataTs = element.getAttribute('data-timestamp');
    if (dataTs) {
        let ts = parseInt(dataTs, 10);
        if (!isNaN(ts)) {
            // If it looks like seconds instead of milliseconds, convert
            if (ts < 1e12) ts *= 1000;
            return ts;
        }
    }

    // Try datetime attribute (ISO 8601)
    const datetime = element.getAttribute('datetime');
    if (datetime) {
        const d = new Date(datetime);
        if (!isNaN(d.getTime())) return d.getTime();
    }

    // Try title attribute as a displayable date
    const title = element.getAttribute('title');
    if (title) {
        const d = new Date(title);
        if (!isNaN(d.getTime())) return d.getTime();
    }

    // Fallback: use textContent as a date string
    const text = element.textContent?.trim();
    if (text) {
        const d = new Date(text);
        if (!isNaN(d.getTime())) return d.getTime();
    }

    return Date.now(); // Last resort fallback
}

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

        const sender = senderEl?.textContent?.trim() || 'Unknown';
        const text = contentEl?.textContent?.trim() || '';
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
            const name = nameEl?.textContent?.trim();
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