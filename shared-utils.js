/**
 * Teams Chat Downloader - Shared Utilities
 * Contains common functions used by both the Chrome extension and Tampermonkey script
 */

/**
 * Selector strategies — ordered by priority.
 * Teams frequently changes its DOM; having fallbacks makes the extension more resilient.
 */
let SELECTOR_STRATEGIES = [
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
        attachmentName: '.ts-attachment-name, .file-name',
    }
];

// Asynchronously fetch latest selectors from GitHub (failsafe back to hardcoded)
(async function initSelectors() {
    try {
        const res = await fetch('https://raw.githubusercontent.com/rahulmasal/TeamsChatDownloader/main/selectors.json');
        if (res.ok) {
            SELECTOR_STRATEGIES = await res.json();
        }
    } catch (e) {
        // use fallback hardcoded strategies
    }
})();

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
async function loadFullChat(container, onProgress, maxAttempts = 50, scrollDelay = 1200) {
    let previousHeight = 0;
    let attempts = 0;

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

        await new Promise(resolve => setTimeout(resolve, scrollDelay));
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
 * Sanitize text content to prevent XSS vulnerabilities
 */
function sanitizeText(text) {
    if (!text) return '';

    // Basic HTML entity encoding to prevent XSS
    let sanitized = text;
    sanitized = sanitized.replace(/&/g, '&');
    sanitized = sanitized.replace(/</g, '<');
    sanitized = sanitized.replace(/>/g, '>');
    sanitized = sanitized.replace(/"/g, '"');
    sanitized = sanitized.replace(/'/g, '&#x27;');
    return sanitized;
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

// Export for different environments
if (typeof module !== 'undefined' && module.exports) {
    // Node.js environment
    module.exports = {
        SELECTOR_STRATEGIES,
        querySelectorFallback,
        querySelectorAllFallback,
        detectStrategy,
        loadFullChat,
        extractTimestamp,
        extractMessages,
        sanitizeText
    };
} else if (typeof window !== 'undefined') {
    // Browser environment
    window.TeamsChatUtils = {
        SELECTOR_STRATEGIES,
        querySelectorFallback,
        querySelectorAllFallback,
        detectStrategy,
        loadFullChat,
        extractTimestamp,
        extractMessages,
        sanitizeText
    };
}