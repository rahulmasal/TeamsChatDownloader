// ==UserScript==
// @name         Teams Chat Downloader
// @namespace    https://github.com/rahulmasal/TeamsChatDownloader
// @version      1.1.1
// @description  Download Microsoft Teams chat history with one click. Supports TXT, JSON, CSV, HTML export.
// @author       Rahul Masal
// @match        https://teams.microsoft.com/*
// @match        https://teams.live.com/*
// @grant        GM_download
// @grant        GM_addStyle
// @grant        GM_notification
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    // Define shared utilities directly in this script
    // (Avoiding dynamic script injection to prevent syntax issues)

    // Selector strategies
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
        sanitized = sanitized.replace(/&/g, '&amp;');
        sanitized = sanitized.replace(/</g, '&lt;');
        sanitized = sanitized.replace(/>/g, '&gt;');
        sanitized = sanitized.replace(/"/g, '&quot;');
        sanitized = sanitized.replace(/'/g, '&#x27;');
        return sanitized;
    }

    /**
     * Auto-scroll the chat container to load older messages.
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

    // Note: We've defined the shared utilities directly above
    // so we don't need to import them from window.TeamsChatUtils

    // ========================================================
    // CONFIGURATION
    // ========================================================
    const CONFIG = {
        SCROLL_DELAY: 1200,        // ms between scroll attempts (updated from shared utils default)
        MAX_SCROLL_ATTEMPTS: 60,   // max scrolls for full history
        FLOAT_BTN_POSITION: { bottom: '24px', right: '24px' },
    };

    // ========================================================
    // STYLES
    // ========================================================
    GM_addStyle(
        "/* Floating action button */\n" +
        "#tcd-fab {\n" +
        "    position: fixed;\n" +
        "    bottom: " + CONFIG.FLOAT_BTN_POSITION.bottom + ";\n" +
        "    right: " + CONFIG.FLOAT_BTN_POSITION.right + ";\n" +
        "    z-index: 999999;\n" +
        "    display: flex;\n" +
        "    align-items: center;\n" +
        "    gap: 8px;\n" +
        "    padding: 12px 20px;\n" +
        "    background: linear-gradient(135deg, #7b68ee, #6c5ce7);\n" +
        "    color: #fff;\n" +
        "    border: none;\n" +
        "    border-radius: 50px;\n" +
        "    font-family: 'Segoe UI', system-ui, sans-serif;\n" +
        "    font-size: 14px;\n" +
        "    font-weight: 600;\n" +
        "    cursor: pointer;\n" +
        "    box-shadow: 0 4px 20px rgba(123, 104, 238, 0.4);\n" +
        "    transition: all 0.25s ease;\n" +
        "    user-select: none;\n" +
        "}\n" +
        "#tcd-fab:hover {\n" +
        "    transform: translateY(-2px);\n" +
        "    box-shadow: 0 6px 28px rgba(123, 104, 238, 0.55);\n" +
        "}\n" +
        "#tcd-fab:active { transform: scale(0.96); }\n" +
        "\n" +
        "/* Panel */\n" +
        "#tcd-panel {\n" +
        "    position: fixed;\n" +
        "    bottom: 80px;\n" +
        "    right: 24px;\n" +
        "    z-index: 999999;\n" +
        "    width: 340px;\n" +
        "    background: #0f0f1a;\n" +
        "    border: 1px solid rgba(123, 104, 238, 0.2);\n" +
        "    border-radius: 16px;\n" +
        "    box-shadow: 0 12px 48px rgba(0,0,0,0.5);\n" +
        "    font-family: 'Segoe UI', system-ui, sans-serif;\n" +
        "    color: #e8e8f0;\n" +
        "    overflow: hidden;\n" +
        "    display: none;\n" +
        "    animation: tcdSlideUp 0.25s ease;\n" +
        "}\n" +
        "@keyframes tcdSlideUp {\n" +
        "    from { opacity: 0; transform: translateY(12px); }\n" +
        "    to { opacity: 1; transform: translateY(0); }\n" +
        "}\n" +
        "#tcd-panel.open { display: block; }\n" +
        "\n" +
        "#tcd-panel-header {\n" +
        "    display: flex; align-items: center; justify-content: space-between;\n" +
        "    padding: 16px 18px; border-bottom: 1px solid rgba(255,255,255,0.06);\n" +
        "}\n" +
        "#tcd-panel-header h3 {\n" +
        "    margin: 0; font-size: 15px; font-weight: 700; color: #e8e8f0;\n" +
        "}\n" +
        "#tcd-panel-close {\n" +
        "    background: none; border: none; color: #6c6c80; font-size: 18px;\n" +
        "    cursor: pointer; padding: 2px 6px; border-radius: 6px;\n" +
        "}\n" +
        "#tcd-panel-close:hover { color: #e8e8f0; background: rgba(255,255,255,0.06); }\n" +
        "\n" +
        "#tcd-panel-body { padding: 16px 18px; }\n" +
        "\n" +
        ".tcd-status {\n" +
        "    display: flex; align-items: center; gap: 8px;\n" +
        "    padding: 10px 12px; background: #1a1a2e; border-radius: 8px;\n" +
        "    font-size: 13px; color: #9a9ab0; margin-bottom: 14px;\n" +
        "    border-left: 3px solid #7b68ee;\n" +
        "}\n" +
        ".tcd-status.success { border-left-color: #00c853; color: #00c853; background: rgba(0,200,83,0.08); }\n" +
        ".tcd-status.error { border-left-color: #ef5350; color: #ef5350; background: rgba(239,83,80,0.08); }\n" +
        ".tcd-status.working { border-left-color: #ffa726; color: #ffa726; background: rgba(255,167,38,0.08); }\n" +
        "\n" +
        ".tcd-progress {\n" +
        "    height: 4px; background: #1a1a2e; border-radius: 20px;\n" +
        "    margin-bottom: 14px; overflow: hidden; display: none;\n" +
        "}\n" +
        ".tcd-progress-bar {\n" +
        "    height: 100%; width: 0%; border-radius: 20px;\n" +
        "    background: linear-gradient(90deg, #7b68ee, #a78bfa);\n" +
        "    transition: width 0.3s ease;\n" +
        "}\n" +
        "\n" +
        ".tcd-btn-row { display: flex; gap: 8px; margin-bottom: 10px; }\n" +
        "\n" +
        ".tcd-btn {\n" +
        "    flex: 1; padding: 9px 14px; border: none; border-radius: 8px;\n" +
        "    font-family: inherit; font-size: 13px; font-weight: 600;\n" +
        "    cursor: pointer; transition: all 0.2s ease;\n" +
        "}\n" +
        ".tcd-btn:active { transform: scale(0.97); }\n" +
        ".tcd-btn:disabled { opacity: 0.4; cursor: not-allowed; }\n" +
        "\n" +
        ".tcd-btn-primary { background: #7b68ee; color: #fff; }\n" +
        ".tcd-btn-primary:hover:not(:disabled) { background: #6c5ce7; }\n" +
        "\n" +
        ".tcd-btn-outline {\n" +
        "    background: transparent; color: #7b68ee;\n" +
        "    border: 1px solid rgba(123,104,238,0.25);\n" +
        "}\n" +
        ".tcd-btn-outline:hover:not(:disabled) { background: rgba(123,104,238,0.1); }\n" +
        "\n" +
        ".tcd-btn-success { background: #00c853; color: #fff; }\n" +
        ".tcd-btn-success:hover:not(:disabled) { background: #00b84a; }\n" +
        "\n" +
        ".tcd-select {\n" +
        "    width: 100%; padding: 9px 12px; background: #1a1a2e; color: #e8e8f0;\n" +
        "    border: 1px solid rgba(255,255,255,0.06); border-radius: 8px;\n" +
        "    font-family: inherit; font-size: 13px; margin-bottom: 10px;\n" +
        "    cursor: pointer; appearance: none;\n" +
        "    background-image: url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%239a9ab0' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E\");\n" +
        "    background-repeat: no-repeat; background-position: right 10px center;\n" +
        "}\n" +
        ".tcd-select:focus { outline: none; border-color: #7b68ee; }\n" +
        "\n" +
        ".tcd-count {\n" +
        "    text-align: center; font-size: 12px; font-weight: 600;\n" +
        "    color: #7b68ee; padding: 6px; background: rgba(123,104,238,0.08);\n" +
        "    border-radius: 6px; margin-bottom: 10px; display: none;\n" +
        "}\n" +
        "\n" +
        ".tcd-footer {\n" +
        "    padding: 10px 18px; border-top: 1px solid rgba(255,255,255,0.06);\n" +
        "    text-align: center; font-size: 11px; color: #6c6c80;\n" +
        "}"
    );

    // ========================================================
    // HELPER FUNCTIONS adapted to use shared utilities
    // ========================================================

    function qsf(root, selectors) {
        // Use the shared utility if available, otherwise fall back to the original implementation
        if (typeof querySelectorFallback === 'function') {
            return querySelectorFallback(root, selectors);
        }
        // Original implementation
        for (const sel of selectors.split(',').map(s => s.trim())) {
            const el = root.querySelector(sel);
            if (el) return el;
        }
        return null;
    }

    function qsaf(root, selectors) {
        // Use the shared utility if available, otherwise fall back to the original implementation
        if (typeof querySelectorAllFallback === 'function') {
            return querySelectorAllFallback(root, selectors);
        }
        // Original implementation
        for (const sel of selectors.split(',').map(s => s.trim())) {
            const els = root.querySelectorAll(sel);
            if (els.length > 0) return els;
        }
        return [];
    }

    function detectStrategy() {
        // Use the shared utility if available, otherwise fall back to the original implementation
        if (typeof detectStrategyFunc === 'function') {
            return detectStrategyFunc();
        }

        // Original implementation
        for (const s of SELECTOR_STRATEGIES) {
            const container = qsf(document, s.chatContainer);
            if (container && qsaf(container, s.message).length > 0) {
                return { strategy: s, container };
            }
        }
        return null;
    }

    // Define the function separately so we can reference it
    function detectStrategyFunc() {
        for (const s of SELECTOR_STRATEGIES) {
            const container = querySelectorFallback(document, s.chatContainer);
            if (container && querySelectorAllFallback(container, s.message).length > 0) {
                console.log('[TeamsChatDownloader] Using selector strategy: ' + s.name);
                return { strategy: s, container };
            }
        }
        return null;
    }

    function extractTimestamp(el) {
        // Use the shared utility if available, otherwise fall back to the original implementation
        if (typeof extractTimestampFunc === 'function') {
            return extractTimestampFunc(el, null); // Pass null for strategy as it's not needed in shared version
        }

        // Original implementation
        if (!el) return Date.now();
        const dataTs = el.getAttribute('data-timestamp');
        if (dataTs) { let t = parseInt(dataTs); if (!isNaN(t)) { return t < 1e12 ? t * 1000 : t; } }
        const dt = el.getAttribute('datetime');
        if (dt) { const d = new Date(dt); if (!isNaN(d)) return d.getTime(); }
        const title = el.getAttribute('title');
        if (title) { const d = new Date(title); if (!isNaN(d)) return d.getTime(); }
        const txt = el.textContent?.trim();
        if (txt) { const d = new Date(txt); if (!isNaN(d)) return d.getTime(); }
        return Date.now();
    }

    // Define the function separately so we can reference it
    function extractTimestampFunc(el, strategy) {
        if (!el) return null;

        // Try data-timestamp attribute first
        const dataTs = el.getAttribute('data-timestamp');
        if (dataTs) {
            let ts = parseInt(dataTs, 10);
            if (!isNaN(ts)) {
                // If it looks like seconds instead of milliseconds, convert
                if (ts < 1e12) ts *= 1000;
                return ts;
            }
        }

        // Try datetime attribute (ISO 8601)
        const datetime = el.getAttribute('datetime');
        if (datetime) {
            const d = new Date(datetime);
            if (!isNaN(d.getTime())) return d.getTime();
        }

        // Try title attribute as a displayable date
        const title = el.getAttribute('title');
        if (title) {
            const d = new Date(title);
            if (!isNaN(d.getTime())) return d.getTime();
        }

        // Fallback: use textContent as a date string
        const text = el.textContent?.trim();
        if (text) {
            const d = new Date(text);
            if (!isNaN(d.getTime())) return d.getTime();
        }

        return Date.now(); // Last resort fallback
    }

    function extractMessages(container, strategy) {
        const messages = [];

        const seen = new Set();

        const elements = qsaf(container, strategy.message);

        elements.forEach(el => {

            const sender = sanitizeText(qsf(el, strategy.sender)?.textContent?.trim() || 'Unknown');

            const text = sanitizeText(qsf(el, strategy.content)?.textContent?.trim() || '');

            const timestamp = extractTimestamp(qsf(el, strategy.timestamp));

            if (!text && sender === 'Unknown') return;

            const key = sender + '|' + timestamp + '|' + text.substring(0, 50);

            if (seen.has(key)) return;

            seen.add(key);

            const attachments = [];

            qsaf(el, strategy.attachment).forEach(a => {

                const name = sanitizeText(qsf(a, strategy.attachmentName)?.textContent?.trim());

                if (name) attachments.push(name);

            });

            messages.push({ sender, text, timestamp, attachments });

        });

        return messages.sort((a, b) => a.timestamp - b.timestamp);

    }

    async function autoScroll(container, onProgress) {
        // Use the shared utility if available, otherwise fall back to the original implementation
        if (typeof loadFullChat === 'function') {
            return loadFullChat(container, onProgress, CONFIG.MAX_SCROLL_ATTEMPTS, CONFIG.SCROLL_DELAY);
        }

        // Original implementation
        let prevH = 0, attempts = 0;
        while (attempts < CONFIG.MAX_SCROLL_ATTEMPTS) {
            const curH = container.scrollHeight;
            if (curH === prevH && attempts > 2) break;
            prevH = curH;
            container.scrollTop = 0;
            onProgress(Math.min(90, Math.round((attempts / CONFIG.MAX_SCROLL_ATTEMPTS) * 90)));
            await new Promise(r => setTimeout(r, CONFIG.SCROLL_DELAY));
            attempts++;
        }
        container.scrollTop = container.scrollHeight;
    }

    // ========================================================
    // FORMATTERS - Updated HTML sanitizer
    // ========================================================

    function toText(data) {
        let r = 'Microsoft Teams Chat Export\nExported: ' + new Date().toLocaleString() + '\nMessages: ' + data.length + '\n' + '═'.repeat(50) + '\n\n';
        data.forEach(m => {
            r += '[' + new Date(m.timestamp).toLocaleString() + '] ' + m.sender + ':\n  ' + m.text + '\n';
            if (m.attachments.length) r += '  📎 ' + m.attachments.join(', ') + '\n';
            r += '\n';
        });
        return r;
    }

    function toJSON(data) {
        return JSON.stringify({
            exportDate: new Date().toISOString(),
            messageCount: data.length,
            messages: data.map(m => ({
                sender: m.sender, text: m.text,
                timestamp: new Date(m.timestamp).toISOString(),
                timestampMs: m.timestamp,
                attachments: m.attachments
            }))
        }, null, 2);
    }

    function toCSV(data) {
        const esc = s => '"' + (s || '').replace(/"/g, '""').replace(/\n/g, ' ') + '"';
        let csv = 'Timestamp,Sender,Message,Attachments\n';
        data.forEach(m => {
            csv += esc(new Date(m.timestamp).toISOString()) + ',' + esc(m.sender) + ',' + esc(m.text) + ',' + esc(m.attachments.join('; ')) + '\n';
        });
        return csv;
    }

    function toHTML(data) {
        // Use the improved sanitize function from shared utilities
        const esc = s => sanitizeText(s);
        let h = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Teams Chat Export</title>\n';
        h += '<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:\'Segoe UI\',system-ui,sans-serif;background:#1a1a2e;color:#e0e0e0;padding:24px;max-width:800px;margin:0 auto}\n';
        h += 'h1{color:#7b68ee;margin-bottom:8px}.meta{color:#888;margin-bottom:24px;font-size:.9em}\n';
        h += '.m{background:#16213e;border-radius:12px;padding:14px 18px;margin-bottom:8px;border-left:3px solid #7b68ee}\n';
        h += '.s{color:#7b68ee;font-weight:600}.t{color:#666;font-size:.8em;margin-left:8px}\n';
        h += '.txt{margin-top:6px;line-height:1.5;white-space:pre-wrap}.att{margin-top:8px;padding:6px 10px;background:#0f3460;border-radius:6px;font-size:.85em;color:#a0c4ff}</style></head>\n';
        h += '<body><h1>📋 Teams Chat Export</h1><p class="meta">Exported: ' + esc(new Date().toLocaleString()) + ' · ' + data.length + ' messages</p>\n';
        data.forEach(m => {
            h += '<div class="m"><span class="s">' + esc(m.sender) + '</span><span class="t">' + esc(new Date(m.timestamp).toLocaleString()) + '</span><div class="txt">' + esc(m.text) + '</div>';
            if (m.attachments.length) h += '<div class="att">📎 ' + esc(m.attachments.join(', ')) + '</div>';
            h += '</div>\n';
        });
        return h + '</body></html>';
    }

    // ========================================================
    // DOWNLOAD
    // ========================================================

    function downloadFile(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    // ========================================================
    // UI
    // ========================================================

    let chatData = null;
    let isWorking = false;

    // Create FAB
    const fab = document.createElement('button');
    fab.id = 'tcd-fab';
    fab.innerHTML = '💬 Chat Downloader';
    document.body.appendChild(fab);

    // Create Panel
    const panel = document.createElement('div');
    panel.id = 'tcd-panel';
    panel.innerHTML = '<div id="tcd-panel-header"><h3>💬 Teams Chat Downloader</h3><button id="tcd-panel-close">✕</button></div><div id="tcd-panel-body"><div class="tcd-status" id="tcd-status">Ready — open a chat and scan</div><div class="tcd-progress" id="tcd-progress"><div class="tcd-progress-bar" id="tcd-pbar"></div></div><div class="tcd-btn-row"><button class="tcd-btn tcd-btn-primary" id="tcd-scan">🔍 Quick Scan</button><button class="tcd-btn tcd-btn-outline" id="tcd-full">📜 Full History</button></div><select class="tcd-select" id="tcd-format"><option value="txt">📄 Plain Text (.txt)</option><option value="json">📋 JSON (.json)</option><option value="csv">📊 CSV (.csv)</option><option value="html">🌐 HTML (.html)</option></select><div class="tcd-count" id="tcd-count"></div><button class="tcd-btn tcd-btn-success" id="tcd-download" disabled style="width:100%">⬇️ Download</button></div><div class="tcd-footer">All data stays in your browser · v1.1.1</div>';
    document.body.appendChild(panel);

    // DOM refs
    const $ = id => document.getElementById(id);
    const statusEl = $('#tcd-status');
    const progressEl = $('#tcd-progress');
    const pbar = $('#tcd-pbar');
    const countEl = $('#tcd-count');
    const scanBtn = $('#tcd-scan');
    const fullBtn = $('#tcd-full');
    const downloadBtn = $('#tcd-download');
    const formatSel = $('#tcd-format');

    // Events
    fab.onclick = () => panel.classList.toggle('open');
    $('#tcd-panel-close').onclick = () => panel.classList.remove('open');
    scanBtn.onclick = () => scan(false);
    fullBtn.onclick = () => scan(true);
    downloadBtn.onclick = doDownload;

    function setStatus(text, type = '') {
        statusEl.textContent = text;
        statusEl.className = 'tcd-status ' + type;
    }

    function setProgress(v) {
        progressEl.style.display = 'block';
        pbar.style.width = v + '%';
    }

    async function scan(full) {
        if (isWorking) return;
        isWorking = true;
        chatData = null;
        downloadBtn.disabled = true;
        scanBtn.disabled = fullBtn.disabled = true;
        countEl.style.display = 'none';

        setStatus(full ? '📜 Loading full history...' : '🔍 Scanning chat...', 'working');
        setProgress(5);

        try {
            const detected = detectStrategyFunc();
            if (!detected) {
                setStatus('❌ Chat not found — open a chat first', 'error');
                progressEl.style.display = 'none';
                return;
            }

            const { strategy, container } = detected;

            if (full) {
                await autoScroll(container, setProgress);
            }

            setProgress(95);
            const messages = extractMessages(container, strategy);

            if (messages.length === 0) {
                setStatus('❌ No messages found', 'error');
                return;
            }

            chatData = messages;

            setStatus('✅ Found ' + messages.length + ' messages!', 'success');

            countEl.textContent = messages.length + ' messages · ' + strategy.name + ' strategy';

            countEl.style.display = 'block';
            downloadBtn.disabled = false;
            setProgress(100);
        } catch (err) {
            setStatus(`❌ Error: ${err.message}`, 'error');
        } finally {
            setTimeout(() => { progressEl.style.display = 'none'; }, 600);
            scanBtn.disabled = fullBtn.disabled = false;
            isWorking = false;
        }
    }

    function doDownload() {
        if (!chatData) return;
        const fmt = formatSel.value;
        const date = new Date().toISOString().slice(0, 10);
        const map = {
            txt:  { fn: toText, mime: 'text/plain' },
            json: { fn: toJSON, mime: 'application/json' },
            csv:  { fn: toCSV,  mime: 'text/csv' },
            html: { fn: toHTML, mime: 'text/html' },
        };
        const { fn, mime } = map[fmt];
        downloadFile(fn(chatData), `teams-chat-${date}.${fmt}`, mime);
        setStatus('🎉 Downloaded!', 'success');
    }


})();
