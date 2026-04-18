// ==UserScript==
// @name         Teams Chat Downloader
// @namespace    https://github.com/rahulmasal/TeamsChatDownloader
// @version      1.1.0
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

    // ========================================================
    // CONFIGURATION
    // ========================================================
    const CONFIG = {
        SCROLL_DELAY: 1000,        // ms between scroll attempts
        MAX_SCROLL_ATTEMPTS: 60,   // max scrolls for full history
        FLOAT_BTN_POSITION: { bottom: '24px', right: '24px' },
    };

    // ========================================================
    // SELECTOR STRATEGIES
    // ========================================================
    const STRATEGIES = [
        {
            name: 'data-tid',
            chat: '[data-tid="chat-pane-list"], [data-tid="chat-list"], [data-tid="message-pane-list"]',
            msg: '[data-tid="chat-pane-message"], [data-tid="message"], [data-tid="message-body"]',
            sender: '[data-tid="message-author-name"], [data-tid="message-sender"]',
            content: '[data-tid="message-text"], [data-tid="message-content"]',
            time: '[data-tid="message-timestamp"], [data-tid="ts-message-timestamp"]',
            attach: '[data-tid="file-card"], [data-tid="message-attachment"]',
            attachName: '[data-tid="file-card-name"], [data-tid="attachment-name"]',
        },
        {
            name: 'aria-role',
            chat: '[role="main"] [role="list"], [role="log"]',
            msg: '[role="listitem"], [data-is-focusable="true"]',
            sender: '[data-testid="message-author"], .ui-chat__message__author',
            content: '[data-testid="message-body"], .ui-chat__message__content',
            time: 'time[datetime], [data-testid="message-timestamp"]',
            attach: '.ui-attachment, [data-testid="file-attachment"]',
            attachName: '.ui-attachment__header, [data-testid="file-name"]',
        },
        {
            name: 'class-based',
            chat: '.ts-message-list-container, .message-list',
            msg: '.ts-message, .message-body-content',
            sender: '.ts-message-sender, .message-sender-name',
            content: '.ts-message-text, .message-body-text',
            time: '.ts-message-timestamp, time',
            attach: '.ts-attachment, .file-attachment',
            attachName: '.ts-attachment-name, .file-name',
        },
    ];

    // ========================================================
    // STYLES
    // ========================================================
    GM_addStyle(`
        /* Floating action button */
        #tcd-fab {
            position: fixed;
            bottom: ${CONFIG.FLOAT_BTN_POSITION.bottom};
            right: ${CONFIG.FLOAT_BTN_POSITION.right};
            z-index: 999999;
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 12px 20px;
            background: linear-gradient(135deg, #7b68ee, #6c5ce7);
            color: #fff;
            border: none;
            border-radius: 50px;
            font-family: 'Segoe UI', system-ui, sans-serif;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            box-shadow: 0 4px 20px rgba(123, 104, 238, 0.4);
            transition: all 0.25s ease;
            user-select: none;
        }
        #tcd-fab:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 28px rgba(123, 104, 238, 0.55);
        }
        #tcd-fab:active { transform: scale(0.96); }

        /* Panel */
        #tcd-panel {
            position: fixed;
            bottom: 80px;
            right: 24px;
            z-index: 999999;
            width: 340px;
            background: #0f0f1a;
            border: 1px solid rgba(123, 104, 238, 0.2);
            border-radius: 16px;
            box-shadow: 0 12px 48px rgba(0,0,0,0.5);
            font-family: 'Segoe UI', system-ui, sans-serif;
            color: #e8e8f0;
            overflow: hidden;
            display: none;
            animation: tcdSlideUp 0.25s ease;
        }
        @keyframes tcdSlideUp {
            from { opacity: 0; transform: translateY(12px); }
            to { opacity: 1; transform: translateY(0); }
        }
        #tcd-panel.open { display: block; }

        #tcd-panel-header {
            display: flex; align-items: center; justify-content: space-between;
            padding: 16px 18px; border-bottom: 1px solid rgba(255,255,255,0.06);
        }
        #tcd-panel-header h3 {
            margin: 0; font-size: 15px; font-weight: 700; color: #e8e8f0;
        }
        #tcd-panel-close {
            background: none; border: none; color: #6c6c80; font-size: 18px;
            cursor: pointer; padding: 2px 6px; border-radius: 6px;
        }
        #tcd-panel-close:hover { color: #e8e8f0; background: rgba(255,255,255,0.06); }

        #tcd-panel-body { padding: 16px 18px; }

        .tcd-status {
            display: flex; align-items: center; gap: 8px;
            padding: 10px 12px; background: #1a1a2e; border-radius: 8px;
            font-size: 13px; color: #9a9ab0; margin-bottom: 14px;
            border-left: 3px solid #7b68ee;
        }
        .tcd-status.success { border-left-color: #00c853; color: #00c853; background: rgba(0,200,83,0.08); }
        .tcd-status.error { border-left-color: #ef5350; color: #ef5350; background: rgba(239,83,80,0.08); }
        .tcd-status.working { border-left-color: #ffa726; color: #ffa726; background: rgba(255,167,38,0.08); }

        .tcd-progress {
            height: 4px; background: #1a1a2e; border-radius: 20px;
            margin-bottom: 14px; overflow: hidden; display: none;
        }
        .tcd-progress-bar {
            height: 100%; width: 0%; border-radius: 20px;
            background: linear-gradient(90deg, #7b68ee, #a78bfa);
            transition: width 0.3s ease;
        }

        .tcd-btn-row { display: flex; gap: 8px; margin-bottom: 10px; }

        .tcd-btn {
            flex: 1; padding: 9px 14px; border: none; border-radius: 8px;
            font-family: inherit; font-size: 13px; font-weight: 600;
            cursor: pointer; transition: all 0.2s ease;
        }
        .tcd-btn:active { transform: scale(0.97); }
        .tcd-btn:disabled { opacity: 0.4; cursor: not-allowed; }

        .tcd-btn-primary { background: #7b68ee; color: #fff; }
        .tcd-btn-primary:hover:not(:disabled) { background: #6c5ce7; }

        .tcd-btn-outline {
            background: transparent; color: #7b68ee;
            border: 1px solid rgba(123,104,238,0.25);
        }
        .tcd-btn-outline:hover:not(:disabled) { background: rgba(123,104,238,0.1); }

        .tcd-btn-success { background: #00c853; color: #fff; }
        .tcd-btn-success:hover:not(:disabled) { background: #00b84a; }

        .tcd-select {
            width: 100%; padding: 9px 12px; background: #1a1a2e; color: #e8e8f0;
            border: 1px solid rgba(255,255,255,0.06); border-radius: 8px;
            font-family: inherit; font-size: 13px; margin-bottom: 10px;
            cursor: pointer; appearance: none;
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%239a9ab0' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E");
            background-repeat: no-repeat; background-position: right 10px center;
        }
        .tcd-select:focus { outline: none; border-color: #7b68ee; }

        .tcd-count {
            text-align: center; font-size: 12px; font-weight: 600;
            color: #7b68ee; padding: 6px; background: rgba(123,104,238,0.08);
            border-radius: 6px; margin-bottom: 10px; display: none;
        }

        .tcd-footer {
            padding: 10px 18px; border-top: 1px solid rgba(255,255,255,0.06);
            text-align: center; font-size: 11px; color: #6c6c80;
        }
    `);

    // ========================================================
    // HELPER FUNCTIONS
    // ========================================================

    function qsf(root, selectors) {
        for (const sel of selectors.split(',').map(s => s.trim())) {
            const el = root.querySelector(sel);
            if (el) return el;
        }
        return null;
    }

    function qsaf(root, selectors) {
        for (const sel of selectors.split(',').map(s => s.trim())) {
            const els = root.querySelectorAll(sel);
            if (els.length > 0) return els;
        }
        return [];
    }

    function detectStrategy() {
        for (const s of STRATEGIES) {
            const container = qsf(document, s.chat);
            if (container && qsaf(container, s.msg).length > 0) {
                return { strategy: s, container };
            }
        }
        return null;
    }

    function extractTimestamp(el) {
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

    function extractMessages(container, strategy) {
        const messages = [];
        const seen = new Set();
        const elements = qsaf(container, strategy.msg);

        elements.forEach(el => {
            const sender = qsf(el, strategy.sender)?.textContent?.trim() || 'Unknown';
            const text = qsf(el, strategy.content)?.textContent?.trim() || '';
            const timestamp = extractTimestamp(qsf(el, strategy.time));

            if (!text && sender === 'Unknown') return;

            const key = `${sender}|${timestamp}|${text.substring(0, 50)}`;
            if (seen.has(key)) return;
            seen.add(key);

            const attachments = [];
            qsaf(el, strategy.attach).forEach(a => {
                const name = qsf(a, strategy.attachName)?.textContent?.trim();
                if (name) attachments.push(name);
            });

            messages.push({ sender, text, timestamp, attachments });
        });

        return messages.sort((a, b) => a.timestamp - b.timestamp);
    }

    async function autoScroll(container, onProgress) {
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
    // FORMATTERS
    // ========================================================

    function toText(data) {
        let r = `Microsoft Teams Chat Export\nExported: ${new Date().toLocaleString()}\nMessages: ${data.length}\n${'═'.repeat(50)}\n\n`;
        data.forEach(m => {
            r += `[${new Date(m.timestamp).toLocaleString()}] ${m.sender}:\n  ${m.text}\n`;
            if (m.attachments.length) r += `  📎 ${m.attachments.join(', ')}\n`;
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
            csv += `${esc(new Date(m.timestamp).toISOString())},${esc(m.sender)},${esc(m.text)},${esc(m.attachments.join('; '))}\n`;
        });
        return csv;
    }

    function toHTML(data) {
        const esc = s => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
        let h = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Teams Chat Export</title>
<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',system-ui,sans-serif;background:#1a1a2e;color:#e0e0e0;padding:24px;max-width:800px;margin:0 auto}
h1{color:#7b68ee;margin-bottom:8px}.meta{color:#888;margin-bottom:24px;font-size:.9em}
.m{background:#16213e;border-radius:12px;padding:14px 18px;margin-bottom:8px;border-left:3px solid #7b68ee}
.s{color:#7b68ee;font-weight:600}.t{color:#666;font-size:.8em;margin-left:8px}
.txt{margin-top:6px;line-height:1.5;white-space:pre-wrap}.att{margin-top:8px;padding:6px 10px;background:#0f3460;border-radius:6px;font-size:.85em;color:#a0c4ff}</style></head>
<body><h1>📋 Teams Chat Export</h1><p class="meta">Exported: ${esc(new Date().toLocaleString())} · ${data.length} messages</p>\n`;
        data.forEach(m => {
            h += `<div class="m"><span class="s">${esc(m.sender)}</span><span class="t">${esc(new Date(m.timestamp).toLocaleString())}</span><div class="txt">${esc(m.text)}</div>`;
            if (m.attachments.length) h += `<div class="att">📎 ${esc(m.attachments.join(', '))}</div>`;
            h += `</div>\n`;
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
    panel.innerHTML = `
        <div id="tcd-panel-header">
            <h3>💬 Teams Chat Downloader</h3>
            <button id="tcd-panel-close">✕</button>
        </div>
        <div id="tcd-panel-body">
            <div class="tcd-status" id="tcd-status">Ready — open a chat and scan</div>
            <div class="tcd-progress" id="tcd-progress"><div class="tcd-progress-bar" id="tcd-pbar"></div></div>
            <div class="tcd-btn-row">
                <button class="tcd-btn tcd-btn-primary" id="tcd-scan">🔍 Quick Scan</button>
                <button class="tcd-btn tcd-btn-outline" id="tcd-full">📜 Full History</button>
            </div>
            <select class="tcd-select" id="tcd-format">
                <option value="txt">📄 Plain Text (.txt)</option>
                <option value="json">📋 JSON (.json)</option>
                <option value="csv">📊 CSV (.csv)</option>
                <option value="html">🌐 HTML (.html)</option>
            </select>
            <div class="tcd-count" id="tcd-count"></div>
            <button class="tcd-btn tcd-btn-success" id="tcd-download" disabled style="width:100%">⬇️ Download</button>
        </div>
        <div class="tcd-footer">All data stays in your browser · v1.1.0</div>
    `;
    document.body.appendChild(panel);

    // DOM refs
    const $ = id => document.getElementById(id);
    const statusEl = $('tcd-status');
    const progressEl = $('tcd-progress');
    const pbar = $('tcd-pbar');
    const countEl = $('tcd-count');
    const scanBtn = $('tcd-scan');
    const fullBtn = $('tcd-full');
    const downloadBtn = $('tcd-download');
    const formatSel = $('tcd-format');

    // Events
    fab.onclick = () => panel.classList.toggle('open');
    $('tcd-panel-close').onclick = () => panel.classList.remove('open');
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
            const detected = detectStrategy();
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
            setStatus(`✅ Found ${messages.length} messages!`, 'success');
            countEl.textContent = `${messages.length} messages · ${strategy.name} strategy`;
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
