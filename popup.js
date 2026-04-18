/**
 * Teams Chat Downloader - Popup Script
 * Handles UI interactions, scan requests, formatting, and downloads.
 */

// ---- DOM References ----
const statusEl = document.getElementById('status');
const statusText = statusEl.querySelector('.status-text');
const statusIcon = statusEl.querySelector('.status-icon');
const scanBtn = document.getElementById('scanBtn');
const scanFullBtn = document.getElementById('scanFullBtn');
const downloadBtn = document.getElementById('downloadBtn');
const formatSelect = document.getElementById('formatSelect');
const progressContainer = document.getElementById('progress');
const progressBar = document.getElementById('progressBar');
const progressLabel = document.getElementById('progressLabel');
const messagesEl = document.getElementById('messages');
const messageCountEl = document.getElementById('messageCount');
const countText = document.getElementById('countText');
const strategyInfo = document.getElementById('strategyInfo');

let chatData = null;
let isProcessing = false;

// ---- Event Listeners ----
scanBtn.addEventListener('click', () => startScan(false));
scanFullBtn.addEventListener('click', () => startScan(true));
downloadBtn.addEventListener('click', downloadChat);

// Restore state from session storage if available
restoreState();

// ---- Core Functions ----

function startScan(fullHistory) {
    if (isProcessing) return;

    isProcessing = true;
    chatData = null;
    downloadBtn.disabled = true;
    scanBtn.disabled = true;
    scanFullBtn.disabled = true;
    messagesEl.innerHTML = '';
    messageCountEl.style.display = 'none';

    setStatus('🔄', fullHistory ? 'Loading full chat history...' : 'Scanning for chat messages...', 'info');
    showProgress(0);

    chrome.tabs.query({ active: true, currentWindow: true }, function (tabs) {
        const tab = tabs[0];

        // Verify we're on a Teams page
        if (!tab?.url?.match(/teams\.(microsoft|live)\.com/)) {
            setStatus('⚠️', 'Please navigate to Microsoft Teams first.', 'warning');
            resetButtons();
            return;
        }

        chrome.tabs.sendMessage(tab.id, { action: 'scanChat', fullHistory }, function (response) {
            // Check for communication errors
            if (chrome.runtime.lastError) {
                setStatus('❌', `Connection error: ${chrome.runtime.lastError.message}`, 'error');
                showErrorHelp('The content script may not be loaded. Try refreshing the Teams page.');
                resetButtons();
                return;
            }

            if (response && response.success) {
                chatData = response.chatData;

                const count = response.messageCount || chatData.length;
                setStatus('✅', `Found ${count} message${count !== 1 ? 's' : ''}! Choose a format and download.`, 'success');
                downloadBtn.disabled = false;

                // Show message count badge
                countText.textContent = `${count} message${count !== 1 ? 's' : ''} found`;
                messageCountEl.style.display = 'flex';

                // Show strategy used
                if (response.strategy) {
                    strategyInfo.textContent = `Strategy: ${response.strategy}`;
                }

                // Show preview
                showMessages(response.preview || []);

                // Persist state
                saveState();
            } else {
                const errMsg = response?.error || 'No chat found. Make sure you have a chat open.';
                setStatus('❌', errMsg, 'error');

                if (response?.diagnostics) {
                    console.log('[TeamsChatDownloader] Diagnostics:', response.diagnostics);
                }
            }

            hideProgress();
            resetButtons();
        });
    });
}

function downloadChat() {
    if (!chatData || chatData.length === 0) return;

    const format = formatSelect.value;
    downloadBtn.disabled = true;
    setStatus('📦', 'Preparing download...', 'info');

    let content, mimeType, extension;

    switch (format) {
        case 'json':
            content = formatAsJSON(chatData);
            mimeType = 'application/json';
            extension = 'json';
            break;
        case 'csv':
            content = formatAsCSV(chatData);
            mimeType = 'text/csv';
            extension = 'csv';
            break;
        case 'html':
            content = formatAsHTML(chatData);
            mimeType = 'text/html';
            extension = 'html';
            break;
        case 'txt':
        default:
            content = formatAsText(chatData);
            mimeType = 'text/plain';
            extension = 'txt';
            break;
    }

    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const timestamp = new Date().toISOString().slice(0, 10);

    chrome.downloads.download(
        {
            url: url,
            filename: `teams-chat-${timestamp}.${extension}`,
            saveAs: true
        },
        function (downloadId) {
            // Always revoke the blob URL to prevent memory leaks
            URL.revokeObjectURL(url);

            if (chrome.runtime.lastError) {
                setStatus('❌', `Download failed: ${chrome.runtime.lastError.message}`, 'error');
            } else {
                setStatus('🎉', 'Download started!', 'success');
                setTimeout(() => {
                    setStatus('💬', 'Ready. Scan again or download in another format.', 'info');
                }, 2500);
            }
            downloadBtn.disabled = false;
        }
    );
}

// ---- Formatters ----

function formatAsText(data) {
    let result = `Microsoft Teams Chat Export\n`;
    result += `Exported: ${new Date().toLocaleString()}\n`;
    result += `Messages: ${data.length}\n`;
    result += '═'.repeat(50) + '\n\n';

    data.forEach(msg => {
        const ts = new Date(msg.timestamp).toLocaleString();
        result += `[${ts}] ${msg.sender}:\n`;
        result += `  ${msg.text}\n`;
        if (msg.attachments && msg.attachments.length > 0) {
            result += `  📎 Attachments: ${msg.attachments.join(', ')}\n`;
        }
        result += '\n';
    });

    result += '═'.repeat(50) + '\n';
    result += `End of export — ${data.length} messages\n`;
    return result;
}

function formatAsJSON(data) {
    const exportData = {
        exportDate: new Date().toISOString(),
        messageCount: data.length,
        messages: data.map(msg => ({
            sender: msg.sender,
            text: msg.text,
            timestamp: new Date(msg.timestamp).toISOString(),
            timestampMs: msg.timestamp,
            attachments: msg.attachments || []
        }))
    };
    return JSON.stringify(exportData, null, 2);
}

function formatAsCSV(data) {
    const escapeCSV = (str) => {
        if (!str) return '""';
        // Escape double quotes and wrap in quotes
        return '"' + str.replace(/"/g, '""').replace(/\n/g, ' ') + '"';
    };

    let csv = 'Timestamp,Sender,Message,Attachments\n';
    data.forEach(msg => {
        const ts = new Date(msg.timestamp).toISOString();
        const attachments = (msg.attachments || []).join('; ');
        csv += `${escapeCSV(ts)},${escapeCSV(msg.sender)},${escapeCSV(msg.text)},${escapeCSV(attachments)}\n`;
    });
    return csv;
}

function formatAsHTML(data) {
    const escapeHTML = (str) => {
        if (!str) return '';
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/\n/g, '<br>');
    };

    let html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Teams Chat Export - ${new Date().toLocaleDateString()}</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            background: #1a1a2e; color: #e0e0e0; padding: 24px;
            max-width: 800px; margin: 0 auto;
        }
        h1 { color: #7b68ee; margin-bottom: 8px; font-size: 1.5em; }
        .meta { color: #888; margin-bottom: 24px; font-size: 0.9em; }
        .message {
            background: #16213e; border-radius: 12px; padding: 14px 18px;
            margin-bottom: 8px; border-left: 3px solid #7b68ee;
        }
        .message:hover { background: #1a2744; }
        .sender { color: #7b68ee; font-weight: 600; }
        .time { color: #666; font-size: 0.8em; margin-left: 8px; }
        .text { margin-top: 6px; line-height: 1.5; white-space: pre-wrap; }
        .attachments {
            margin-top: 8px; padding: 6px 10px; background: #0f3460;
            border-radius: 6px; font-size: 0.85em; color: #a0c4ff;
        }
    </style>
</head>
<body>
    <h1>📋 Teams Chat Export</h1>
    <p class="meta">Exported: ${escapeHTML(new Date().toLocaleString())} · ${data.length} messages</p>
`;

    data.forEach(msg => {
        const ts = new Date(msg.timestamp).toLocaleString();
        html += `    <div class="message">
        <span class="sender">${escapeHTML(msg.sender)}</span>
        <span class="time">${escapeHTML(ts)}</span>
        <div class="text">${escapeHTML(msg.text)}</div>`;

        if (msg.attachments && msg.attachments.length > 0) {
            html += `\n        <div class="attachments">📎 ${escapeHTML(msg.attachments.join(', '))}</div>`;
        }
        html += `\n    </div>\n`;
    });

    html += `</body>\n</html>`;
    return html;
}

// ---- UI Helpers ----

function setStatus(icon, text, type) {
    statusIcon.textContent = icon;
    statusText.textContent = text;
    statusEl.className = `status status-${type}`;
}

function showProgress(value) {
    progressContainer.style.display = 'flex';
    progressBar.style.width = `${value}%`;
    progressLabel.textContent = `${value}%`;
}

function hideProgress() {
    setTimeout(() => {
        progressContainer.style.display = 'none';
        progressBar.style.width = '0%';
        progressLabel.textContent = '0%';
    }, 500);
}

function resetButtons() {
    scanBtn.disabled = false;
    scanFullBtn.disabled = false;
    isProcessing = false;
}

function showMessages(messages) {
    messagesEl.innerHTML = '';
    if (!messages || messages.length === 0) return;

    messages.forEach(msg => {
        const messageEl = document.createElement('div');
        messageEl.className = 'message-preview';

        const header = document.createElement('div');
        header.className = 'message-preview-header';

        const senderSpan = document.createElement('strong');
        senderSpan.textContent = msg.sender;

        const timeSpan = document.createElement('span');
        timeSpan.className = 'message-preview-time';
        timeSpan.textContent = new Date(msg.timestamp).toLocaleTimeString();

        header.appendChild(senderSpan);
        header.appendChild(timeSpan);

        const textP = document.createElement('p');
        textP.textContent = msg.text;

        messageEl.appendChild(header);
        messageEl.appendChild(textP);
        messagesEl.appendChild(messageEl);
    });
}

function showErrorHelp(helpText) {
    const helpEl = document.createElement('div');
    helpEl.className = 'error-help';
    helpEl.textContent = helpText;
    messagesEl.innerHTML = '';
    messagesEl.appendChild(helpEl);
}

// ---- State Persistence ----

function saveState() {
    try {
        chrome.storage.session.set({
            chatData: chatData,
            lastScanTime: Date.now()
        });
    } catch (e) {
        // session storage not available in older Chrome — ignore
    }
}

function restoreState() {
    try {
        chrome.storage.session.get(['chatData', 'lastScanTime'], (result) => {
            if (chrome.runtime.lastError) return;
            if (result.chatData && result.lastScanTime) {
                const ageMinutes = (Date.now() - result.lastScanTime) / 60000;
                if (ageMinutes < 10) { // Only restore if less than 10 minutes old
                    chatData = result.chatData;
                    downloadBtn.disabled = false;
                    const count = chatData.length;
                    setStatus('📋', `Previous scan: ${count} message${count !== 1 ? 's' : ''}. Download or scan again.`, 'info');
                    countText.textContent = `${count} message${count !== 1 ? 's' : ''} (cached)`;
                    messageCountEl.style.display = 'flex';
                }
            }
        });
    } catch (e) {
        // Ignore
    }
}

// ---- Listen for progress updates from content script ----
chrome.runtime.onMessage.addListener(function (request, sender, sendResponse) {
    if (request.action === 'scanProgress') {
        showProgress(request.progress);
    }
});