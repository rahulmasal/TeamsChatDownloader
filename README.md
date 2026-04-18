# Teams Chat Downloader

A Chrome extension for downloading Microsoft Teams chat history from the web version.

## ✨ Features

- **Multiple Export Formats** — Download as Plain Text, JSON, CSV, or HTML
- **Full History Loading** — Auto-scrolls to capture the entire chat, not just visible messages
- **Resilient DOM Parsing** — Multiple selector strategies handle Teams UI changes gracefully
- **Message Deduplication** — Automatically removes duplicate messages
- **Smart Timestamp Handling** — Supports Unix timestamps (ms and seconds), ISO 8601, and text dates
- **Session Persistence** — Scan results survive popup close/reopen (up to 10 minutes)
- **Privacy First** — All processing happens locally in your browser; no data is transmitted externally
- **Beautiful Dark UI** — Modern, polished interface with smooth animations

## 📦 Installation

1. Clone or download this repository
2. Open Chrome and navigate to `chrome://extensions/`
3. Enable **Developer mode** in the top right
4. Click **Load unpacked** and select the `TeamsChatDownloader` folder
5. The extension icon will appear in your Chrome toolbar

## 🚀 Usage

1. Navigate to [Microsoft Teams Web](https://teams.microsoft.com/) (or [teams.live.com](https://teams.live.com/) for personal accounts)
2. Open the chat you want to download
3. Click the extension icon in the toolbar
4. Choose a scan mode:
   - **Scan Chat** — Quick scan of currently visible messages
   - **Full History** — Auto-scrolls to load the entire conversation (may take a minute for long chats)
5. Select your preferred export format (TXT, JSON, CSV, or HTML)
6. Click **Download** to save the file

## 📁 Export Formats

### Plain Text (.txt)
```
[4/18/2026, 9:30:00 PM] John Doe:
  Hey, how's the project going?

[4/18/2026, 9:31:00 PM] Jane Smith:
  Great! Just pushed the latest changes.
  📎 Attachments: report.pdf
```

### JSON (.json)
Structured format with ISO timestamps, ideal for programmatic processing.

### CSV (.csv)
Spreadsheet-compatible with columns: Timestamp, Sender, Message, Attachments.

### HTML (.html)
Beautiful dark-themed standalone page that can be opened in any browser.

## 🔒 Permissions

| Permission | Purpose |
|------------|---------|
| `activeTab` | Access the current Teams tab to read chat messages |
| `storage` | Persist scan results across popup sessions |
| `downloads` | Save exported chat files to disk |
| Host access | Only `teams.microsoft.com` and `teams.live.com` |

## 🛠 Architecture

```
TeamsChatDownloader/
├── manifest.json     # Extension config (Manifest V3)
├── content.js        # Chat extraction logic (injected into Teams pages)
├── popup.html        # Extension popup UI
├── popup.js          # Popup logic, formatting, and download handling
├── styles.css        # Modern dark theme styles
├── icon16.png        # Toolbar icon
├── icon48.png        # Extensions page icon
└── icon128.png       # Chrome Web Store icon
```

- **Content Script** (`content.js`) — Injected into Teams pages. Uses multiple selector strategies to find chat elements. Supports auto-scrolling to load full history.
- **Popup** (`popup.html/js/css`) — User interface. Handles scan requests, format selection, and file downloads. All rendering uses safe DOM APIs (no `innerHTML` with user content).

## 🐛 Troubleshooting

| Problem | Solution |
|---------|----------|
| Extension doesn't detect chat | Make sure you're in a Teams chat (not a channel) and the conversation is open |
| "Connection error" message | Refresh the Teams tab — the content script may need to reload |
| Only a few messages captured | Use **Full History** mode to auto-scroll and load all messages |
| Timestamps appear wrong | The extension handles multiple timestamp formats; report an issue if dates seem off |

## 📄 License

This project is open source and available under the MIT License.
