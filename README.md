# 📋 Teams Chat Downloader

**Export Microsoft Teams chat history with media — 3 tools for every use case.**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Chrome Extension](https://img.shields.io/badge/Chrome-Extension-green.svg)](#-chrome-extension)
[![Tampermonkey](https://img.shields.io/badge/Tampermonkey-Script-yellow.svg)](#-tampermonkey-script)
[![Graph API](https://img.shields.io/badge/Graph_API-CLI-purple.svg)](#-graph-api-exporter)

---

## 🎯 Pick Your Tool

| Tool | Best For | Media Export | Setup Time |
|------|----------|-------------|-----------|
| [**Chrome Extension**](#-chrome-extension) | Sharing with non-technical users | Text only | 30 seconds |
| [**Tampermonkey Script**](#-tampermonkey-script) | Personal use, fastest option | Text only | 1 minute |
| [**Graph API Exporter**](#-graph-api-exporter) | Full export with images & files | ✅ All media | 10 minutes |

---

## 🧩 Chrome Extension

A polished Chrome extension with a dark-themed popup UI.

**Features:**
- 🔍 Quick Scan or Full History (auto-scroll)
- 📄 Export as TXT, JSON, CSV, or HTML
- 🛡️ Multiple DOM selector strategies for resilience
- 🔒 All data stays in your browser

**Quick Start:**
1. Go to `chrome://extensions/` → Enable Developer Mode
2. Click **Load unpacked** → Select the project root folder
3. Navigate to Teams → Click the extension icon → Scan & Download

![Extension UI](https://img.shields.io/badge/UI-Dark_Theme-1a1a2e?style=flat-square)

---

## 🐒 Tampermonkey Script

A single-file userscript that injects a floating download button directly into Teams.

**Features:**
- ⚡ Fastest option — zero message-passing overhead
- 💬 Floating "Chat Downloader" button on every Teams page
- 📄 Same 4 export formats (TXT, JSON, CSV, HTML)
- 🔄 Auto-scroll for full history

**Quick Start:**
1. Install [Tampermonkey](https://www.tampermonkey.net/) in your browser
2. **[Click here to auto-install the script](https://raw.githubusercontent.com/rahulmasal/TeamsChatDownloader/main/tampermonkey/teams-chat-downloader.user.js)** (Tampermonkey will prompt you to install it)
3. Open Teams → Click "💬 Chat Downloader" button in bottom-right

---

## 🔗 Graph API Exporter

A Node.js CLI that uses Microsoft Graph API to export **complete chat history with all media and attachments**.

**Features:**
- 📥 Complete message history via official API (no DOM scraping)
- 🖼️ Downloads inline images, file attachments, GIFs
- 👍 Preserves reactions, edits, deletions, mentions
- 📊 Interactive mode — browse & select chats
- 🔄 Batch export all chats at once
- 🔐 Secure device-code authentication with token caching
- ⏱️ Rate limit handling with automatic retry

**Quick Start:**
```bash
cd graph-exporter
npm install
# Edit config.js → Add your Azure AD Client ID
node index.js
```

**Output structure:**
```
exported-chats/
├── John Doe, Jane Smith/
│   ├── chat.html          ← Beautiful dark-themed viewer
│   ├── chat.json          ← Structured data
│   ├── chat.txt           ← Plain text
│   ├── metadata.json      ← Chat info & stats
│   └── media/
│       ├── inline_abc.png ← Pasted images
│       └── report.pdf     ← File attachments
```

> 📖 See [`graph-exporter/README.md`](graph-exporter/README.md) for Azure AD setup instructions.

---

## 📂 Project Structure

```
TeamsChatDownloader/
│
├── manifest.json           ┐
├── content.js              │ Chrome Extension
├── popup.html / .js        │ (browser-based)
├── styles.css              │
├── icon16/48/128.png       ┘
│
├── tampermonkey/
│   └── teams-chat-downloader.user.js    ← Tampermonkey script
│
└── graph-exporter/
    ├── index.js            ← CLI entry point
    ├── auth.js             ← MSAL authentication
    ├── graph.js            ← Graph API client
    ├── exporter.js         ← Export engine
    ├── config.js           ← Azure AD configuration
    └── package.json        ← Dependencies
```

---

## 🔒 Privacy & Security

- **No data is transmitted externally** — all processing happens locally
- Chrome Extension & Tampermonkey run entirely in your browser
- Graph API Exporter authenticates via Microsoft's official OAuth2 flow
- Token cache is stored locally and git-ignored
- No telemetry, no analytics, no tracking

---

## 🛡️ Permissions

### Chrome Extension
| Permission | Purpose |
|-----------|---------|
| `activeTab` | Read chat messages from the current Teams tab |
| `storage` | Cache scan results between popup sessions |
| `downloads` | Save exported files to disk |

### Graph API Exporter
| Permission | Purpose |
|-----------|---------|
| `Chat.Read` | Read your chat messages |
| `Files.Read.All` | Download shared files and attachments |
| `User.Read` | Display your profile name |

---

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

---

**Made with ❤️ by [Rahul Masal](https://github.com/rahulmasal)**
