# Teams Chat Graph Exporter

A Node.js CLI tool that exports Microsoft Teams chat history **with all media and attachments** using the Microsoft Graph API.

Unlike the browser extension/Tampermonkey script (which scrape the web UI), this tool uses the **official API** — meaning it works regardless of which Teams client you use (desktop, web, or mobile), and it can export complete chat history without any scrolling.

## ✨ Features

- **Complete chat export** — All messages, not just visible ones
- **Media download** — Inline images, file attachments, GIFs, stickers
- **Multiple formats** — HTML (beautiful dark theme), JSON, Plain Text
- **Reactions & metadata** — Preserves reactions, edits, deletions, mentions
- **Interactive mode** — Browse and select chats to export
- **Batch export** — Export all chats at once
- **Token caching** — Sign in once, stays authenticated
- **Rate limit handling** — Automatic retry on throttling
- **Progress tracking** — Real-time progress bars

## 📋 Prerequisites

- **Node.js 18+**
- **Microsoft 365 account** (work/school or personal)
- **Azure AD App Registration** (see setup below)

## 🔧 Azure AD Setup (One-Time)

1. Go to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations**
2. Click **New registration**:
   - **Name:** `Teams Chat Exporter`
   - **Supported account types:** `Accounts in any organizational directory and personal Microsoft accounts`
   - **Redirect URI:** Select `Mobile and desktop applications` → `https://login.microsoftonline.com/common/oauth2/nativeclient`
3. Copy the **Application (client) ID**
4. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**:
   - `User.Read`
   - `Chat.Read`
   - `Chat.ReadBasic`
   - `Files.Read`
   - `Files.Read.All`
5. *(If admin)* Click **Grant admin consent**

## 🚀 Installation

```bash
cd graph-exporter
npm install
```

Edit `config.js` and paste your **Application (client) ID**:
```javascript
clientId: 'paste-your-client-id-here',
```

## 📖 Usage

### Interactive Mode (Recommended)
```bash
node index.js
```
Walks you through: sign in → browse chats → select chats → choose format → export.

### List All Chats
```bash
node index.js --list
```

### Export All Chats
```bash
node index.js --export
```

### Export Specific Chat
```bash
node index.js --export --chat CHAT_ID_HERE
```

### Options
| Flag | Description |
|---|---|
| `-l, --list` | List available chats |
| `-e, --export` | Export chats |
| `-c, --chat <id>` | Export a specific chat |
| `-o, --output <dir>` | Output directory (default: `./exported-chats`) |
| `-f, --format <fmt>` | Format: `html`, `json`, `txt`, `all` (default: `all`) |
| `--no-media` | Skip downloading media/attachments |
| `--max <n>` | Max messages per chat |
| `--logout` | Clear saved token |

## 📂 Output Structure

```
exported-chats/
├── John Doe, Jane Smith/
│   ├── chat.html          ← Beautiful dark-themed HTML export
│   ├── chat.json          ← Structured JSON with all metadata
│   ├── chat.txt           ← Plain text export
│   ├── metadata.json      ← Chat info, participants, stats
│   └── media/
│       ├── inline_abc123.png    ← Images pasted in chat
│       ├── report.pdf           ← File attachments
│       └── screenshot.jpg       ← Shared images
├── Project Alpha Team/
│   ├── chat.html
│   ├── chat.json
│   ├── ...
```

## 🔒 Security

- Token is cached locally in `.token-cache.json` (git-ignored)
- Uses **delegated permissions** — only accesses chats you can already see
- No data is sent to any third party
- Run `node index.js --logout` to clear saved credentials

## 📄 License

MIT License
