/**
 * Microsoft Graph API Configuration
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to https://portal.azure.com → Azure Active Directory → App registrations
 * 2. Click "New registration"
 *    - Name: "Teams Chat Exporter"
 *    - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
 *    - Redirect URI: Select "Mobile and desktop applications" → https://login.microsoftonline.com/common/oauth2/nativeclient
 * 3. Copy the "Application (client) ID" and paste below
 * 4. Go to "API permissions" → Add permission → Microsoft Graph → Delegated permissions
 *    - Add: Chat.Read, Chat.ReadBasic, Files.Read, Files.Read.All, User.Read
 * 5. Click "Grant admin consent" (if you are an admin) or ask your admin
 */

const config = {
    auth: {
        // Replace with YOUR Application (client) ID from Azure Portal
        clientId: 'YOUR_CLIENT_ID_HERE',

        // Use 'common' for multi-tenant + personal accounts
        // Use your tenant ID for single-tenant
        authority: 'https://login.microsoftonline.com/common',
    },

    // Graph API scopes (permissions)
    scopes: [
        'User.Read',
        'Chat.Read',
        'Chat.ReadBasic',
        'Files.Read',
        'Files.Read.All',
    ],

    // Export settings
    export: {
        // Output directory for exported chats
        outputDir: './exported-chats',

        // Maximum messages to fetch per chat (null = all)
        maxMessages: null,

        // Download media/attachments
        downloadMedia: true,

        // Export format: 'html', 'json', 'txt', or 'all'
        format: 'all',

        // Maximum concurrent downloads
        concurrency: 3,

        // Include chat metadata (participants, topic, etc.)
        includeMetadata: true,
    },
};

export default config;
