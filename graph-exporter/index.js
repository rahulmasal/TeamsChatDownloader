#!/usr/bin/env node

/**
 * Teams Chat Graph Exporter — CLI Entry Point
 *
 * Usage:
 *   node index.js                    → Interactive mode (list chats, select, export)
 *   node index.js --list             → List all chats
 *   node index.js --export           → Export all chats
 *   node index.js --export --chat ID → Export specific chat
 *   node index.js --logout           → Clear saved token
 */

import { program } from 'commander';
import inquirer from 'inquirer';
import chalk from 'chalk';
import cliProgress from 'cli-progress';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

import config from './config.js';
import { initAuth, clearAuth, getCurrentAccount } from './auth.js';
import { initGraphClient, getMe, listChats } from './graph.js';
import { exportChat } from './exporter.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ================================================================
// CLI SETUP
// ================================================================

program
    .name('teams-chat-exporter')
    .description('Export Microsoft Teams chats with media using Graph API')
    .version('1.0.0')
    .option('-l, --list', 'List all chats')
    .option('-e, --export', 'Export chats')
    .option('-c, --chat <id>', 'Export a specific chat by ID')
    .option('-o, --output <dir>', 'Output directory', config.export.outputDir)
    .option('-f, --format <format>', 'Export format: html, json, txt, all', config.export.format)
    .option('--no-media', 'Skip media/attachment downloads')
    .option('--max <n>', 'Max messages per chat', parseInt)
    .option('--logout', 'Clear saved authentication')
    .parse(process.argv);

const opts = program.opts();

// ================================================================
// MAIN
// ================================================================

async function main() {
    printBanner();

    // Handle logout
    if (opts.logout) {
        clearAuth();
        console.log(chalk.green('✅ Token cache cleared. You will need to sign in again.\n'));
        return;
    }

    // Validate config
    if (config.auth.clientId === 'YOUR_CLIENT_ID_HERE') {
        console.log(chalk.red.bold('\n⚠️  Azure AD App not configured!\n'));
        console.log(chalk.yellow('You need to register an Azure AD application first.\n'));
        printSetupInstructions();
        return;
    }

    // Initialize auth and Graph client
    initAuth(config);
    initGraphClient();

    // Get current user
    try {
        const me = await getMe();
        console.log(chalk.blue(`👤 Signed in as: ${chalk.bold(me.displayName)} (${me.mail || me.userPrincipalName})\n`));
    } catch (err) {
        console.log(chalk.red(`❌ Authentication failed: ${err.message}\n`));
        return;
    }

    // List mode
    if (opts.list) {
        await listAllChats();
        return;
    }

    // Export mode (specific chat)
    if (opts.export && opts.chat) {
        await exportSingleChat(opts.chat);
        return;
    }

    // Export all mode
    if (opts.export) {
        await exportAllChats();
        return;
    }

    // Interactive mode
    await interactiveMode();
}

// ================================================================
// MODES
// ================================================================

async function listAllChats() {
    console.log(chalk.blue('📋 Fetching chat list...\n'));
    const chats = await listChats();

    if (chats.length === 0) {
        console.log(chalk.yellow('No chats found.'));
        return;
    }

    console.log(chalk.green(`Found ${chats.length} chats:\n`));

    const table = chats.map((chat, i) => {
        const name = getChatDisplayName(chat);
        const type = chat.chatType === 'oneOnOne' ? '1:1' : chat.chatType === 'group' ? 'Group' : chat.chatType;
        const updated = new Date(chat.lastUpdatedDateTime).toLocaleDateString();
        const members = (chat.members || []).length;
        return {
            '#': i + 1,
            Type: type,
            Name: name.substring(0, 50),
            Members: members,
            'Last Active': updated,
            ID: chat.id.substring(0, 20) + '...',
        };
    });

    console.table(table);
}

async function exportSingleChat(chatId) {
    console.log(chalk.blue(`📥 Fetching chat ${chatId.substring(0, 20)}...\n`));
    const chats = await listChats();
    const chat = chats.find((c) => c.id === chatId);

    if (!chat) {
        console.log(chalk.red('❌ Chat not found. Use --list to see available chats.'));
        return;
    }

    const outputDir = path.resolve(opts.output);
    const exportOpts = {
        downloadMedia: opts.media !== false,
        maxMessages: opts.max || null,
        format: opts.format,
    };

    const bar = createProgressBar();
    bar.start(100, 0, { status: 'Starting...' });

    const result = await exportChat(chat, outputDir, exportOpts, (stage, value, msg) => {
        bar.update(value, { status: msg });
    });

    bar.stop();
    printExportResult(result);
}

async function exportAllChats() {
    console.log(chalk.blue('📋 Fetching chat list...\n'));
    const chats = await listChats();

    if (chats.length === 0) {
        console.log(chalk.yellow('No chats found.'));
        return;
    }

    const { confirm } = await inquirer.prompt([
        {
            type: 'confirm',
            name: 'confirm',
            message: `Export all ${chats.length} chats? This may take a while.`,
            default: false,
        },
    ]);

    if (!confirm) return;

    const outputDir = path.resolve(opts.output);
    const exportOpts = {
        downloadMedia: opts.media !== false,
        maxMessages: opts.max || null,
        format: opts.format,
    };

    console.log(chalk.blue(`\n📂 Exporting to: ${outputDir}\n`));

    const results = [];
    for (let i = 0; i < chats.length; i++) {
        const chat = chats[i];
        const name = getChatDisplayName(chat);
        console.log(chalk.cyan(`\n[${i + 1}/${chats.length}] ${name}`));

        const bar = createProgressBar();
        bar.start(100, 0, { status: 'Starting...' });

        try {
            const result = await exportChat(chat, outputDir, exportOpts, (stage, value, msg) => {
                bar.update(Math.min(value, 100), { status: msg });
            });
            bar.stop();
            results.push(result);
            console.log(
                chalk.green(
                    `  ✅ ${result.messageCount} messages, ${result.mediaCount} media files`
                )
            );
        } catch (err) {
            bar.stop();
            console.log(chalk.red(`  ❌ Failed: ${err.message}`));
            results.push({ chatName: name, error: err.message });
        }
    }

    // Summary
    console.log(chalk.blue('\n' + '═'.repeat(60)));
    console.log(chalk.green.bold('\n📊 Export Summary:\n'));
    const totalMsgs = results.reduce((sum, r) => sum + (r.messageCount || 0), 0);
    const totalMedia = results.reduce((sum, r) => sum + (r.mediaCount || 0), 0);
    const failures = results.filter((r) => r.error).length;
    console.log(`  Chats exported: ${results.length - failures}/${chats.length}`);
    console.log(`  Total messages: ${totalMsgs}`);
    console.log(`  Total media:    ${totalMedia}`);
    if (failures) console.log(chalk.red(`  Failures:       ${failures}`));
    console.log(`\n  Output: ${outputDir}\n`);
}

async function interactiveMode() {
    console.log(chalk.blue('📋 Fetching chat list...\n'));
    const chats = await listChats();

    if (chats.length === 0) {
        console.log(chalk.yellow('No chats found.'));
        return;
    }

    const choices = chats.map((chat, i) => {
        const name = getChatDisplayName(chat);
        const type = chat.chatType === 'oneOnOne' ? '1:1' : 'Group';
        const updated = new Date(chat.lastUpdatedDateTime).toLocaleDateString();
        return {
            name: `[${type}] ${name.substring(0, 45).padEnd(45)} (${updated})`,
            value: chat.id,
            short: name,
        };
    });

    const { selectedChats } = await inquirer.prompt([
        {
            type: 'checkbox',
            name: 'selectedChats',
            message: 'Select chats to export (Space to select, Enter to confirm):',
            choices,
            pageSize: 20,
        },
    ]);

    if (selectedChats.length === 0) {
        console.log(chalk.yellow('No chats selected.'));
        return;
    }

    const { format, downloadMedia } = await inquirer.prompt([
        {
            type: 'list',
            name: 'format',
            message: 'Export format:',
            choices: [
                { name: 'All formats (HTML + JSON + TXT)', value: 'all' },
                { name: 'HTML (with embedded media)', value: 'html' },
                { name: 'JSON (structured data)', value: 'json' },
                { name: 'Plain Text', value: 'txt' },
            ],
        },
        {
            type: 'confirm',
            name: 'downloadMedia',
            message: 'Download media & attachments?',
            default: true,
        },
    ]);

    const outputDir = path.resolve(opts.output);
    fs.mkdirSync(outputDir, { recursive: true });

    console.log(chalk.blue(`\n📂 Exporting ${selectedChats.length} chats to: ${outputDir}\n`));

    for (let i = 0; i < selectedChats.length; i++) {
        const chatId = selectedChats[i];
        const chat = chats.find((c) => c.id === chatId);
        const name = getChatDisplayName(chat);

        console.log(chalk.cyan(`\n[${i + 1}/${selectedChats.length}] ${name}`));

        const bar = createProgressBar();
        bar.start(100, 0, { status: 'Starting...' });

        try {
            const result = await exportChat(
                chat,
                outputDir,
                { format, downloadMedia, maxMessages: opts.max || null },
                (stage, value, msg) => {
                    bar.update(Math.min(value, 100), { status: msg });
                }
            );
            bar.stop();
            console.log(
                chalk.green(
                    `  ✅ ${result.messageCount} messages, ${result.mediaCount} media files → ${result.outputPath}`
                )
            );
        } catch (err) {
            bar.stop();
            console.log(chalk.red(`  ❌ Failed: ${err.message}`));
        }
    }

    console.log(chalk.green.bold(`\n🎉 Export complete! Files saved to: ${outputDir}\n`));
}

// ================================================================
// HELPERS
// ================================================================

function printBanner() {
    console.log(chalk.blue.bold(`
╔══════════════════════════════════════════╗
║   📋 Teams Chat Graph Exporter v1.0.0   ║
║   Export chats with media via Graph API  ║
╚══════════════════════════════════════════╝
`));
}

function printSetupInstructions() {
    console.log(chalk.white(`
${chalk.bold('Step 1:')} Go to ${chalk.cyan('https://portal.azure.com')}
${chalk.bold('Step 2:')} Navigate to Azure Active Directory → App registrations
${chalk.bold('Step 3:')} Click "New registration"
         - Name: ${chalk.green('Teams Chat Exporter')}
         - Supported account types: ${chalk.green('Accounts in any organizational directory')}
         - Redirect URI: Select ${chalk.green('Mobile and desktop')} → 
           ${chalk.cyan('https://login.microsoftonline.com/common/oauth2/nativeclient')}
${chalk.bold('Step 4:')} Copy the ${chalk.yellow('Application (client) ID')}
${chalk.bold('Step 5:')} Go to API permissions → Add:
         - ${chalk.green('Chat.Read')}
         - ${chalk.green('Chat.ReadBasic')}
         - ${chalk.green('Files.Read')}
         - ${chalk.green('Files.Read.All')}
         - ${chalk.green('User.Read')}
${chalk.bold('Step 6:')} Paste the Client ID in ${chalk.yellow('config.js')}
         → clientId: '${chalk.red('YOUR_CLIENT_ID_HERE')}'
`));
}

function getChatDisplayName(chat) {
    if (chat.topic) return chat.topic;
    const members = (chat.members || []).map((m) => m.displayName).filter(Boolean);
    if (members.length > 0) return members.join(', ');
    return `Chat_${chat.id.substring(0, 8)}`;
}

function createProgressBar() {
    return new cliProgress.SingleBar(
        {
            format: `  ${chalk.blue('{bar}')} {percentage}% | {status}`,
            barCompleteChar: '█',
            barIncompleteChar: '░',
            hideCursor: true,
        },
        cliProgress.Presets.shades_classic
    );
}

function printExportResult(result) {
    console.log(chalk.green.bold('\n✅ Export complete!\n'));
    console.log(`  Chat:     ${result.chatName}`);
    console.log(`  Messages: ${result.messageCount}`);
    console.log(`  Media:    ${result.mediaCount} files`);
    if (result.mediaErrors) console.log(chalk.yellow(`  Errors:   ${result.mediaErrors}`));
    console.log(`  Path:     ${result.outputPath}\n`);
}

// ================================================================
// RUN
// ================================================================

main().catch((err) => {
    console.error(chalk.red(`\n❌ Fatal error: ${err.message}\n`));
    if (process.env.DEBUG) console.error(err.stack);
    process.exit(1);
});
