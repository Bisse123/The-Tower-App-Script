#!/usr/bin/env node

// Simple wrapper to pass arguments to PowerShell script
const { spawn } = require('child_process');
const path = require('path');

// Get the version argument
const version = process.argv[2];

if (!version) {
    console.log('Usage: npm run archive <version>');
    console.log('Example: npm run archive 35');
    process.exit(1);
}

// Get optional environment argument (defaults to dev)
const environment = process.argv[3] || 'dev';

// Run the PowerShell script
const scriptPath = path.join(__dirname, 'archive-deployments.ps1');
const ps = spawn('powershell', [
    '-ExecutionPolicy', 'Bypass',
    '-File', scriptPath,
    '-MaxVersion', version,
    '-Environment', environment
], {
    stdio: 'inherit'
});

ps.on('exit', (code) => {
    process.exit(code);
});
