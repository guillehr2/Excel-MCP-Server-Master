#!/usr/bin/env node
/*
 * Excel MCP Server - NPM Wrapper
 * Author: Guillem Hermida (https://github.com/guillehr2)
 * License: MIT
 */

const spawn = require('cross-spawn');
const path = require('path');
const fs = require('fs');
const os = require('os');

// Path to the Python script
const pythonScript = path.join(__dirname, 'master_excel_mcp.py');

// Check if Python script exists
if (!fs.existsSync(pythonScript)) {
    console.error('Error: master_excel_mcp.py not found');
    process.exit(1);
}

// Determine Python command based on OS
const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';

// Check if uv is installed, otherwise use pip
const checkCommand = (cmd) => {
    try {
        const result = spawn.sync(cmd, ['--version'], { stdio: 'pipe' });
        return result.status === 0;
    } catch (e) {
        return false;
    }
};

const hasUv = checkCommand('uv');

// Function to install dependencies silently
const installDependencies = () => {
    const deps = [
        'fastmcp',
        'openpyxl',
        'pandas',
        'numpy',
        'xlsxwriter',
        'xlrd',
        'xlwt',
        'matplotlib'
    ];
    
    if (hasUv) {
        // Use uv for faster installation with --system flag
        const result = spawn.sync('uv', ['pip', 'install', '--system', '--quiet', ...deps], {
            stdio: 'ignore'
        });
        return result.status === 0;
    } else {
        // Fallback to pip
        const pipCmd = process.platform === 'win32' ? 'pip' : 'pip3';
        const result = spawn.sync(pipCmd, ['install', '--quiet', ...deps], {
            stdio: 'ignore'
        });
        return result.status === 0;
    }
};

// Check if this is the first run by checking for a marker file
const markerFile = path.join(os.homedir(), '.excel-mcp-server-installed-v3');

if (!fs.existsSync(markerFile)) {
    // Install dependencies silently
    if (installDependencies()) {
        // Create marker file
        fs.writeFileSync(markerFile, new Date().toISOString());
    }
}

// Run the Python script directly without extra logging
const args = process.argv.slice(2);

if (hasUv) {
    // Use uv run for better dependency management (without --system flag)
    const uvArgs = [
        'run',
        '--with', 'matplotlib',
        '--with', 'mcp[cli]',
        '--with', 'numpy',
        '--with', 'openpyxl',
        '--with', 'pandas',
        '--with', 'xlsxwriter',
        '--with', 'xlrd',
        '--with', 'xlwt',
        'mcp',
        'run',
        pythonScript,
        ...args
    ];
    
    const child = spawn('uv', uvArgs, {
        stdio: 'inherit'
    });
    
    child.on('error', (err) => {
        console.error('Failed to start server:', err);
        process.exit(1);
    });
    
    child.on('exit', (code) => {
        process.exit(code || 0);
    });
} else {
    // Fallback to direct Python execution
    const child = spawn(pythonCmd, [pythonScript, ...args], {
        stdio: 'inherit'
    });
    
    child.on('error', (err) => {
        console.error('Failed to start server:', err);
        process.exit(1);
    });
    
    child.on('exit', (code) => {
        process.exit(code || 0);
    });
}
