#!/usr/bin/env node
/**
 * Cache Refresh Utility
 * 
 * This script triggers a full refresh of the notebook cache,
 * populating .notebook-cache.json with all personal and team notebooks.
 * 
 * Usage:
 *   node refresh-cache.mjs
 * 
 * This is useful for:
 * - Initial setup to populate cache for other apps using this MCP server
 * - Manual cache refresh when you know notebooks have changed
 * - Testing cache persistence
 */

import { spawn } from 'child_process';

console.log('\n📚 OneNote Cache Refresh Utility');
console.log('═'.repeat(60));
console.log('This will fetch all notebooks (personal + team) and save to cache.');
console.log('Please wait, this may take 30-60 seconds for large organizations...\n');

const server = spawn('node', ['onenote-mcp.mjs'], {
  stdio: ['pipe', 'pipe', 'inherit']
});

let messageId = 1;
let buffer = '';
let startTime = Date.now();

server.stdout.on('data', (data) => {
  buffer += data.toString();
  
  const lines = buffer.split('\n');
  buffer = lines.pop();
  
  lines.forEach(line => {
    if (line.trim()) {
      try {
        const message = JSON.parse(line);
        
        if (message.result) {
          const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
          
          if (message.result.content) {
            const content = message.result.content[0]?.text || '';
            
            // Parse the notebook count from the response
            const match = content.match(/(\d+) notebooks?/);
            const notebookCount = match ? match[1] : 'unknown';
            
            console.log('\n✅ Cache refresh complete!');
            console.log(`⏱️  Time elapsed: ${elapsed}s`);
            console.log(`📚 Notebooks cached: ${notebookCount}`);
            console.log(`💾 Cache file: .notebook-cache.json`);
            console.log('\nThe cache is now available for all apps using this MCP server.\n');
          } else {
            console.log('\n✅ Cache refresh result:');
            console.log(JSON.stringify(message.result, null, 2));
          }
          
          setTimeout(() => {
            server.kill();
            process.exit(0);
          }, 100);
        } else if (message.error) {
          console.log('\n❌ Error refreshing cache:');
          console.log(JSON.stringify(message.error, null, 2));
          setTimeout(() => {
            server.kill();
            process.exit(1);
          }, 100);
        }
      } catch (e) {
        // Ignore parse errors
      }
    }
  });
});

server.on('close', (code) => {
  if (code !== 0) {
    console.error(`\n❌ Server exited with code ${code}`);
  }
  process.exit(code);
});

server.on('error', (err) => {
  console.error(`\n❌ Failed to start server: ${err.message}`);
  process.exit(1);
});

// Send initialization
setTimeout(() => {
  const initRequest = {
    jsonrpc: '2.0',
    id: messageId++,
    method: 'initialize',
    params: {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: { name: 'cache-refresh-utility', version: '1.0.0' }
    }
  };
  
  server.stdin.write(JSON.stringify(initRequest) + '\n');
}, 100);

// Listen for initialization response, then send notifications/initialized
let initComplete = false;
const originalStdoutHandler = server.stdout.on;

server.stdout.on('data', (data) => {
  if (!initComplete) {
    const str = data.toString();
    if (str.includes('"method":"initialize"') || str.includes('serverInfo')) {
      initComplete = true;
      
      // Send notifications/initialized
      setTimeout(() => {
        const notifRequest = {
          jsonrpc: '2.0',
          method: 'notifications/initialized'
        };
        server.stdin.write(JSON.stringify(notifRequest) + '\n');
        
        // Now send the actual tool call
        setTimeout(() => {
          console.log('🔄 Requesting full notebook list (this triggers cache refresh)...\n');
          
          const toolRequest = {
            jsonrpc: '2.0',
            id: messageId++,
            method: 'tools/call',
            params: {
              name: 'listNotebooks',
              arguments: {
                includeTeamNotebooks: true
              }
            }
          };
          
          server.stdin.write(JSON.stringify(toolRequest) + '\n');
        }, 500);
      }, 500);
    }
  }
});

// Timeout after 5 minutes
setTimeout(() => {
  console.error('\n⏱️  Timeout: Cache refresh took too long (>5 minutes)');
  console.error('This may indicate an issue with Microsoft Graph API or network.');
  server.kill();
  process.exit(1);
}, 5 * 60 * 1000);
