#!/usr/bin/env node
/**
 * Simple MCP Tool Tester
 * 
 * Usage:
 *   node test-mcp-tool.mjs listNotebooks
 *   node test-mcp-tool.mjs searchPages '{"query": "test"}'
 *   node test-mcp-tool.mjs getPageContent '{"pageId": "xxx", "format": "text"}'
 */

import { spawn } from 'child_process';

const toolName = process.argv[2];
const argsJson = process.argv[3] || '{}';

if (!toolName) {
  console.error('Usage: node test-mcp-tool.mjs <toolName> [arguments-json]');
  console.error('Example: node test-mcp-tool.mjs listNotebooks');
  console.error('Example: node test-mcp-tool.mjs searchPages \'{"query": "test"}\'');
  process.exit(1);
}

let args;
try {
  args = JSON.parse(argsJson);
} catch (error) {
  console.error('Invalid JSON arguments:', argsJson);
  process.exit(1);
}

console.log(`\n🔧 Testing tool: ${toolName}`);
console.log(`📥 Arguments:`, JSON.stringify(args, null, 2));
console.log('─'.repeat(60));

const server = spawn('node', ['onenote-mcp.mjs'], {
  stdio: ['pipe', 'pipe', 'inherit']
});

let messageId = 1;
let buffer = '';

server.stdout.on('data', (data) => {
  buffer += data.toString();
  
  // Try to parse complete JSON-RPC messages
  const lines = buffer.split('\n');
  buffer = lines.pop(); // Keep incomplete line in buffer
  
  lines.forEach(line => {
    if (line.trim()) {
      try {
        const message = JSON.parse(line);
        if (message.result) {
          console.log('\n✅ Result:');
          console.log(JSON.stringify(message.result, null, 2));
          
          // If it's an error result, show details
          if (message.result.isError) {
            console.log('\n⚠️  This is an error response from the tool');
          }
        } else if (message.error) {
          console.log('\n❌ JSON-RPC Error:');
          console.log(JSON.stringify(message.error, null, 2));
        }
      } catch (e) {
        // Ignore parse errors for incomplete messages
      }
    }
  });
});

server.on('close', (code) => {
  process.exit(code);
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
      clientInfo: { name: 'test-client', version: '1.0.0' }
    }
  };
  
  server.stdin.write(JSON.stringify(initRequest) + '\n');
  
  // Wait for server to fully initialize (increased from 500ms to 2000ms)
  setTimeout(() => {
    const toolRequest = {
      jsonrpc: '2.0',
      id: messageId++,
      method: 'tools/call',
      params: {
        name: toolName,
        arguments: args
      }
    };
    
    console.log('\n📤 Sending request...\n');
    server.stdin.write(JSON.stringify(toolRequest) + '\n');
    
    // Give it time to respond, then close
    setTimeout(() => {
      server.stdin.end();
    }, 10000);
  }, 2000);
}, 100);
