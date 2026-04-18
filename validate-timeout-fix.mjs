#!/usr/bin/env node
/**
 * Timeout Fix Validation Script
 * 
 * Tests the fixes for getPageContent and getPageByTitle timeout issues.
 * Uses the same MCP stdio protocol approach as test-mcp-tool.mjs.
 * 
 * Usage:
 *   node validate-timeout-fix.mjs
 */

import { spawn } from 'child_process';

const colors = {
  reset: '\x1b[0m',
  green: '\x1b[32m',
  red: '\x1b[31m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  cyan: '\x1b[36m',
  bold: '\x1b[1m',
};

function log(msg, color = colors.reset) {
  console.log(`${color}${msg}${colors.reset}`);
}

/**
 * Call an MCP tool by spawning the server, sending initialize + tool call.
 * Returns { text, elapsed } on success.
 */
function callTool(toolName, args = {}, timeoutMs = 120000) {
  return new Promise((resolve, reject) => {
    const server = spawn('node', ['onenote-mcp.mjs'], {
      stdio: ['pipe', 'pipe', 'pipe']
    });

    let buffer = '';
    const startTime = Date.now();
    let resolved = false;

    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true;
        server.kill();
        reject(new Error(`Timed out after ${timeoutMs / 1000}s`));
      }
    }, timeoutMs);

    server.stdout.on('data', (data) => {
      buffer += data.toString();
      const lines = buffer.split('\n');
      buffer = lines.pop();

      for (const line of lines) {
        if (!line.trim()) continue;
        try {
          const msg = JSON.parse(line);
          // Wait for the tool call response (id === 2)
          if (msg.id === 2 && !resolved) {
            resolved = true;
            clearTimeout(timer);
            const elapsed = Date.now() - startTime;
            server.kill();

            if (msg.error) {
              reject(new Error(msg.error.message || JSON.stringify(msg.error)));
            } else {
              const text = msg.result?.content?.[0]?.text || '';
              resolve({ text, elapsed });
            }
          }
        } catch (e) {
          // not JSON yet
        }
      }
    });

    server.on('error', (err) => {
      if (!resolved) { resolved = true; clearTimeout(timer); reject(err); }
    });

    // Send initialize (id=1), then tool call (id=2) after 2s
    setTimeout(() => {
      server.stdin.write(JSON.stringify({
        jsonrpc: '2.0', id: 1, method: 'initialize',
        params: {
          protocolVersion: '2024-11-05',
          capabilities: {},
          clientInfo: { name: 'validator', version: '1.0.0' }
        }
      }) + '\n');

      setTimeout(() => {
        server.stdin.write(JSON.stringify({
          jsonrpc: '2.0', id: 2, method: 'tools/call',
          params: { name: toolName, arguments: args }
        }) + '\n');
      }, 2000);
    }, 100);
  });
}

async function runTest(name, fn, maxTimeMs) {
  process.stdout.write(`${colors.blue}▶${colors.reset} ${name}... `);
  try {
    const result = await fn();
    const secs = (result.elapsed / 1000).toFixed(1);
    if (maxTimeMs && result.elapsed > maxTimeMs) {
      log(`⚠️  SLOW (${secs}s > ${maxTimeMs / 1000}s target)`, colors.yellow);
      return { status: 'warn', ...result };
    }
    log(`✅ PASS (${secs}s)`, colors.green);
    return { status: 'pass', ...result };
  } catch (err) {
    log(`❌ FAIL: ${err.message}`, colors.red);
    return { status: 'fail', error: err.message };
  }
}

// ── Main ──────────────────────────────────────────────────────────

async function main() {
  log('\n' + '═'.repeat(70), colors.cyan);
  log('  TIMEOUT FIX VALIDATION', colors.bold + colors.cyan);
  log('═'.repeat(70) + '\n', colors.cyan);

  let passed = 0, failed = 0, warned = 0;
  const count = (r) => { if (r.status === 'pass') passed++; else if (r.status === 'fail') failed++; else warned++; };

  // ── Step 1: Get a test page ID via searchPagesByDate ──
  let testPageId = null;
  let testPageTitle = null;

  const r1 = await runTest(
    'searchPagesByDate returns page IDs',
    async () => {
      const { text, elapsed } = await callTool('searchPagesByDate', { days: 7 }, 120000);
      if (!text.includes('ID:')) throw new Error('Page IDs not in results');

      const idMatch = text.match(/^\d+\.\s+\*\*[^*]+\*\*\n\s+ID:\s*(.+)/m);
      const titleMatch = text.match(/^\d+\.\s+\*\*([^*]+)\*\*/m);
      testPageId = idMatch?.[1]?.trim();
      testPageTitle = titleMatch?.[1];

      if (!testPageId) throw new Error('Could not extract page ID');
      log(`\n   Test page: "${testPageTitle}" (ID: ${testPageId.substring(0, 30)}...)`, colors.cyan);
      return { text, elapsed };
    },
    60000
  );
  count(r1);

  if (!testPageId) {
    log('\n❌ Cannot continue without a page ID.', colors.red);
    process.exit(1);
  }

  // ── Step 2: getPageContent (text) ──
  count(await runTest(
    'getPageContent (text)',
    async () => await callTool('getPageContent', { pageId: testPageId, format: 'text' }),
    30000
  ));

  // ── Step 3: getPageContent (summary) ──
  count(await runTest(
    'getPageContent (summary)',
    async () => await callTool('getPageContent', { pageId: testPageId, format: 'summary' }),
    10000
  ));

  // ── Step 4: getPageContent (html) ──
  count(await runTest(
    'getPageContent (html)',
    async () => await callTool('getPageContent', { pageId: testPageId, format: 'html' }),
    30000
  ));

  // ── Step 5: getPageByTitle ──
  if (testPageTitle) {
    count(await runTest(
      `getPageByTitle ("${testPageTitle.substring(0, 30)}")`,
      async () => {
        const result = await callTool('getPageByTitle', { title: testPageTitle, format: 'summary' }, 120000);
        if (result.text.includes('timed out')) throw new Error('Content fetch timed out');
        return result;
      },
      15000
    ));
  }

  // ── Step 6: searchPageContent ──
  count(await runTest(
    'searchPageContent returns page IDs',
    async () => {
      const result = await callTool('searchPageContent', { query: 'the', days: 7, maxPages: 3 }, 120000);
      if (!result.text.includes('No pages found') && !result.text.includes('0 matches') && !result.text.includes('ID:')) {
        throw new Error('Page IDs not in results');
      }
      return result;
    },
    30000
  ));

  // ── Step 7: listNotebooks regression ──
  count(await runTest(
    'Regression: listNotebooks',
    async () => await callTool('listNotebooks', { refresh: false }),
    5000
  ));

  // ── Summary ──
  log('\n' + '═'.repeat(70), colors.cyan);
  log('  RESULTS', colors.bold + colors.cyan);
  log('═'.repeat(70), colors.cyan);
  log(`\n  ✅ Passed:   ${passed}`, colors.green);
  log(`  ❌ Failed:   ${failed}`, colors.red);
  log(`  ⚠️  Warnings: ${warned}`, colors.yellow);

  const total = passed + failed;
  log(`\n  Pass Rate: ${total > 0 ? ((passed / total) * 100).toFixed(0) : 0}%`, colors.bold);

  if (failed === 0) {
    log('\n  🎉 All tests passed!', colors.green + colors.bold);
  }
  log('\n' + '═'.repeat(70) + '\n', colors.cyan);
  process.exit(failed > 0 ? 1 : 0);
}

main().catch(err => {
  console.error(`Fatal: ${err.message}`);
  process.exit(1);
});
