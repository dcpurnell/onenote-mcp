#!/usr/bin/env node
/**
 * Timeout Fix Validation Script
 * 
 * Tests the fixes for getPageContent and getPageByTitle timeout issues.
 * Validates that:
 * - Tools complete within expected timeframes
 * - Page IDs are included in search results
 * - No regression in other tools
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

class TestRunner {
  constructor() {
    this.results = {
      passed: 0,
      failed: 0,
      warnings: 0,
    };
    this.testPageId = null;
    this.testPageTitle = null;
  }

  log(message, color = colors.reset) {
    console.log(`${color}${message}${colors.reset}`);
  }

  async callTool(toolName, args = {}) {
    return new Promise((resolve, reject) => {
      const server = spawn('node', ['onenote-mcp.mjs'], {
        stdio: ['pipe', 'pipe', 'pipe']
      });

      let buffer = '';
      let errorBuffer = '';
      let messageId = 1;
      const startTime = Date.now();

      server.stdout.on('data', (data) => {
        buffer += data.toString();
        
        const lines = buffer.split('\n');
        buffer = lines.pop();
        
        lines.forEach(line => {
          if (line.trim()) {
            try {
              const response = JSON.parse(line);
              if (response.id === messageId) {
                const elapsed = Date.now() - startTime;
                server.kill();
                resolve({ response, elapsed });
              }
            } catch (e) {
              // Not JSON, ignore
            }
          }
        });
      });

      server.stderr.on('data', (data) => {
        errorBuffer += data.toString();
      });

      server.on('error', (error) => {
        reject(new Error(`Server error: ${error.message}`));
      });

      server.on('close', (code) => {
        if (code !== 0 && code !== null) {
          reject(new Error(`Server exited with code ${code}\nStderr: ${errorBuffer}`));
        }
      });

      // Send initialize request
      const initRequest = {
        jsonrpc: '2.0',
        id: messageId,
        method: 'initialize',
        params: {
          protocolVersion: '2024-11-05',
          capabilities: {},
          clientInfo: { name: 'validation-script', version: '1.0.0' }
        }
      };
      server.stdin.write(JSON.stringify(initRequest) + '\n');

      // Wait a bit then send tool call
      setTimeout(() => {
        messageId++;
        const toolRequest = {
          jsonrpc: '2.0',
          id: messageId,
          method: 'tools/call',
          params: {
            name: toolName,
            arguments: args
          }
        };
        server.stdin.write(JSON.stringify(toolRequest) + '\n');
      }, 500);

      // Timeout after 65 seconds (should never hit this if fix works)
      setTimeout(() => {
        server.kill();
        reject(new Error(`Test timed out after 65 seconds`));
      }, 65000);
    });
  }

  async test(name, testFn, expectedMaxTime = null) {
    process.stdout.write(`\n${colors.blue}▶${colors.reset} ${name}... `);
    
    try {
      const result = await testFn();
      
      if (expectedMaxTime && result.elapsed > expectedMaxTime) {
        this.log(`⚠️  SLOW (${(result.elapsed / 1000).toFixed(1)}s > ${expectedMaxTime / 1000}s expected)`, colors.yellow);
        this.results.warnings++;
      } else {
        this.log(`✅ PASS (${(result.elapsed / 1000).toFixed(1)}s)`, colors.green);
        this.results.passed++;
      }
      
      return result;
    } catch (error) {
      this.log(`❌ FAIL: ${error.message}`, colors.red);
      this.results.failed++;
      return null;
    }
  }

  extractPageId(responseText) {
    // Match ID from a numbered list item (e.g. "1. **Title**\n   ID: abc123")
    const idMatch = responseText.match(/^\d+\.\s+\*\*[^*]+\*\*\n\s+ID:\s*(.+)/m);
    return idMatch ? idMatch[1].trim() : null;
  }

  extractPageTitle(responseText) {
    // Match title from a numbered list item (e.g. "1. **My Page Title**")
    const titleMatch = responseText.match(/^\d+\.\s+\*\*([^*]+)\*\*/m);
    return titleMatch ? titleMatch[1] : null;
  }

  async run() {
    this.log('\n' + '═'.repeat(70), colors.cyan);
    this.log('  TIMEOUT FIX VALIDATION SCRIPT', colors.bold + colors.cyan);
    this.log('═'.repeat(70) + '\n', colors.cyan);

    // Test 1: Search pages by date (should return page IDs now)
    const searchResult = await this.test(
      'searchPagesByDate returns page IDs',
      async () => {
        const result = await this.callTool('searchPagesByDate', { days: 7 });
        
        if (result.response.error) {
          throw new Error(result.response.error.message);
        }
        
        const content = result.response.result?.content?.[0]?.text || '';
        
        // Check if page IDs are present
        if (!content.includes('ID:')) {
          throw new Error('Page IDs not found in search results');
        }
        
        // Extract first page ID and title for later tests
        this.testPageId = this.extractPageId(content);
        this.testPageTitle = this.extractPageTitle(content);
        
        if (!this.testPageId) {
          throw new Error('Could not extract page ID from results');
        }
        
        this.log(`\n   Found test page: "${this.testPageTitle}" (ID: ${this.testPageId.substring(0, 20)}...)`, colors.cyan);
        
        return result;
      },
      5000 // Should complete in < 5 seconds
    );

    if (!this.testPageId) {
      this.log('\n❌ Cannot continue without a valid page ID. Please ensure you have pages in OneNote.', colors.red);
      return this.printSummary();
    }

    // Test 2: getPageContent with text format
    await this.test(
      'getPageContent (format: text) completes quickly',
      async () => {
        const result = await this.callTool('getPageContent', {
          pageId: this.testPageId,
          format: 'text'
        });
        
        if (result.response.error) {
          throw new Error(result.response.error.message);
        }
        
        const content = result.response.result?.content?.[0]?.text || '';
        if (content.includes('timed out')) {
          throw new Error('Request timed out');
        }
        
        return result;
      },
      30000 // Should complete in < 30 seconds
    );

    // Test 3: getPageContent with summary format
    await this.test(
      'getPageContent (format: summary) completes quickly',
      async () => {
        const result = await this.callTool('getPageContent', {
          pageId: this.testPageId,
          format: 'summary'
        });
        
        if (result.response.error) {
          throw new Error(result.response.error.message);
        }
        
        const content = result.response.result?.content?.[0]?.text || '';
        if (content.includes('timed out')) {
          throw new Error('Request timed out');
        }
        
        return result;
      },
      10000 // Should complete in < 10 seconds
    );

    // Test 4: getPageContent with html format
    await this.test(
      'getPageContent (format: html) completes quickly',
      async () => {
        const result = await this.callTool('getPageContent', {
          pageId: this.testPageId,
          format: 'html'
        });
        
        if (result.response.error) {
          throw new Error(result.response.error.message);
        }
        
        const content = result.response.result?.content?.[0]?.text || '';
        if (content.includes('timed out')) {
          throw new Error('Request timed out');
        }
        
        return result;
      },
      30000 // Should complete in < 30 seconds
    );

    // Test 5: getPageByTitle
    if (this.testPageTitle) {
      await this.test(
        'getPageByTitle completes quickly',
        async () => {
          const result = await this.callTool('getPageByTitle', {
            title: this.testPageTitle,
            format: 'summary'
          });
          
          if (result.response.error) {
            throw new Error(result.response.error.message);
          }
          
          const content = result.response.result?.content?.[0]?.text || '';
          if (content.includes('timed out')) {
            throw new Error('Request timed out');
          }
          
          return result;
        },
        15000 // Should complete in < 15 seconds
      );
    }

    // Test 6: searchPageContent returns page IDs
    await this.test(
      'searchPageContent returns page IDs',
      async () => {
        const result = await this.callTool('searchPageContent', {
          query: 'the',
          days: 7,
          maxPages: 3
        });
        
        if (result.response.error) {
          throw new Error(result.response.error.message);
        }
        
        const content = result.response.result?.content?.[0]?.text || '';
        
        // Check if page IDs are present (if any results found)
        if (content.includes('No pages found') || content.includes('0 matches')) {
          this.log(`\n   No search results found (query may be too narrow), skipping ID check`, colors.yellow);
        } else if (!content.includes('ID:')) {
          throw new Error('Page IDs not found in search results');
        }
        
        return result;
      },
      10000 // Should complete in < 10 seconds
    );

    // Test 7: Regression test - listNotebooks still works
    await this.test(
      'Regression: listNotebooks still works',
      async () => {
        const result = await this.callTool('listNotebooks', {
          refresh: false
        });
        
        if (result.response.error) {
          throw new Error(result.response.error.message);
        }
        
        return result;
      },
      5000 // Should complete in < 5 seconds
    );

    this.printSummary();
  }

  printSummary() {
    this.log('\n' + '═'.repeat(70), colors.cyan);
    this.log('  TEST SUMMARY', colors.bold + colors.cyan);
    this.log('═'.repeat(70), colors.cyan);
    
    this.log(`\n✅ Passed:   ${this.results.passed}`, colors.green);
    this.log(`❌ Failed:   ${this.results.failed}`, colors.red);
    this.log(`⚠️  Warnings: ${this.results.warnings}`, colors.yellow);
    
    const total = this.results.passed + this.results.failed;
    const passRate = total > 0 ? ((this.results.passed / total) * 100).toFixed(1) : 0;
    
    this.log(`\nPass Rate: ${passRate}%`, colors.bold);
    
    if (this.results.failed === 0 && this.results.warnings === 0) {
      this.log('\n🎉 All tests passed! The timeout fix is working correctly.', colors.green + colors.bold);
      this.log('\nNext steps:', colors.cyan);
      this.log('  1. Merge the optimization branch to main', colors.reset);
      this.log('  2. Test with your downstream Claude workflow', colors.reset);
    } else if (this.results.failed === 0) {
      this.log('\n⚠️  All tests passed but some were slower than expected.', colors.yellow + colors.bold);
      this.log('Consider investigating the slow operations.', colors.yellow);
    } else {
      this.log('\n❌ Some tests failed. Please review the errors above.', colors.red + colors.bold);
    }
    
    this.log('\n' + '═'.repeat(70) + '\n', colors.cyan);
    
    process.exit(this.results.failed > 0 ? 1 : 0);
  }
}

// Run the validation
const runner = new TestRunner();
runner.run().catch(error => {
  console.error(`${colors.red}Fatal error: ${error.message}${colors.reset}`);
  process.exit(1);
});
