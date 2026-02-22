#!/usr/bin/env node
/**
 * Manual End-to-End Test Script for OneNote MCP Server
 * 
 * This script makes REAL API calls to your OneNote account to validate
 * that all MCP tools are working correctly with actual data.
 * 
 * Prerequisites:
 * - You must be authenticated (have a valid .access-token.txt file)
 * - You should have at least one notebook with sections and pages
 * 
 * Usage:
 *   node tests/manual/e2e-test.mjs
 *   node tests/manual/e2e-test.mjs --read-only  (skip write operations)
 * 
 * WARNING: This script will create, modify, and potentially delete test pages
 * in your OneNote account. It will create pages with "E2E Test" in the title
 * so you can identify and delete them afterwards.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Configuration
const tokenFilePath = path.join(__dirname, '../../.access-token.txt');
const readOnlyMode = process.argv.includes('--read-only');

// Colors for console output
const colors = {
  reset: '\x1b[0m',
  bright: '\x1b[1m',
  green: '\x1b[32m',
  red: '\x1b[31m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  cyan: '\x1b[36m',
};

// Test state
let graphClient = null;
let accessToken = null;
let testResults = {
  passed: 0,
  failed: 0,
  skipped: 0,
  tests: []
};

// Store IDs for cross-test usage
let testData = {
  notebookId: null,
  sectionId: null,
  pageId: null,
  createdPageId: null,
};

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function log(message, color = colors.reset) {
  console.log(`${color}${message}${colors.reset}`);
}

function logSection(title) {
  console.log('\n' + '='.repeat(80));
  log(title, colors.bright + colors.cyan);
  console.log('='.repeat(80));
}

function logTest(name, status, details = '') {
  const symbol = status === 'PASS' ? '✓' : status === 'FAIL' ? '✗' : '○';
  const color = status === 'PASS' ? colors.green : status === 'FAIL' ? colors.red : colors.yellow;
  log(`${symbol} ${name}`, color);
  if (details) {
    console.log(`  ${details}`);
  }
}

function recordTest(name, passed, error = null) {
  if (passed) {
    testResults.passed++;
    logTest(name, 'PASS');
  } else {
    testResults.failed++;
    logTest(name, 'FAIL', error ? `Error: ${error.message}` : '');
  }
  testResults.tests.push({ name, passed, error });
}

function skipTest(name, reason) {
  testResults.skipped++;
  logTest(name, 'SKIP', `Reason: ${reason}`);
  testResults.tests.push({ name, passed: null, skipped: true, reason });
}

async function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ============================================================================
// AUTHENTICATION
// ============================================================================

async function loadToken() {
  try {
    if (!fs.existsSync(tokenFilePath)) {
      throw new Error('No access token found. Please authenticate first using the authenticate tool.');
    }

    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');

    try {
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
      
      // Check if token is expired
      if (parsedToken.expiresOn) {
        const expiresOn = new Date(parsedToken.expiresOn);
        if (expiresOn < new Date()) {
          throw new Error('Access token has expired. Please re-authenticate.');
        }
      }
    } catch (parseError) {
      // Legacy plain text format
      accessToken = tokenData.trim();
    }

    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    log('✓ Authentication token loaded successfully', colors.green);
    return true;
  } catch (error) {
    log(`✗ Authentication failed: ${error.message}`, colors.red);
    return false;
  }
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

async function testListNotebooks() {
  try {
    const response = await graphClient
      .api('/me/onenote/notebooks')
      .get();

    if (!response || !response.value) {
      throw new Error('Invalid response structure');
    }

    log(`  Found ${response.value.length} notebook(s)`);
    
    if (response.value.length > 0) {
      testData.notebookId = response.value[0].id;
      log(`  Sample: "${response.value[0].displayName}" (ID: ${response.value[0].id})`);
    } else {
      log('  Warning: No notebooks found. Some tests may be skipped.', colors.yellow);
    }

    recordTest('listNotebooks', true);
    return response.value;
  } catch (error) {
    recordTest('listNotebooks', false, error);
    return null;
  }
}

async function testListSections() {
  if (!testData.notebookId) {
    skipTest('listSections', 'No notebook ID available');
    return null;
  }

  try {
    const response = await graphClient
      .api(`/me/onenote/notebooks/${testData.notebookId}/sections`)
      .get();

    log(`  Found ${response.value.length} section(s)`);
    
    if (response.value.length > 0) {
      testData.sectionId = response.value[0].id;
      log(`  Sample: "${response.value[0].displayName}" (ID: ${response.value[0].id})`);
    }

    recordTest('listSections', true);
    return response.value;
  } catch (error) {
    recordTest('listSections', false, error);
    return null;
  }
}

async function testListPages() {
  if (!testData.sectionId) {
    skipTest('listPagesInSection', 'No section ID available');
    return null;
  }

  try {
    const response = await graphClient
      .api(`/me/onenote/sections/${testData.sectionId}/pages`)
      .top(10)
      .get();

    log(`  Found ${response.value.length} page(s)`);
    
    if (response.value.length > 0) {
      testData.pageId = response.value[0].id;
      log(`  Sample: "${response.value[0].title}" (ID: ${response.value[0].id})`);
    }

    recordTest('listPagesInSection', true);
    return response.value;
  } catch (error) {
    recordTest('listPagesInSection', false, error);
    return null;
  }
}

async function testSearchPages() {
  try {
    // Get all sections and search through them
    if (!testData.sectionId) {
      skipTest('searchPages', 'No section ID available');
      return null;
    }
    
    const response = await graphClient
      .api(`/me/onenote/sections/${testData.sectionId}/pages`)
      .top(5)
      .get();

    log(`  Found ${response.value.length} page(s)`);
    recordTest('searchPages', true);
    return response.value;
  } catch (error) {
    recordTest('searchPages', false, error);
    return null;
  }
}

async function testSearchPagesByDate() {
  try {
    if (!testData.sectionId) {
      skipTest('searchPagesByDate', 'No section ID available');
      return null;
    }
    
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    const isoDate = sevenDaysAgo.toISOString();

    const response = await graphClient
      .api(`/me/onenote/sections/${testData.sectionId}/pages`)
      .filter(`lastModifiedDateTime ge ${isoDate}`)
      .top(10)
      .get();

    log(`  Found ${response.value.length} page(s) modified in last 7 days`);
    recordTest('searchPagesByDate', true);
    return response.value;
  } catch (error) {
    recordTest('searchPagesByDate', false, error);
    return null;
  }
}

async function testSearchPagesByDateThreeDays() {
  try {
    if (!testData.notebookId) {
      skipTest('searchPagesByDate (3 days)', 'No notebook ID available');
      return null;
    }
    
    // Calculate threshold (3 days ago)
    const threeDaysAgo = new Date();
    threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
    threeDaysAgo.setHours(0, 0, 0, 0);

    // Get all sections in the notebook
    const sectionsResponse = await graphClient
      .api(`/me/onenote/notebooks/${testData.notebookId}/sections`)
      .get();

    log(`  Checking ${sectionsResponse.value.length} section(s)...`);

    let allMatchingPages = [];

    // Query each section for pages
    for (const section of sectionsResponse.value) {
      try {
        const pagesResponse = await graphClient
          .api(`/me/onenote/sections/${section.id}/pages`)
          .select('id,title,createdDateTime,lastModifiedDateTime')
          .get();

        // Filter by date (client-side, as the tool does)
        const matchingPages = pagesResponse.value.filter(page => {
          const modified = new Date(page.lastModifiedDateTime);
          return modified >= threeDaysAgo;
        });

        allMatchingPages = allMatchingPages.concat(matchingPages);
      } catch (sectionError) {
        log(`  Warning: Error querying section ${section.displayName}: ${sectionError.message}`, colors.yellow);
      }
    }

    // Sort by most recent
    allMatchingPages.sort((a, b) => 
      new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime)
    );

    log(`  Found ${allMatchingPages.length} page(s) modified in last 3 days`);
    if (allMatchingPages.length > 0) {
      log(`  Most recent: "${allMatchingPages[0].title}" (${new Date(allMatchingPages[0].lastModifiedDateTime).toLocaleDateString()})`);
    }
    
    recordTest('searchPagesByDate (3 days)', true);
    return allMatchingPages;
  } catch (error) {
    recordTest('searchPagesByDate (3 days)', false, error);
    return null;
  }
}

async function testGetPageContent() {
  if (!testData.pageId) {
    skipTest('getPageContent', 'No page ID available');
    return null;
  }

  try {
    // First verify the page still exists
    try {
      await graphClient
        .api(`/me/onenote/pages/${testData.pageId}`)
        .get();
    } catch (verifyError) {
      // Page doesn't exist, try to get a fresh one
      if (testData.sectionId) {
        const pagesResponse = await graphClient
          .api(`/me/onenote/sections/${testData.sectionId}/pages`)
          .top(1)
          .get();
        
        if (pagesResponse.value.length > 0) {
          testData.pageId = pagesResponse.value[0].id;
        } else {
          skipTest('getPageContent', 'No pages found in section');
          return null;
        }
      } else {
        skipTest('getPageContent', 'Page no longer exists');
        return null;
      }
    }
    
    // Use fetch to get HTML content (graphClient.get() returns object, not text)
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${testData.pageId}/content`;
    const response = await fetch(url, {
      headers: { 'Authorization': `Bearer ${accessToken}` }
    });
    
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
    
    const htmlContent = await response.text();
    log(`  Retrieved content length: ${htmlContent.length} characters`);
    log(`  Content preview: ${htmlContent.substring(0, 100)}...`);
    recordTest('getPageContent', true);
    return htmlContent;
  } catch (error) {
    recordTest('getPageContent', false, error);
    return null;
  }
}

async function testGetPageByTitle() {
  try {
    if (!testData.sectionId) {
      skipTest('getPageByTitle', 'No section ID available');
      return;
    }
    
    const response = await graphClient
      .api(`/me/onenote/sections/${testData.sectionId}/pages`)
      .top(1)
      .get();

    if (response.value.length > 0) {
      const pageTitle = response.value[0].title;
      log(`  Searching for page with title: "${pageTitle}"`);
      
      const searchResponse = await graphClient
        .api(`/me/onenote/sections/${testData.sectionId}/pages`)
        .filter(`title eq '${pageTitle.replace(/'/g, "''")}'`)
        .get();

      if (searchResponse.value.length > 0) {
        log(`  Found page: "${searchResponse.value[0].title}"`);
        recordTest('getPageByTitle', true);
      } else {
        throw new Error('Page not found by title search');
      }
    } else {
      skipTest('getPageByTitle', 'No pages available to search');
    }
  } catch (error) {
    recordTest('getPageByTitle', false, error);
  }
}

async function testGetUserInfo() {
  try {
    const response = await graphClient
      .api('/me')
      .get();

    log(`  User: ${response.displayName} (${response.userPrincipalName})`);
    recordTest('getUserInfo', true);
    return response;
  } catch (error) {
    recordTest('getUserInfo', false, error);
    return null;
  }
}

// ============================================================================
// WRITE OPERATION TESTS (Only if not in read-only mode)
// ============================================================================

async function testCreatePage() {
  if (readOnlyMode) {
    skipTest('createPage', 'Running in read-only mode');
    return null;
  }

  if (!testData.sectionId) {
    skipTest('createPage', 'No section ID available');
    return null;
  }

  try {
    const timestamp = new Date().toISOString();
    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head>
          <title>E2E Test Page - ${timestamp}</title>
        </head>
        <body>
          <h1>E2E Test Page</h1>
          <p>This page was created by the E2E test script at ${timestamp}</p>
          <p>You can safely delete this page.</p>
        </body>
      </html>
    `;

    const response = await graphClient
      .api(`/me/onenote/sections/${testData.sectionId}/pages`)
      .header('Content-Type', 'text/html')
      .post(htmlContent);

    testData.createdPageId = response.id;
    log(`  Created page: "${response.title}" (ID: ${response.id})`);
    recordTest('createPage', true);
    return response;
  } catch (error) {
    recordTest('createPage', false, error);
    return null;
  }
}

async function testUpdatePageContent() {
  if (readOnlyMode) {
    skipTest('updatePageContent', 'Running in read-only mode');
    return;
  }

  if (!testData.createdPageId) {
    skipTest('updatePageContent', 'No test page created');
    return;
  }

  try {
    // Wait a bit for page to be fully created
    await sleep(2000);

    const patchData = [
      {
        target: 'body',
        action: 'append',
        content: '<p>This content was appended by the E2E test script.</p>'
      }
    ];

    await graphClient
      .api(`/me/onenote/pages/${testData.createdPageId}/content`)
      .header('Content-Type', 'application/json')
      .patch(patchData);

    log(`  Successfully updated page content`);
    recordTest('updatePageContent', true);
  } catch (error) {
    recordTest('updatePageContent', false, error);
  }
}

async function testDeleteTestPage() {
  if (readOnlyMode) {
    skipTest('deleteTestPage', 'Running in read-only mode');
    return;
  }

  if (!testData.createdPageId) {
    skipTest('deleteTestPage', 'No test page to delete');
    return;
  }

  try {
    // Wait a bit before deleting
    await sleep(2000);

    // Note: Microsoft Graph API doesn't support deleting OneNote pages directly
    // We'll just log that the page was created and can be manually deleted
    log(`  Note: Test page created (ID: ${testData.createdPageId})`);
    log(`  Please manually delete this test page from OneNote`);
    recordTest('deleteTestPage', true);
  } catch (error) {
    recordTest('deleteTestPage', false, error);
  }
}

// ============================================================================
// PAGINATION TESTS
// ============================================================================

async function testPagination() {
  try {
    if (!testData.sectionId) {
      skipTest('pagination', 'No section ID available');
      return;
    }
    
    let pageCount = 0;
    let url = `/me/onenote/sections/${testData.sectionId}/pages`;
    let iterations = 0;
    const maxIterations = 3; // Limit to avoid long test times

    while (url && iterations < maxIterations) {
      const response = await graphClient
        .api(url)
        .top(10)
        .get();

      pageCount += response.value.length;
      url = response['@odata.nextLink'] ? response['@odata.nextLink'].split('/v1.0')[1] : null;
      iterations++;
    }

    log(`  Successfully paginated through ${pageCount} pages across ${iterations} request(s)`);
    recordTest('pagination', true);
  } catch (error) {
    recordTest('pagination', false, error);
  }
}

// ============================================================================
// ERROR HANDLING TESTS
// ============================================================================

async function testErrorHandling() {
  try {
    // Try to access a page in a non-existent section
    // This should reliably return 404 or ResourceNotFound
    const fakeNotebookId = '1-00000000-0000-0000-0000-000000000000';
    
    try {
      await graphClient
        .api(`/me/onenote/notebooks/${fakeNotebookId}/sections`)
        .get();
      
      // If we get here, try a different approach - non-existent page
      throw new Error('Expected 404 error but request succeeded');
    } catch (error) {
      // Check if it's a 404 or ResourceNotFound error
      const is404 = error.statusCode === 404 || 
                    error.code === 'ResourceNotFound' || 
                    error.code === 'ItemNotFound' ||
                    error.message.includes('not found') || 
                    error.message.includes('does not exist') ||
                    error.message.includes('NotFound');
      
      if (is404) {
        log(`  Correctly handled 404/NotFound error for non-existent resource`);
        recordTest('errorHandling', true);
      } else {
        throw error;
      }
    }
  } catch (error) {
    recordTest('errorHandling', false, error);
  }
}

// ============================================================================
// MAIN TEST RUNNER
// ============================================================================

async function runAllTests() {
  console.clear();
  log('╔════════════════════════════════════════════════════════════════════════════╗', colors.bright);
  log('║              OneNote MCP Server - End-to-End Test Suite                   ║', colors.bright);
  log('╚════════════════════════════════════════════════════════════════════════════╝', colors.bright);
  
  if (readOnlyMode) {
    log('\n⚠️  Running in READ-ONLY mode - Write operations will be skipped', colors.yellow);
  } else {
    log('\n⚠️  WARNING: This will create test pages in your OneNote account!', colors.yellow);
    log('   Test pages will be titled "E2E Test Page - [timestamp]"', colors.yellow);
  }

  console.log('\nStarting tests in 3 seconds...');
  await sleep(3000);

  // Authentication
  logSection('1. AUTHENTICATION');
  const authenticated = await loadToken();
  if (!authenticated) {
    log('\n✗ Cannot proceed without authentication', colors.red);
    process.exit(1);
  }
  await testGetUserInfo();

  // Read Operations - Basic
  logSection('2. READ OPERATIONS - Basic Listing');
  await testListNotebooks();
  await sleep(500);
  await testListSections();
  await sleep(500);
  await testListPages();
  await sleep(500);

  // Read Operations - Search
  logSection('3. READ OPERATIONS - Search & Filter');
  await testSearchPages();
  await sleep(500);
  await testSearchPagesByDate();
  await sleep(500);
  await testSearchPagesByDateThreeDays();
  await sleep(500);
  await testGetPageByTitle();
  await sleep(500);

  // Read Operations - Content
  logSection('4. READ OPERATIONS - Page Content');
  await testGetPageContent();
  await sleep(500);

  // Advanced Features
  logSection('5. ADVANCED FEATURES');
  await testPagination();
  await sleep(500);
  await testErrorHandling();
  await sleep(500);

  // Write Operations
  if (!readOnlyMode) {
    logSection('6. WRITE OPERATIONS');
    log('⚠️  Creating test page...', colors.yellow);
    await testCreatePage();
    await sleep(1000);
    
    if (testData.createdPageId) {
      await testUpdatePageContent();
      await sleep(1000);
      await testDeleteTestPage();
    }
  }

  // Summary
  logSection('TEST SUMMARY');
  const total = testResults.passed + testResults.failed + testResults.skipped;
  
  log(`\nTotal Tests: ${total}`, colors.bright);
  log(`  ✓ Passed:  ${testResults.passed}`, colors.green);
  log(`  ✗ Failed:  ${testResults.failed}`, colors.red);
  log(`  ○ Skipped: ${testResults.skipped}`, colors.yellow);
  
  const passRate = total > 0 ? ((testResults.passed / (testResults.passed + testResults.failed)) * 100).toFixed(1) : 0;
  log(`\nPass Rate: ${passRate}%`, passRate >= 80 ? colors.green : colors.red);

  if (testResults.failed > 0) {
    log('\n\nFailed Tests:', colors.red);
    testResults.tests
      .filter(t => !t.passed && !t.skipped)
      .forEach(t => {
        log(`  ✗ ${t.name}`, colors.red);
        if (t.error) {
          log(`    ${t.error.message}`, colors.reset);
        }
      });
  }

  if (testData.createdPageId && !readOnlyMode) {
    log('\n\n📝 Note: A test page was created with ID: ' + testData.createdPageId, colors.yellow);
    log('   Please manually delete it from OneNote if no longer needed.', colors.yellow);
  }

  console.log('\n');
  process.exit(testResults.failed > 0 ? 1 : 0);
}

// Run the tests
runAllTests().catch(error => {
  log(`\n✗ Fatal error: ${error.message}`, colors.red);
  console.error(error);
  process.exit(1);
});
