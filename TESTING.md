# Testing Guide for OneNote MCP Server

This document describes the testing strategy, setup, and best practices for the OneNote MCP Server project.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Quick Start](#quick-start)
- [Test Structure](#test-structure)
- [Interactive Testing with MCP Inspector](#interactive-testing-with-mcp-inspector)
- [Running Tests](#running-tests)
- [Writing Tests](#writing-tests)
- [Continuous Integration](#continuous-integration)
- [Coverage Requirements](#coverage-requirements)
- [Troubleshooting](#troubleshooting)

## Overview

The OneNote MCP Server uses a comprehensive testing strategy to ensure code quality and reliability:

- **Unit Tests**: Test individual functions in isolation (HTML utilities, helper functions)
- **Integration Tests**: Test integration with Microsoft Graph API using mocked HTTP requests
- **End-to-End Tests**: Manual testing with real OneNote data (not automated)

### Testing Tools

- **Jest**: Test framework and runner with ESM support
- **nock**: HTTP mocking library for Microsoft Graph API
- **@jest/globals**: Jest utilities for ES modules
- **JSDOM**: HTML parsing and manipulation testing

## Prerequisites

Before running tests, ensure you have:

- Node.js 18.x or later installed
- npm installed
- All project dependencies installed: `npm install`

## Quick Start

```bash
# Install dependencies (includes test dependencies)
npm install

# Run automated tests (unit + integration with mocks)
npm test

# Run tests in watch mode (great for development)
npm run test:watch

# Run only unit tests (fastest, always works)
npm run test:unit

# Run only integration tests (with mocked API)
npm run test:integration

# Generate coverage report
npm run test:coverage

# Run manual E2E tests (requires auth, makes REAL API calls)
npm run test:e2e

# Run E2E tests in read-only mode (no write operations)
npm run test:e2e:readonly
```

## Test Structure

```text
tests/
├── unit/                    # Unit tests for pure functions
│   └── htmlUtils.test.mjs   # HTML processing utilities
├── integration/             # Integration tests with mocked APIs
│   ├── authentication.test.mjs  # Auth flow tests
│   └── mcpTools.test.mjs        # MCP tool templates
├── manual/                  # Manual E2E tests (real API calls)
│   └── e2e-test.mjs         # End-to-end validation script
├── fixtures/                # Test data and mock responses
│   ├── htmlContent.mjs      # Sample HTML for testing
│   └── apiResponses.mjs     # Mock Graph API responses
└── helpers/                 # Test utilities
    ├── mockFactory.mjs      # Nock-based API mocking
    └── testUtils.mjs        # General test utilities
```

### Unit Tests (`tests/unit/`)

Unit tests focus on testing individual functions in isolation without external dependencies:

- **HTML Utilities**: `extractReadableText()`, `extractTextSummary()`, `textToHtml()`
- **Helper Functions**: Date formatting, text processing, etc.

**Example:**

```javascript
import { describe, it, expect } from '@jest/globals';

describe('extractReadableText', () => {
  it('should extract text from simple HTML', () => {
    const html = '<html><body><h1>Title</h1><p>Content</p></body></html>';
    const result = extractReadableText(html);
    expect(result).toContain('Title');
    expect(result).toContain('Content');
  });
});
```

### Integration Tests (`tests/integration/`)

Integration tests verify interactions with external services using mocked HTTP requests:

- **Authentication**: Token management, Graph client initialization
- **MCP Tools**: All 20 tools that interact with Microsoft Graph API
- **Error Handling**: Rate limiting, authentication errors, not found errors

**Example:**

```javascript
import { mockListNotebooks } from '../helpers/mockFactory.mjs';

describe('listNotebooks Tool', () => {
  it('should fetch notebooks from Graph API', async () => {
    mockListNotebooks();
    
    const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
    const data = await response.json();
    
    expect(data.value).toHaveLength(2);
  });
});
```

### Manual E2E Tests (`tests/manual/`)

Manual end-to-end tests validate the entire system with **real OneNote API calls**. Unlike unit and integration tests that use mocks, E2E tests require:

- **Valid authentication**: Active `.access-token.txt` file
- **Real OneNote account**: Tests will read/write actual data
- **Network connectivity**: Makes live calls to Microsoft Graph API

**When to use E2E tests:**

- Validating all MCP tools work with real OneNote API
- Testing after OAuth changes or Graph API updates
- Verifying functionality before releases
- Debugging issues that don't reproduce with mocks

**Running E2E tests:**

```bash
# Full E2E test suite (includes write operations)
npm run test:e2e

# Read-only mode (skips createPage and updatePageContent)
npm run test:e2e:readonly
```

**What E2E tests validate:**

✅ **Authentication**

- Token file exists and is valid
- GraphServiceClient initialization

✅ **Read Operations**

- `listNotebooks` - List all notebooks
- `listSections` - List sections in notebook
- `listPagesInSection` - List pages in section
- `searchPages` - Search by query string
- `searchPagesByDate` - Filter by date ranges
- `getPageByTitle` - Find page by exact/partial title
- `getPageContent` - Retrieve page HTML content
- `getUserInfo` - Get current user information

✅ **Write Operations** (skipped with `--read-only` flag)

- `createPage` - Create new test page
- `updatePageContent` - Modify existing page

✅ **Error Handling**

- 404 errors for non-existent resources
- Pagination for large result sets

**Example output:**

```text
===================================
OneNote MCP Server - E2E Tests
===================================

Starting E2E tests in 3 seconds...
This will make REAL API calls to your OneNote account.

✓ Authentication check passed
✓ User Info retrieved: John Doe
✓ List Notebooks: Found 3 notebooks
✓ List Sections: Found 5 sections
✓ List Pages: Found 12 pages
✓ Search Pages: Found 4 results
✓ Create Page: Created test page
✓ Update Page: Updated content successfully

===================================
Test Summary
===================================
Passed: 15 | Failed: 0 | Skipped: 0
Success Rate: 100%
```

**Important notes:**

⚠️ **Test pages**: E2E tests create pages titled "E2E Test Page - [timestamp]". Remember to delete them manually after testing.

⚠️ **Rate limits**: Making many API calls may trigger Microsoft Graph rate limiting. Use read-only mode for frequent testing.

⚠️ **Production data**: These tests interact with your real OneNote account. Use a test account if possible.

## Interactive Testing with MCP Inspector

The **MCP Inspector** is an interactive testing tool provided by the Model Context Protocol team. It allows you to test your MCP server in real-time with a graphical interface, making it ideal for:

- **Manual testing** of individual tools
- **Debugging** tool parameters and responses
- **Exploring** available tools and prompts
- **Validating** real API interactions during development

### Installing MCP Inspector

Install the inspector globally via npm:

```bash
npm install -g @modelcontextprotocol/inspector
```

### Running the Inspector

Start the MCP Inspector with your OneNote MCP server:

```bash
mcp-inspector node onenote-mcp.mjs
```

This will:

1. Start your MCP server
2. Launch a web-based inspector interface
3. Open your default browser to the inspector UI (typically `http://localhost:5173`)

### Using the Inspector Interface

The inspector provides several panels:

#### 1. Tools Panel

- View all available tools (20+ OneNote operations)
- See tool descriptions and parameter schemas
- Test tools interactively with custom parameters

#### 2. Prompts Panel

- View available prompts (like daily standup)
- Test prompt execution with arguments

#### 3. Resources Panel

- View available resources (if any)
- Test resource loading

#### 4. Server Logs Panel

- See real-time console output from your server
- Debug errors and trace execution
- View `console.error()` statements for debugging

### Testing Workflow Example

Here's a typical workflow for testing with the inspector:

1. **Start the inspector:**

   ```bash
   mcp-inspector node onenote-mcp.mjs
   ```

2. **Authenticate:**

   - Ensure you have a valid `.access-token.txt` file
   - Or run the OAuth flow first

3. **Test a simple tool:**

   - Select `listNotebooks` from the Tools panel
   - Click "Run" (no parameters needed)
   - View the response with your actual notebooks

4. **Test a complex tool:**

   - Select `searchPagesByDate` from the Tools panel
   - Set parameters:
     - `days`: 1
     - `notebookName`: "SQLNikon" (optional)
   - Click "Run"
   - View matching pages in the response

5. **Check server logs:**

   - Switch to the Logs panel
   - See debug output like:

     ```text
     Searching pages from last 1 day(s)...
     Found 9 notebook(s), fetching sections...
     Checked notebook "SQLNikon", 5 matches so far...
     ```

6. **Debug errors:**

   - If a tool fails, check the error message in the response
   - Review server logs for detailed stack traces
   - Verify your authentication token is valid

### Inspector Benefits vs Automated Tests

**Use Inspector for:**

- ✅ Quick manual testing during development
- ✅ Validating real API responses
- ✅ Debugging authentication issues
- ✅ Exploring tool capabilities
- ✅ Testing edge cases interactively

**Use Automated Tests for:**

- ✅ Regression testing
- ✅ CI/CD pipelines
- ✅ Unit testing individual functions
- ✅ Testing with mocked APIs (no rate limits)
- ✅ Validating code coverage

### Common Inspector Use Cases

**Testing a new tool:**

```bash
# 1. Start inspector
mcp-inspector node onenote-mcp.mjs

# 2. Select your new tool from the list
# 3. Fill in test parameters
# 4. Click "Run" and verify the response
# 5. Check logs for any warnings/errors
```

**Debugging authentication:**

```bash
# If you see auth errors in the inspector:
# 1. Check the Logs panel for specific error messages
# 2. Verify .access-token.txt exists and is valid
# 3. Try re-authenticating with the OAuth flow
# 4. Test getUserInfo tool to verify token works
```

**Performance testing:**

```bash
# Test performance improvements:
# 1. Run searchPagesByDate without notebookName parameter
# 2. Note the time in server logs
# 3. Run again with notebookName parameter
# 4. Compare execution time
```

### Tips for Using the Inspector

1. **Keep logs open**: The Logs panel shows valuable debugging information
2. **Test incrementally**: Start with simple tools, then move to complex ones
3. **Save test parameters**: Document successful parameter combinations
4. **Watch for rate limits**: Making many requests may trigger API throttling
5. **Use real data**: Inspector hits the actual OneNote API, not mocks

### Stopping the Inspector

To stop the inspector:

- Press `Ctrl+C` in the terminal where it's running
- Or close the terminal window

## Running Tests

### All Tests

Run the complete test suite:

```bash
npm test
```

### Watch Mode

Run tests automatically when files change (recommended during development):

```bash
npm run test:watch
```

Press `p` to filter by filename pattern, `t` to filter by test name pattern.

### Specific Test Suites

Run only unit tests:

```bash
npm run test:unit
```

Run only integration tests:

```bash
npm run test:integration
```

### Single Test File

Run a specific test file:

```bash
npm test -- tests/unit/htmlUtils.test.mjs
```

### Single Test Suite

Run tests matching a pattern:

```bash
npm test -- --testNamePattern="extractReadableText"
```

### Coverage Report

Generate a code coverage report:

```bash
npm run test:coverage
```

View the HTML coverage report:

```bash
open coverage/lcov-report/index.html  # macOS
xdg-open coverage/lcov-report/index.html  # Linux
start coverage/lcov-report/index.html  # Windows
```

## Writing Tests

### Adding a New Unit Test

1. Create a test file in `tests/unit/` with `.test.mjs` extension
2. Import test utilities: `import { describe, it, expect } from '@jest/globals';`
3. Write test suites using `describe()` and individual tests using `it()`
4. Use Jest matchers: `expect().toBe()`, `expect().toContain()`, etc.

**Example:**

```javascript
// tests/unit/myFunction.test.mjs
import { describe, it, expect } from '@jest/globals';
import { myFunction } from '../../onenote-mcp.mjs';

describe('myFunction', () => {
  it('should do something specific', () => {
    const result = myFunction('input');
    expect(result).toBe('expected output');
  });

  it('should handle edge cases', () => {
    expect(myFunction('')).toBe('');
    expect(myFunction(null)).toBe('');
  });
});
```

### Adding a New Integration Test

1. Create a test file in `tests/integration/` with `.test.mjs` extension
2. Import mocking utilities from `tests/helpers/mockFactory.mjs`
3. Set up and tear down mocks in `beforeEach()` and `afterEach()`
4. Mock Graph API endpoints and verify responses

**Example:**

```javascript
// tests/integration/myTool.test.mjs
import { describe, it, expect, beforeEach, afterEach } from '@jest/globals';
import { setupGraphAPIMocks, teardownGraphAPIMocks } from '../helpers/mockFactory.mjs';

describe('My Tool Integration Tests', () => {
  beforeEach(() => {
    setupGraphAPIMocks();
  });

  afterEach(() => {
    teardownGraphAPIMocks();
  });

  it('should interact with Graph API', async () => {
    // Your test code here
  });
});
```

### Creating Mock Data

Add new mock data to `tests/fixtures/`:

**API Responses** (`tests/fixtures/apiResponses.mjs`):

```javascript
export const mockMyData = {
  value: [
    { id: '1', name: 'Test' }
  ]
};
```

**HTML Content** (`tests/fixtures/htmlContent.mjs`):

```javascript
export const myTestHTML = `
  <html>
    <body>
      <h1>Test Content</h1>
    </body>
  </html>
`;
```

### Using Mock Factory

The mock factory (`tests/helpers/mockFactory.mjs`) provides helpers for mocking Graph API:

```javascript
import {
  mockListNotebooks,
  mockListSections,
  mockGetPageContent,
  mockRateLimitError
} from '../helpers/mockFactory.mjs';

// Mock successful API call
mockListNotebooks();

// Mock with custom response
mockListNotebooks({ value: [{ id: '1', displayName: 'Custom' }] });

// Mock error response
mockRateLimitError('/v1.0/me/onenote/notebooks');
```

## Continuous Integration

### GitHub Actions

The project includes a GitHub Actions workflow (`.github/workflows/test.yml`) that automatically runs tests on:

- **Triggers**: Push to `main`/`develop` branches, pull requests
- **Operating Systems**: Ubuntu, macOS, Windows
- **Node Versions**: 18.x, 20.x, 22.x

The CI pipeline includes:

1. **Test Job**: Runs all tests with coverage reporting
2. **Lint Job**: Checks for syntax errors
3. **Security Job**: Runs `npm audit` for vulnerabilities

### Viewing CI Results

1. Go to the repository's "Actions" tab on GitHub
2. Click on a workflow run to see results
3. Download test artifacts for detailed reports

## Coverage Requirements

The project maintains the following minimum coverage thresholds (configured in `jest.config.js`):

- **Branches**: 70%
- **Functions**: 70%
- **Lines**: 80%
- **Statements**: 80%

### Checking Coverage

```bash
npm run test:coverage
```

If coverage falls below thresholds, tests will fail. To view uncovered code:

1. Run `npm run test:coverage`
2. Open `coverage/lcov-report/index.html`
3. Click on files to see line-by-line coverage

## Troubleshooting

### Common Issues

#### Tests Fail with "Cannot use import statement outside a module"

**Solution**: Ensure you're using the correct file extension (`.mjs`) and running tests with the ESM flag:

```bash
node --experimental-vm-modules node_modules/jest/bin/jest.js
```

#### Nock Mocks Not Working

**Solution**: Verify that `setupGraphAPIMocks()` is called in `beforeEach()` and `teardownGraphAPIMocks()` in `afterEach()`:

```javascript
beforeEach(() => {
  setupGraphAPIMocks();
});

afterEach(() => {
  teardownGraphAPIMocks();
});
```

#### Tests Timeout

**Solution**: Increase timeout in `jest.config.js` or for individual tests:

```javascript
it('long running test', async () => {
  // test code
}, 30000); // 30 second timeout
```

#### Coverage Report Not Generated

**Solution**: Ensure you're running with the coverage flag:

```bash
npm run test:coverage
```

#### Mock HTTP Requests Not Matching

**Solution**: Check that your nock mock matches the exact URL and query parameters:

```javascript
// Exact match
nock('https://graph.microsoft.com')
  .get('/v1.0/me/onenote/notebooks')
  .reply(200, data);

// Match any query params
nock('https://graph.microsoft.com')
  .get('/v1.0/me/onenote/pages')
  .query(true)
  .reply(200, data);
```

### Debug Mode

Run tests with additional debugging:

```bash
# Show console.log output
npm test -- --verbose

# Run a single test file
npm test -- tests/unit/htmlUtils.test.mjs

# Show which tests are running
npm test -- --verbose --no-coverage
```

## Best Practices

1. **Write Tests First**: Consider TDD (Test-Driven Development) for new features
2. **Test Edge Cases**: Empty inputs, null values, malformed data
3. **Keep Tests Isolated**: Each test should be independent
4. **Use Descriptive Names**: Test names should clearly describe what they test
5. **Mock External Dependencies**: Never make real API calls in tests
6. **Maintain High Coverage**: Aim for >80% code coverage
7. **Update Tests with Code Changes**: Keep tests in sync with implementation

## Resources

- [Jest Documentation](https://jestjs.io/docs/getting-started)
- [Nock Documentation](https://github.com/nock/nock)
- [Microsoft Graph API Reference](https://learn.microsoft.com/en-us/graph/api/overview)
- [JSDOM Documentation](https://github.com/jsdom/jsdom)

## Contributing

When contributing tests:

1. Follow the existing test structure and naming conventions
2. Ensure all tests pass before submitting a PR
3. Add tests for any new functionality
4. Maintain or improve code coverage
5. Update this documentation if adding new test patterns

---

For questions or issues with testing, please open an issue on the GitHub repository.
