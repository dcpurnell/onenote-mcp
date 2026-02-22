/**
 * Test utilities and helper functions
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Creates a temporary token file for testing
 * @param {Object} tokenData - Token data to write
 * @param {string} filePath - Path to token file
 */
export function createMockTokenFile(tokenData, filePath) {
  fs.writeFileSync(filePath, JSON.stringify(tokenData, null, 2), 'utf8');
}

/**
 * Removes a temporary token file after testing
 * @param {string} filePath - Path to token file
 */
export function removeMockTokenFile(filePath) {
  if (fs.existsSync(filePath)) {
    fs.unlinkSync(filePath);
  }
}

/**
 * Creates a mock Graph client for testing
 * @param {Object} responses - Mock responses for different endpoints
 * @returns {Object} Mock Graph client
 */
export function createMockGraphClient(responses = {}) {
  return {
    api: (endpoint) => ({
      get: async () => responses[endpoint] || { value: [] },
      post: async (data) => ({ id: 'mock-created-id', ...data }),
      patch: async (data) => ({ success: true, ...data }),
      delete: async () => ({ success: true }),
      select: function(fields) { return this; },
      expand: function(fields) { return this; },
      filter: function(query) { return this; },
      orderby: function(field) { return this; },
      top: function(count) { return this; },
    })
  };
}

/**
 * Delays execution for testing async operations
 * @param {number} ms - Milliseconds to delay
 * @returns {Promise}
 */
export function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Captures console output during test execution
 * @param {Function} fn - Function to execute
 * @returns {Object} Object with stdout and stderr output
 */
export async function captureConsole(fn) {
  const originalLog = console.log;
  const originalError = console.error;
  const stdout = [];
  const stderr = [];

  console.log = (...args) => stdout.push(args.join(' '));
  console.error = (...args) => stderr.push(args.join(' '));

  try {
    await fn();
  } finally {
    console.log = originalLog;
    console.error = originalError;
  }

  return { stdout, stderr };
}

/**
 * Generates a random ID for testing
 * @param {string} prefix - Prefix for the ID
 * @returns {string} Random ID
 */
export function generateMockId(prefix = 'mock') {
  return `${prefix}-${Math.random().toString(36).substring(2, 15)}`;
}

/**
 * Creates a mock page object
 * @param {Object} overrides - Properties to override
 * @returns {Object} Mock page object
 */
export function createMockPage(overrides = {}) {
  const id = generateMockId('page');
  return {
    id,
    title: 'Test Page',
    createdDateTime: new Date().toISOString(),
    lastModifiedDateTime: new Date().toISOString(),
    contentUrl: `https://graph.microsoft.com/v1.0/me/onenote/pages/${id}/content`,
    links: {
      oneNoteWebUrl: {
        href: `https://onenote.com/${id}`
      }
    },
    ...overrides
  };
}

/**
 * Creates a mock notebook object
 * @param {Object} overrides - Properties to override
 * @returns {Object} Mock notebook object
 */
export function createMockNotebook(overrides = {}) {
  const id = generateMockId('notebook');
  return {
    id,
    displayName: 'Test Notebook',
    createdDateTime: new Date().toISOString(),
    lastModifiedDateTime: new Date().toISOString(),
    links: {
      oneNoteWebUrl: {
        href: `https://onenote.com/${id}`
      }
    },
    ...overrides
  };
}

/**
 * Creates a mock section object
 * @param {Object} overrides - Properties to override
 * @returns {Object} Mock section object
 */
export function createMockSection(overrides = {}) {
  const id = generateMockId('section');
  return {
    id,
    displayName: 'Test Section',
    createdDateTime: new Date().toISOString(),
    lastModifiedDateTime: new Date().toISOString(),
    pagesUrl: `https://graph.microsoft.com/v1.0/me/onenote/sections/${id}/pages`,
    ...overrides
  };
}

/**
 * Asserts that a value matches expected type
 * @param {*} value - Value to check
 * @param {string} expectedType - Expected type
 * @throws {Error} If types don't match
 */
export function assertType(value, expectedType) {
  const actualType = typeof value;
  if (actualType !== expectedType) {
    throw new Error(`Expected type ${expectedType}, got ${actualType}`);
  }
}

/**
 * Asserts that an array contains specific items
 * @param {Array} array - Array to check
 * @param {Array} items - Items to find
 * @throws {Error} If items not found
 */
export function assertArrayContains(array, items) {
  for (const item of items) {
    if (!array.includes(item)) {
      throw new Error(`Array does not contain: ${item}`);
    }
  }
}
