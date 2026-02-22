/**
 * Mock factory for Microsoft Graph API using nock
 */
import nock from 'nock';
import {
  mockNotebooks,
  mockSections,
  mockPages,
  mockPageContentHTML,
  mockPaginatedResponse,
  mockUserInfo,
  mockErrorResponses
} from '../fixtures/apiResponses.mjs';

const GRAPH_API_BASE = 'https://graph.microsoft.com';
const GRAPH_API_VERSION = '/v1.0';

/**
 * Sets up nock to mock all Microsoft Graph API calls
 */
export function setupGraphAPIMocks() {
  // Disable real HTTP requests
  nock.disableNetConnect();
}

/**
 * Cleans up all nock mocks
 */
export function teardownGraphAPIMocks() {
  nock.cleanAll();
  nock.enableNetConnect();
}

/**
 * Mocks the notebooks list endpoint
 * @param {Object} response - Custom response (defaults to mockNotebooks)
 * @returns {nock.Scope} Nock scope
 */
export function mockListNotebooks(response = mockNotebooks) {
  return nock(GRAPH_API_BASE)
    .get(`${GRAPH_API_VERSION}/me/onenote/notebooks`)
    .reply(200, response);
}

/**
 * Mocks the sections list endpoint
 * @param {string} notebookId - Notebook ID
 * @param {Object} response - Custom response (defaults to mockSections)
 * @returns {nock.Scope} Nock scope
 */
export function mockListSections(notebookId, response = mockSections) {
  return nock(GRAPH_API_BASE)
    .get(`${GRAPH_API_VERSION}/me/onenote/notebooks/${notebookId}/sections`)
    .reply(200, response);
}

/**
 * Mocks the pages list endpoint
 * @param {string} sectionId - Section ID
 * @param {Object} response - Custom response (defaults to mockPages)
 * @returns {nock.Scope} Nock scope
 */
export function mockListPages(sectionId, response = mockPages) {
  return nock(GRAPH_API_BASE)
    .get(`${GRAPH_API_VERSION}/me/onenote/sections/${sectionId}/pages`)
    .reply(200, response);
}

/**
 * Mocks the page content endpoint
 * @param {string} pageId - Page ID
 * @param {string} response - Custom HTML response (defaults to mockPageContentHTML)
 * @returns {nock.Scope} Nock scope
 */
export function mockGetPageContent(pageId, response = mockPageContentHTML) {
  return nock(GRAPH_API_BASE)
    .get(`${GRAPH_API_VERSION}/me/onenote/pages/${pageId}/content`)
    .reply(200, response, { 'Content-Type': 'text/html' });
}

/**
 * Mocks the page search endpoint
 * @param {string} query - Search query
 * @param {Object} response - Custom response (defaults to mockPages)
 * @returns {nock.Scope} Nock scope
 */
export function mockSearchPages(query, response = mockPages) {
  return nock(GRAPH_API_BASE)
    .get(`${GRAPH_API_VERSION}/me/onenote/pages`)
    .query(true) // Accept any query parameters
    .reply(200, response);
}

/**
 * Mocks the create page endpoint
 * @param {string} sectionId - Section ID
 * @param {Object} response - Custom response
 * @returns {nock.Scope} Nock scope
 */
export function mockCreatePage(sectionId, response) {
  return nock(GRAPH_API_BASE)
    .post(`${GRAPH_API_VERSION}/me/onenote/sections/${sectionId}/pages`)
    .reply(201, response || { id: 'new-page-id', title: 'New Page' });
}

/**
 * Mocks the update page content endpoint (PATCH)
 * @param {string} pageId - Page ID
 * @param {Object} response - Custom response
 * @returns {nock.Scope} Nock scope
 */
export function mockUpdatePageContent(pageId, response) {
  return nock(GRAPH_API_BASE)
    .patch(`${GRAPH_API_VERSION}/me/onenote/pages/${pageId}/content`)
    .reply(204, response || {});
}

/**
 * Mocks user info endpoint
 * @param {Object} response - Custom response (defaults to mockUserInfo)
 * @returns {nock.Scope} Nock scope
 */
export function mockGetUserInfo(response = mockUserInfo) {
  return nock(GRAPH_API_BASE)
    .get(`${GRAPH_API_VERSION}/me`)
    .reply(200, response);
}

/**
 * Mocks a 429 rate limit response
 * @param {string} endpoint - Endpoint path
 * @returns {nock.Scope} Nock scope
 */
export function mockRateLimitError(endpoint) {
  return nock(GRAPH_API_BASE)
    .get(endpoint)
    .reply(429, mockErrorResponses.rateLimit, {
      'Retry-After': '1'
    });
}

/**
 * Mocks a 401 unauthorized response
 * @param {string} endpoint - Endpoint path
 * @returns {nock.Scope} Nock scope
 */
export function mockUnauthorizedError(endpoint) {
  return nock(GRAPH_API_BASE)
    .get(endpoint)
    .reply(401, mockErrorResponses.unauthorized);
}

/**
 * Mocks a 404 not found response
 * @param {string} endpoint - Endpoint path
 * @returns {nock.Scope} Nock scope
 */
export function mockNotFoundError(endpoint) {
  return nock(GRAPH_API_BASE)
    .get(endpoint)
    .reply(404, mockErrorResponses.notFound);
}

/**
 * Mocks a paginated response with nextLink
 * @param {string} endpoint - Initial endpoint
 * @param {Array} pages - Array of page responses
 * @returns {Array<nock.Scope>} Array of nock scopes
 */
export function mockPaginatedRequest(endpoint, pages) {
  const scopes = [];
  
  pages.forEach((page, index) => {
    const isLast = index === pages.length - 1;
    const response = {
      value: page,
      ...(isLast ? {} : { '@odata.nextLink': `${endpoint}?$skip=${(index + 1) * 10}` })
    };
    
    const queryParams = index === 0 ? {} : { $skip: (index * 10).toString() };
    
    scopes.push(
      nock(GRAPH_API_BASE)
        .get(endpoint)
        .query(queryParams)
        .reply(200, response)
    );
  });
  
  return scopes;
}

/**
 * Verifies that all mocked requests were called
 * @throws {Error} If there are pending mocks
 */
export function verifyAllMocksCalled() {
  if (!nock.isDone()) {
    const pending = nock.pendingMocks();
    throw new Error(`Pending mocks were not called: ${pending.join(', ')}`);
  }
}
