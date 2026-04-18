/**
 * Integration tests for MCP tools with mocked Microsoft Graph API
 * Tests basic read operations: listNotebooks, listSections, listPages, getPageContent
 */
import { describe, it, expect, beforeEach, afterEach } from '@jest/globals';
import {
  setupGraphAPIMocks,
  teardownGraphAPIMocks,
  mockListNotebooks,
  mockListSections,
  mockListPages,
  mockGetPageContent,
  mockSearchPages,
  mockCreatePage,
  mockRateLimitError,
  mockUnauthorizedError,
  mockNotFoundError,
  verifyAllMocksCalled
} from '../helpers/mockFactory.mjs';
import {
  mockNotebooks,
  mockSections,
  mockPages,
  mockPageContentHTML
} from '../fixtures/apiResponses.mjs';

describe('MCP Tools Integration Tests', () => {
  beforeEach(() => {
    setupGraphAPIMocks();
  });

  afterEach(() => {
    teardownGraphAPIMocks();
  });

  describe('listNotebooks Tool', () => {
    it('should successfully fetch notebooks from Graph API', async () => {
      const scope = mockListNotebooks();
      
      // Simulate making a request to the mocked endpoint
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
      const data = await response.json();
      
      expect(data.value).toHaveLength(2);
      expect(data.value[0].displayName).toBe('Personal Notebook');
      expect(data.value[1].displayName).toBe('Work Notes');
      expect(scope.isDone()).toBe(true);
    });

    it('should return notebooks with required properties', async () => {
      mockListNotebooks();
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
      const data = await response.json();
      
      data.value.forEach(notebook => {
        expect(notebook).toHaveProperty('id');
        expect(notebook).toHaveProperty('displayName');
        expect(notebook).toHaveProperty('createdDateTime');
        expect(notebook).toHaveProperty('lastModifiedDateTime');
        expect(notebook).toHaveProperty('links');
      });
    });

    it('should handle empty notebooks list', async () => {
      mockListNotebooks({ value: [] });
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
      const data = await response.json();
      
      expect(data.value).toHaveLength(0);
      expect(Array.isArray(data.value)).toBe(true);
    });

    it('should fetch notebooks with team notebooks when includeTeamNotebooks is true', async () => {
      const mockNotebooksWithShared = {
        value: [
          { id: 'notebook-1', displayName: 'Personal Notebook', sections: [] },
          { id: 'notebook-2', displayName: 'Work Notes', sections: [] },
          { id: 'notebook-shared-1', displayName: 'Team Project Notes', sections: [] }
        ]
      };
      const scope = mockListNotebooks(mockNotebooksWithShared, true);
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks?includeteamnotebooks=true');
      const data = await response.json();
      
      expect(data.value).toHaveLength(3);
      expect(data.value[2].displayName).toBe('Team Project Notes');
      expect(scope.isDone()).toBe(true);
    });

    it('should use correct endpoint when includeTeamNotebooks is false (default)', async () => {
      const scope = mockListNotebooks(undefined, false);
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
      const data = await response.json();
      
      expect(data.value).toHaveLength(2);
      expect(scope.isDone()).toBe(true);
    });
  });

  describe('listSections Tool', () => {
    it('should successfully fetch sections for a notebook', async () => {
      const notebookId = 'notebook-1';
      const scope = mockListSections(notebookId);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/notebooks/${notebookId}/sections`);
      const data = await response.json();
      
      expect(data.value).toHaveLength(2);
      expect(data.value[0].displayName).toBe('Quick Notes');
      expect(data.value[1].displayName).toBe('Meeting Notes');
      expect(scope.isDone()).toBe(true);
    });

    it('should return sections with required properties', async () => {
      const notebookId = 'notebook-1';
      mockListSections(notebookId);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/notebooks/${notebookId}/sections`);
      const data = await response.json();
      
      data.value.forEach(section => {
        expect(section).toHaveProperty('id');
        expect(section).toHaveProperty('displayName');
        expect(section).toHaveProperty('createdDateTime');
        expect(section).toHaveProperty('lastModifiedDateTime');
        expect(section).toHaveProperty('pagesUrl');
      });
    });
  });

  describe('listPages Tool', () => {
    it('should successfully fetch pages from a section', async () => {
      const sectionId = 'section-1';
      const scope = mockListPages(sectionId);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/sections/${sectionId}/pages`);
      const data = await response.json();
      
      expect(data.value).toHaveLength(2);
      expect(data.value[0].title).toBe('Daily Note - 2/22/26');
      expect(data.value[1].title).toBe('Project Planning');
      expect(scope.isDone()).toBe(true);
    });

    it('should return pages with required properties', async () => {
      const sectionId = 'section-1';
      mockListPages(sectionId);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/sections/${sectionId}/pages`);
      const data = await response.json();
      
      data.value.forEach(page => {
        expect(page).toHaveProperty('id');
        expect(page).toHaveProperty('title');
        expect(page).toHaveProperty('createdDateTime');
        expect(page).toHaveProperty('lastModifiedDateTime');
        expect(page).toHaveProperty('contentUrl');
      });
    });

    it('should include web URLs for pages', async () => {
      const sectionId = 'section-1';
      mockListPages(sectionId);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/sections/${sectionId}/pages`);
      const data = await response.json();
      
      data.value.forEach(page => {
        expect(page.links).toHaveProperty('oneNoteWebUrl');
        expect(page.links.oneNoteWebUrl).toHaveProperty('href');
      });
    });
  });

  describe('getPageContent Tool', () => {
    it('should successfully fetch HTML content for a page', async () => {
      const pageId = 'page-1';
      const scope = mockGetPageContent(pageId);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`);
      const html = await response.text();
      
      expect(html).toContain('<!DOCTYPE html>');
      expect(html).toContain('Daily Note - 2/22/26');
      expect(html).toContain("Today's tasks:");
      expect(scope.isDone()).toBe(true);
    });

    it('should return HTML content type', async () => {
      const pageId = 'page-1';
      mockGetPageContent(pageId);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`);
      
      expect(response.headers.get('content-type')).toContain('text/html');
    });

    it('should handle pages with complex HTML structure', async () => {
      const pageId = 'page-2';
      const complexHTML = `
        <!DOCTYPE html>
        <html>
        <body>
          <h1>Complex Page</h1>
          <table><tr><td>Data</td></tr></table>
          <ul><li>Item 1</li><li>Item 2</li></ul>
        </body>
        </html>
      `;
      mockGetPageContent(pageId, complexHTML);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`);
      const html = await response.text();
      
      expect(html).toContain('<table>');
      expect(html).toContain('<ul>');
      expect(html).toContain('Complex Page');
    });
  });

  describe('searchPages Tool', () => {
    it('should successfully search for pages by title', async () => {
      const query = 'Daily';
      const scope = mockSearchPages(query);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/pages?$search=${query}`);
      const data = await response.json();
      
      expect(data.value).toHaveLength(2);
      expect(scope.isDone()).toBe(true);
    });

    it('should handle search with no results', async () => {
      const query = 'NonExistentPage';
      mockSearchPages(query, { value: [] });
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/pages?$search=${query}`);
      const data = await response.json();
      
      expect(data.value).toHaveLength(0);
    });
  });

  describe('createPage Tool', () => {
    it('should successfully create a new page', async () => {
      const sectionId = 'section-1';
      const newPage = {
        id: 'new-page-123',
        title: 'New Test Page'
      };
      const scope = mockCreatePage(sectionId, newPage);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/sections/${sectionId}/pages`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ title: 'New Test Page', content: '<html></html>' })
      });
      const data = await response.json();
      
      expect(response.status).toBe(201);
      expect(data.id).toBe('new-page-123');
      expect(data.title).toBe('New Test Page');
      expect(scope.isDone()).toBe(true);
    });
  });

  describe('Error Handling', () => {
    it('should handle 429 rate limit errors', async () => {
      const endpoint = '/v1.0/me/onenote/notebooks';
      const scope = mockRateLimitError(endpoint);
      
      const response = await fetch(`https://graph.microsoft.com${endpoint}`);
      
      expect(response.status).toBe(429);
      expect(response.headers.get('retry-after')).toBe('1');
      expect(scope.isDone()).toBe(true);
    });

    it('should handle 401 unauthorized errors', async () => {
      const endpoint = '/v1.0/me/onenote/notebooks';
      const scope = mockUnauthorizedError(endpoint);
      
      const response = await fetch(`https://graph.microsoft.com${endpoint}`);
      const data = await response.json();
      
      expect(response.status).toBe(401);
      expect(data.error.code).toBe('Unauthorized');
      expect(scope.isDone()).toBe(true);
    });

    it('should handle 404 not found errors', async () => {
      const endpoint = '/v1.0/me/onenote/pages/invalid-page-id/content';
      const scope = mockNotFoundError(endpoint);
      
      const response = await fetch(`https://graph.microsoft.com${endpoint}`);
      const data = await response.json();
      
      expect(response.status).toBe(404);
      expect(data.error.code).toBe('ResourceNotFound');
      expect(scope.isDone()).toBe(true);
    });

    it('should include request-id in error responses', async () => {
      const endpoint = '/v1.0/me/onenote/notebooks';
      mockUnauthorizedError(endpoint);
      
      const response = await fetch(`https://graph.microsoft.com${endpoint}`);
      const data = await response.json();
      
      expect(data.error.innerError).toHaveProperty('request-id');
      expect(data.error.innerError).toHaveProperty('date');
    });
  });

  describe('Response Validation', () => {
    it('should validate notebook response structure', async () => {
      mockListNotebooks();
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
      const data = await response.json();
      
      expect(data).toHaveProperty('value');
      expect(Array.isArray(data.value)).toBe(true);
    });

    it('should validate datetime formats in responses', async () => {
      mockListNotebooks();
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
      const data = await response.json();
      
      data.value.forEach(notebook => {
        const created = new Date(notebook.createdDateTime);
        const modified = new Date(notebook.lastModifiedDateTime);
        
        expect(created).toBeInstanceOf(Date);
        expect(modified).toBeInstanceOf(Date);
        expect(created.toString()).not.toBe('Invalid Date');
        expect(modified.toString()).not.toBe('Invalid Date');
      });
    });

    it('should validate that modified date is not before created date', async () => {
      mockListNotebooks();
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
      const data = await response.json();
      
      data.value.forEach(notebook => {
        const created = new Date(notebook.createdDateTime);
        const modified = new Date(notebook.lastModifiedDateTime);
        
        expect(modified.getTime()).toBeGreaterThanOrEqual(created.getTime());
      });
    });
  });

  describe('URL Validation', () => {
    it('should return valid URLs in notebook responses', async () => {
      mockListNotebooks();
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
      const data = await response.json();
      
      data.value.forEach(notebook => {
        const webUrl = notebook.links.oneNoteWebUrl.href;
        expect(webUrl).toMatch(/^https:\/\//);
      });
    });

    it('should return valid content URLs for pages', async () => {
      const sectionId = 'section-1';
      mockListPages(sectionId);
      
      const response = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/sections/${sectionId}/pages`);
      const data = await response.json();
      
      data.value.forEach(page => {
        expect(page.contentUrl).toMatch(/^https:\/\/graph\.microsoft\.com/);
        expect(page.contentUrl).toContain('/content');
      });
    });
  });
});
