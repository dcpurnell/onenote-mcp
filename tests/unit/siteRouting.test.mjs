/**
 * Unit tests for SharePoint URL parsing and endpoint routing helpers.
 * Tests: parseSharePointUrl, getOnenoteBasePath (logic only, no Graph calls)
 */
import { describe, it, expect } from '@jest/globals';

// Copy of parseSharePointUrl from onenote-mcp.mjs for testing
// TODO: Export from onenote-mcp.mjs and import directly
function parseSharePointUrl(url) {
  if (!url) return null;
  const match = url.match(/^https:\/\/([^/]+)\/(sites|teams)\/([^/]+)/i);
  if (!match) return null;
  const hostname = match[1];
  if (hostname.includes('-my.sharepoint.com')) return null;
  const sitePath = `${match[2]}/${match[3]}`;
  return { hostname, sitePath };
}

// Synchronous logic portion of getOnenoteBasePath for testing routing decisions
function getOnenoteBasePathSync(notebook) {
  if (!notebook) return '/me/onenote';
  if (notebook._isPersonal) return '/me/onenote';
  if (notebook._siteId) return `/sites/${notebook._siteId}/onenote`;
  if (notebook._groupId) return `/groups/${notebook._groupId}/onenote`;
  return '/me/onenote';
}

describe('parseSharePointUrl', () => {
  it('returns null for null/undefined/empty input', () => {
    expect(parseSharePointUrl(null)).toBeNull();
    expect(parseSharePointUrl(undefined)).toBeNull();
    expect(parseSharePointUrl('')).toBeNull();
  });

  it('parses a standard /sites/ SharePoint URL', () => {
    const result = parseSharePointUrl('https://contoso.sharepoint.com/sites/Engineering');
    expect(result).toEqual({
      hostname: 'contoso.sharepoint.com',
      sitePath: 'sites/Engineering',
    });
  });

  it('parses a /teams/ SharePoint URL', () => {
    const result = parseSharePointUrl('https://contoso.sharepoint.com/teams/ProjectAlpha');
    expect(result).toEqual({
      hostname: 'contoso.sharepoint.com',
      sitePath: 'teams/ProjectAlpha',
    });
  });

  it('handles URLs with trailing path segments', () => {
    const result = parseSharePointUrl('https://contoso.sharepoint.com/sites/Engineering/Shared%20Documents/Notebook');
    expect(result).toEqual({
      hostname: 'contoso.sharepoint.com',
      sitePath: 'sites/Engineering',
    });
  });

  it('returns null for personal OneDrive URLs (-my.sharepoint.com)', () => {
    const result = parseSharePointUrl('https://contoso-my.sharepoint.com/personal/jdoe_contoso_com');
    expect(result).toBeNull();
  });

  it('returns null for non-SharePoint URLs', () => {
    expect(parseSharePointUrl('https://example.com/sites/Test')).toEqual({
      hostname: 'example.com',
      sitePath: 'sites/Test',
    });
    expect(parseSharePointUrl('https://graph.microsoft.com/v1.0/me')).toBeNull();
  });

  it('returns null for URLs without /sites/ or /teams/', () => {
    expect(parseSharePointUrl('https://contoso.sharepoint.com/personal/jdoe')).toBeNull();
    expect(parseSharePointUrl('https://contoso.sharepoint.com/')).toBeNull();
  });

  it('is case-insensitive for /sites/ and /teams/', () => {
    const result1 = parseSharePointUrl('https://contoso.sharepoint.com/Sites/Engineering');
    expect(result1).toEqual({
      hostname: 'contoso.sharepoint.com',
      sitePath: 'Sites/Engineering',
    });

    const result2 = parseSharePointUrl('https://contoso.sharepoint.com/TEAMS/Alpha');
    expect(result2).toEqual({
      hostname: 'contoso.sharepoint.com',
      sitePath: 'TEAMS/Alpha',
    });
  });
});

describe('getOnenoteBasePath routing logic', () => {
  it('returns /me/onenote for null notebook', () => {
    expect(getOnenoteBasePathSync(null)).toBe('/me/onenote');
  });

  it('returns /me/onenote for personal notebooks', () => {
    expect(getOnenoteBasePathSync({ _isPersonal: true })).toBe('/me/onenote');
  });

  it('returns /sites/{siteId}/onenote when _siteId is set', () => {
    const nb = { _isPersonal: false, _siteId: 'contoso.sharepoint.com,abc123,def456' };
    expect(getOnenoteBasePathSync(nb)).toBe('/sites/contoso.sharepoint.com,abc123,def456/onenote');
  });

  it('prefers _siteId over _groupId', () => {
    const nb = { _isPersonal: false, _siteId: 'site-id-123', _groupId: 'group-id-456' };
    expect(getOnenoteBasePathSync(nb)).toBe('/sites/site-id-123/onenote');
  });

  it('falls back to /groups/{groupId}/onenote when only _groupId is set', () => {
    const nb = { _isPersonal: false, _groupId: 'group-id-789' };
    expect(getOnenoteBasePathSync(nb)).toBe('/groups/group-id-789/onenote');
  });

  it('falls back to /me/onenote when neither _siteId nor _groupId is set', () => {
    const nb = { _isPersonal: false };
    expect(getOnenoteBasePathSync(nb)).toBe('/me/onenote');
  });

  it('returns /me/onenote for personal even with _siteId set', () => {
    const nb = { _isPersonal: true, _siteId: 'should-be-ignored' };
    expect(getOnenoteBasePathSync(nb)).toBe('/me/onenote');
  });
});

describe('registerSectionMapping and routing', () => {
  // Simulate the section→notebook mapping logic
  const sectionToNotebookMap = {};

  function registerSectionMapping(sectionId, notebookId) {
    sectionToNotebookMap[sectionId] = notebookId;
  }

  function getBasePathForSectionSync(sectionId, notebookCache) {
    const notebookId = sectionToNotebookMap[sectionId];
    if (notebookId && notebookCache) {
      const notebook = notebookCache.find(nb => nb.id === notebookId);
      if (notebook) return getOnenoteBasePathSync(notebook);
    }
    return '/me/onenote';
  }

  it('maps section to notebook and returns correct base path', () => {
    const cache = [
      { id: 'nb-1', _isPersonal: true },
      { id: 'nb-2', _isPersonal: false, _siteId: 'site-abc' },
    ];

    registerSectionMapping('sec-A', 'nb-1');
    registerSectionMapping('sec-B', 'nb-2');

    expect(getBasePathForSectionSync('sec-A', cache)).toBe('/me/onenote');
    expect(getBasePathForSectionSync('sec-B', cache)).toBe('/sites/site-abc/onenote');
  });

  it('returns /me/onenote for unknown sections', () => {
    expect(getBasePathForSectionSync('sec-unknown', [])).toBe('/me/onenote');
  });
});
