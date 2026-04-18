#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { Client } from '@microsoft/microsoft-graph-client';
import { DeviceCodeCredential } from '@azure/identity';
import { JSDOM } from 'jsdom';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import fetch from 'node-fetch';
import { z } from "zod";

// --- Configuration ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const tokenFilePath = path.join(__dirname, '.access-token.txt');
const notebookCacheFilePath = path.join(__dirname, '.notebook-cache.json');
const clientId = process.env.AZURE_CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Default: Microsoft Graph Explorer App ID
// Updated scopes to include .All permissions for accessing shared/team notebooks
const scopes = ['Notes.Read', 'Notes.ReadWrite', 'Notes.Read.All', 'Notes.ReadWrite.All', 'Notes.Create', 'User.Read'];
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes

// --- Global State ---
let accessToken = null;
let graphClient = null;
let notebookCache = null; // Cache of notebooks with group info
let cacheTimestamp = null; // When cache was last updated
let teamNotebooksLoading = false; // Flag to prevent duplicate team loads

// --- MCP Server Initialization ---
const server = new McpServer({
  name: 'onenote',
  version: '1.0.0', 
  description: 'OneNote MCP Server - Read, Write, and Edit OneNote content.'
});

// ============================================================================
// AUTHENTICATION & MICROSOFT GRAPH CLIENT MANAGEMENT
// ============================================================================

/**
 * Loads an existing access token from the local file system.
 */
function loadExistingToken() {
  try {
    if (fs.existsSync(tokenFilePath)) {
      const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
      try {
        const parsedToken = JSON.parse(tokenData); // New format: JSON object
        accessToken = parsedToken.token;
        console.error('Loaded existing token from file (JSON format).');
      } catch (parseError) {
        accessToken = tokenData; // Old format: plain token string
        console.error('Loaded existing token from file (plain text format).');
      }
    }
  } catch (error) {
    console.error(`Error loading token: ${error.message}`);
  }
}

/**
 * Initializes the Microsoft Graph client if an access token is available.
 * @returns {Client | null} The initialized Graph client or null.
 */
function initializeGraphClient() {
  if (accessToken && !graphClient) {
    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });
    console.error('Microsoft Graph client initialized.');
  }
  return graphClient;
}

/**
 * Ensures the Graph client is initialized and authenticated.
 * Loads token if not present, then initializes client.
 * @throws {Error} If no access token is available after attempting to load.
 * @returns {Promise<Client>} The initialized and authenticated Graph client.
 */
async function ensureGraphClient() {
  if (!accessToken) {
    loadExistingToken();
  }
  if (!accessToken) {
    throw new Error('No access token available. Please authenticate first using the "authenticate" tool.');
  }
  if (!graphClient) {
    initializeGraphClient();
  }
  return graphClient;
}

// ============================================================================
// HTML CONTENT PROCESSING UTILITIES
// ============================================================================

/**
 * Extracts readable plain text from HTML content.
 * Removes scripts, styles, and formats headings, paragraphs, lists, and tables.
 * @param {string} html - The HTML content string.
 * @returns {string} The extracted readable text.
 */
function extractReadableText(html) {
  try {
    if (!html) return '';
    const dom = new JSDOM(html);
    const document = dom.window.document;

    document.querySelectorAll('script, style').forEach(element => element.remove());

    let text = '';
    document.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach(heading => {
      const headingText = heading.textContent?.trim();
      if (headingText) text += `\n${headingText}\n${'-'.repeat(headingText.length)}\n`;
    });
    document.querySelectorAll('p').forEach(paragraph => {
      const content = paragraph.textContent?.trim();
      if (content) text += `${content}\n\n`;
    });
    document.querySelectorAll('ul, ol').forEach(list => {
      text += '\n';
      list.querySelectorAll('li').forEach((item, index) => {
        const content = item.textContent?.trim();
        if (content) text += `${list.tagName === 'OL' ? index + 1 + '.' : '-'} ${content}\n`;
      });
      text += '\n';
    });
    document.querySelectorAll('table').forEach(table => {
      text += '\n📊 Table content:\n';
      table.querySelectorAll('tr').forEach(row => {
        const cells = Array.from(row.querySelectorAll('td, th'))
          .map(cell => cell.textContent?.trim())
          .join(' | ');
        if (cells.trim()) text += `${cells}\n`;
      });
      text += '\n';
    });

    if (!text.trim() && document.body) {
      text = document.body.textContent?.trim().replace(/\s+/g, ' ') || '';
    }
    return text.trim();
  } catch (error) {
    console.error(`Error extracting readable text: ${error.message}`);
    return 'Error: Could not extract readable text from HTML content.';
  }
}

/**
 * Extracts a short summary from HTML content.
 * @param {string} html - The HTML content string.
 * @param {number} [maxLength=300] - The maximum length of the summary.
 * @returns {string} A text summary.
 */
function extractTextSummary(html, maxLength = 300) {
  try {
    if (!html) return 'No content to summarize.';
    const dom = new JSDOM(html);
    const document = dom.window.document;
    const bodyText = document.body?.textContent?.trim().replace(/\s+/g, ' ') || '';
    if (!bodyText) return 'No text content found in HTML body.';
    const summary = bodyText.substring(0, maxLength);
    return summary.length < bodyText.length ? `${summary}...` : summary;
  } catch (error) {
    console.error(`Error extracting text summary: ${error.message}`);
    return 'Could not extract text summary.';
  }
}

/**
 * Converts plain text (with simple markdown) to HTML.
 * @param {string} text - The plain text to convert.
 * @returns {string} The HTML representation.
 */
function textToHtml(text) {
  if (!text) return '';
  if (text.includes('<html>') || text.includes('<!DOCTYPE html>')) return text; // Already HTML

  let html = String(text) // Ensure text is a string
    .replace(/&/g, '&').replace(/</g, '<').replace(/>/g, '>') // Basic HTML escaping first
    .replace(/```([\s\S]*?)```/g, (match, code) => `<pre><code>${code.trim()}</code></pre>`)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/^# (.+)$/gm, '<h1>$1</h1>')
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>').replace(/__(.*?)__/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>').replace(/_(.*?)_/g, '<em>$1</em>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>')
    .replace(/^---+$/gm, '<hr>')
    .replace(/^> (.+)$/gm, '<blockquote>$1</blockquote>')
    .replace(/^[\*\-\+] (.+)$/gm, '<li>$1</li>')
    .replace(/^(\d+)\. (.+)$/gm, '<li>$2</li>');

  html = html.split('\n').map(line => {
    const trimmed = line.trim();
    if (!trimmed) return '';
    if (/^<(h[1-6]|li|hr|blockquote|pre|code|strong|em|a)/.test(trimmed) || /^<\/(h[1-6]|li|hr|blockquote|pre|code|strong|em|a)>/.test(trimmed)) {
      return trimmed; // Already an HTML element we processed or a closing tag
    }
    return `<p>${trimmed}</p>`;
  }).filter(line => line).join('\n');

  html = html.replace(/(<li>.*?<\/li>(?:\s*<li>.*?<\/li>)*)/gs, '<ul>$1</ul>');
  html = html.replace(/(<blockquote>.*?<\/blockquote>(?:\s*<blockquote>.*?<\/blockquote>)*)/gs, '<blockquote>$1</blockquote>');
  
  return html;
}

// ============================================================================
// ONENOTE API UTILITIES
// ============================================================================

/**
 * Implements exponential backoff retry logic for rate limiting.
 * @param {Function} apiCall - The async function to execute.
 * @param {number} maxRetries - Maximum number of retries (default: 3).
 * @returns {Promise<any>} The result of the API call.
 */
async function retryWithBackoff(apiCall, maxRetries = 3) {
  let lastError;
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return await apiCall();
    } catch (error) {
      lastError = error;
      const is429 = error.statusCode === 429 || error.message?.includes('429') || error.message?.includes('throttled');
      if (is429 && attempt < maxRetries - 1) {
        const waitTime = Math.pow(2, attempt) * 1000; // 1s, 2s, 4s
        console.error(`Rate limit hit, waiting ${waitTime}ms before retry ${attempt + 1}/${maxRetries - 1}...`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      } else {
        throw error;
      }
    }
  }
  throw lastError;
}

/**
 * Paginates through Microsoft Graph API results using @odata.nextLink.
 * @param {string} apiPath - The initial API path to query.
 * @param {number | null} maxPages - Maximum number of pages to fetch (null for unlimited).
 * @returns {Promise<Array>} Combined array of all results.
 */
async function paginateGraphRequest(apiPath, maxPages = null) {
  await ensureGraphClient();
  let allResults = [];
  let currentPath = apiPath;
  let pageCount = 0;

  while (currentPath && (maxPages === null || pageCount < maxPages)) {
    const response = await retryWithBackoff(() => graphClient.api(currentPath).get());
    
    if (response.value && Array.isArray(response.value)) {
      allResults = allResults.concat(response.value);
      pageCount++;
    }
    
    currentPath = response['@odata.nextLink'] || null;
    
    if (currentPath && pageCount % 10 === 0) {
      console.error(`Fetched ${pageCount} pages, ${allResults.length} items so far...`);
    }
  }
  
  return allResults;
}

/**
 * Queries multiple sections in parallel with concurrency limiting.
 * @param {Array} sections - Array of section objects to query.
 * @param {Function} queryFn - Async function that takes a section and returns pages.
 * @param {number} [concurrency=5] - Number of concurrent requests.
 * @returns {Promise<Array>} Array of page objects with _section and _notebook added.
 */
async function queryAllSectionsParallel(sections, queryFn, concurrency = 5) {
  let allPages = [];
  
  for (let i = 0; i < sections.length; i += concurrency) {
    const batch = sections.slice(i, i + concurrency);
    const batchResults = await Promise.allSettled(
      batch.map(section => queryFn(section))
    );
    
    // Process results
    batchResults.forEach((result, idx) => {
      if (result.status === 'fulfilled') {
        allPages = allPages.concat(result.value);
      } else {
        const section = batch[idx];
        console.error(`Error querying section "${section.displayName || section.name}": ${result.reason?.message}`);
      }
    });
  }
  
  return allPages;
}

/**
 * Searches sections in parallel for a page matching a title, with early exit.
 * @param {Array} sections - Array of section objects to search.
 * @param {string} searchTerm - Lowercase title substring to match.
 * @param {number} [concurrency=5] - Number of concurrent requests.
 * @returns {Promise<object|null>} The first matching page, or null.
 */
async function findPageInSectionsParallel(sections, searchTerm, concurrency = 5) {
  for (let i = 0; i < sections.length; i += concurrency) {
    const batch = sections.slice(i, i + concurrency);
    const batchResults = await Promise.allSettled(
      batch.map(async (section) => {
        const pages = await paginateGraphRequest(
          `/me/onenote/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=50`
        );
        return pages.find(p => p.title && p.title.toLowerCase().includes(searchTerm));
      })
    );

    for (const result of batchResults) {
      if (result.status === 'fulfilled' && result.value) {
        return result.value;
      }
    }
  }
  return null;
}

/**
 * Fetches the content of a OneNote page.
 * @param {string} pageId - The ID of the page.
 * @param {'httpDirect' | 'direct'} [method='httpDirect'] - The method to use for fetching.
 * @returns {Promise<string>} The HTML content of the page.
 */
async function fetchPageContentAdvanced(pageId, method = 'httpDirect') {
  await ensureGraphClient();
  if (method === 'httpDirect') {
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
    try {
      const response = await fetch(url, { 
        headers: { 'Authorization': `Bearer ${accessToken}` },
        signal: AbortSignal.timeout(45000) // 45 second timeout, leaving headroom under 60s MCP limit
      });
      if (!response.ok) throw new Error(`HTTP error fetching page content! Status: ${response.status} ${response.statusText}`);
      return await response.text();
    } catch (error) {
      if (error.name === 'TimeoutError' || error.name === 'AbortError') {
        throw new Error(`Request timed out after 45 seconds while fetching page content. The page may be too large or the server is slow to respond.`);
      }
      throw error;
    }
  } else { // 'direct'
    return await graphClient.api(`/me/onenote/pages/${pageId}/content`).get();
  }
}

/**
 * Formats OneNote page/notebook information for display.
 * @param {object} page - The OneNote page/notebook object from Graph API.
 * @param {number | null} [index=null] - Optional index for numbered lists.
 * @param {boolean} [includeCreator=false] - Include creator/modifier information.
 * @returns {string} Formatted page information string.
 */
function formatPageInfo(page, index = null, includeCreator = false) {
  const prefix = index !== null ? `${index + 1}. ` : '';
  const title = page.title || page.displayName || 'Untitled';
  const webUrl = page.links?.oneNoteWebUrl?.href || '';
  const urlLine = webUrl ? `
   🔗 <${webUrl}>` : '';
  
  let creatorLine = '';
  if (includeCreator && page.createdBy?.user?.displayName) {
    creatorLine = `
   Created by: ${page.createdBy.user.displayName}`;
  }
  if (includeCreator && page.lastModifiedBy?.user?.displayName) {
    creatorLine += `
   Modified by: ${page.lastModifiedBy.user.displayName}`;
  }
  
  // Show team context if present
  const teamLine = page._teamName ? `
   Team: ${page._teamName}` : '';
  
  return `${prefix}**${title}**
   ID: ${page.id}
   Created: ${new Date(page.createdDateTime).toLocaleDateString()}
   Modified: ${new Date(page.lastModifiedDateTime).toLocaleDateString()}${creatorLine}${teamLine}${urlLine}`;
}

/**
 * Gets the correct API path for accessing notebook resources.
 * @param {string} notebookId - The notebook ID.
 * @param {string} resourcePath - The resource path (e.g., 'sections', 'pages').
 * @returns {Promise<string>} The correct API path.
 */
async function getNotebookApiPath(notebookId, resourcePath) {
  // Check cache first
  if (notebookCache) {
    const notebook = notebookCache.find(nb => nb.id === notebookId);
    if (notebook) {
      if (notebook._groupId) {
        return `/groups/${notebook._groupId}/onenote/notebooks/${notebookId}/${resourcePath}`;
      }
      return `/me/onenote/notebooks/${notebookId}/${resourcePath}`;
    }
  }
  
  // Try personal notebook path first
  try {
    await ensureGraphClient();
    await graphClient.api(`/me/onenote/notebooks/${notebookId}`).get();
    return `/me/onenote/notebooks/${notebookId}/${resourcePath}`;
  } catch (error) {
    // If that fails, search in team notebooks
    if (!notebookCache) {
      await refreshNotebookCache();
    }
    const notebook = notebookCache?.find(nb => nb.id === notebookId);
    if (notebook && notebook._groupId) {
      return `/groups/${notebook._groupId}/onenote/notebooks/${notebookId}/${resourcePath}`;
    }
    throw new Error(`Notebook ${notebookId} not found in personal or team notebooks. Error: ${error.message}`);
  }
}

/**
 * Loads notebook cache from disk if available and not expired.
 * @returns {boolean} True if cache was loaded successfully.
 */
function loadNotebookCacheFromDisk() {
  try {
    if (fs.existsSync(notebookCacheFilePath)) {
      const cacheData = JSON.parse(fs.readFileSync(notebookCacheFilePath, 'utf8'));
      const age = Date.now() - cacheData.timestamp;
      
      if (age < CACHE_TTL_MS) {
        notebookCache = cacheData.notebooks;
        cacheTimestamp = cacheData.timestamp;
        console.error(`Loaded notebook cache from disk: ${notebookCache.length} notebooks (age: ${Math.round(age / 1000)}s)`);
        return true;
      } else {
        console.error(`Notebook cache expired (age: ${Math.round(age / 1000)}s), will refresh`);
      }
    }
  } catch (error) {
    console.error(`Error loading notebook cache: ${error.message}`);
  }
  return false;
}

/**
 * Saves notebook cache to disk.
 */
function saveNotebookCacheToDisk() {
  try {
    const cacheData = {
      timestamp: cacheTimestamp,
      notebooks: notebookCache
    };
    fs.writeFileSync(notebookCacheFilePath, JSON.stringify(cacheData, null, 2));
    console.error(`Saved notebook cache to disk: ${notebookCache.length} notebooks`);
  } catch (error) {
    console.error(`Error saving notebook cache: ${error.message}`);
  }
}

/**
 * Refreshes the notebook cache with both personal and team notebooks.
 * Uses aggressive parallelization and progressive loading.
 * @param {boolean} includeTeams - Whether to include team notebooks.
 * @param {boolean} personalOnly - Return immediately after loading personal notebooks.
 * @returns {Promise<Array>} Array of all notebooks with group info.
 */
async function refreshNotebookCache(includeTeams = true, personalOnly = false) {
  await ensureGraphClient();
  const previousCache = notebookCache ? [...notebookCache] : null;
  let allNotebooks = [];
  
  // Get personal notebooks (fast, always complete)
  try {
    const ownedResponse = await graphClient.api('/me/onenote/notebooks').get();
    allNotebooks = (ownedResponse.value || []).map(nb => ({
      ...nb,
      _isPersonal: true,
      _groupId: null,
      _teamName: null
    }));
    
    // Update cache with personal notebooks immediately
    notebookCache = allNotebooks;
    cacheTimestamp = Date.now();
    
    if (personalOnly) {
      console.error(`Loaded ${allNotebooks.length} personal notebooks (team notebooks will load in background)`);
      // Trigger background load without waiting
      refreshTeamNotebooksBackground().catch(err => 
        console.error(`Background team load failed: ${err.message}`)
      );
      return allNotebooks;
    }
  } catch (error) {
    console.error(`Error fetching personal notebooks: ${error.message}`);
  }
  
  // Get team notebooks if requested (slow, may timeout)
  if (includeTeams && !teamNotebooksLoading) {
    teamNotebooksLoading = true;
    
    try {
      const groupsResponse = await graphClient
        .api('/me/joinedTeams')
        .select('id,displayName')
        .get();
      
      const teams = groupsResponse.value || [];
      console.error(`Fetching notebooks from ${teams.length} teams in parallel...`);
      
      // Remove batching - fetch ALL teams in parallel using Promise.allSettled
      const promises = teams.map(async (team) => {
        try {
          const teamNotebooks = await graphClient
            .api(`/groups/${team.id}/onenote/notebooks`)
            .get();
          
          if (teamNotebooks.value && teamNotebooks.value.length > 0) {
            return teamNotebooks.value.map(nb => ({
              ...nb,
              displayName: nb.displayName || nb.name || `${team.displayName} Notebook`,
              _isPersonal: false,
              _groupId: team.id,
              _teamName: team.displayName,
              _isFromTeam: true
            }));
          }
          return [];
        } catch (teamError) {
          console.error(`Error fetching notebooks for team ${team.displayName}: ${teamError.message}`);
          return [];
        }
      });
      
      // Use allSettled to not fail on individual team errors
      const results = await Promise.allSettled(promises);
      const teamNotebooks = results
        .filter(r => r.status === 'fulfilled')
        .flatMap(r => r.value);
      
      allNotebooks.push(...teamNotebooks);
      console.error(`Loaded ${teamNotebooks.length} team notebooks from ${teams.length} teams`);
      
    } catch (teamsError) {
      console.error(`Error getting team notebooks: ${teamsError.message}`);
    } finally {
      teamNotebooksLoading = false;
    }
  }
  
  // Only update cache if we got results, or if there was no prior cache
  if (allNotebooks.length > 0 || !previousCache || previousCache.length === 0) {
    notebookCache = allNotebooks;
    cacheTimestamp = Date.now();
    saveNotebookCacheToDisk(); // Persist to disk
  } else {
    // Restore previous cache to avoid losing data on throttle
    notebookCache = previousCache;
    console.error(`Skipping cache update: API returned 0 notebooks but existing cache has ${previousCache.length}`);
  }
  console.error(`Notebook cache refreshed: ${allNotebooks.length} notebooks (${allNotebooks.filter(nb => nb._isPersonal).length} personal, ${allNotebooks.filter(nb => !nb._isPersonal).length} team)`);
  
  return allNotebooks;
}

/**
 * Refreshes team notebooks in the background without blocking.
 * Updates cache progressively as teams complete.
 */
async function refreshTeamNotebooksBackground() {
  if (teamNotebooksLoading || !graphClient) return;
  
  teamNotebooksLoading = true;
  console.error('Starting background team notebook refresh...');
  
  try {
    await ensureGraphClient();
    const groupsResponse = await graphClient
      .api('/me/joinedTeams')
      .select('id,displayName')
      .get();
    
    const teams = groupsResponse.value || [];
    console.error(`Fetching notebooks from ${teams.length} teams in background...`);
    
    const promises = teams.map(async (team) => {
      try {
        const teamNotebooks = await graphClient
          .api(`/groups/${team.id}/onenote/notebooks`)
          .get();
        
        if (teamNotebooks.value && teamNotebooks.value.length > 0) {
          return teamNotebooks.value.map(nb => ({
            ...nb,
            displayName: nb.displayName || nb.name || `${team.displayName} Notebook`,
            _isPersonal: false,
            _groupId: team.id,
            _teamName: team.displayName,
            _isFromTeam: true
          }));
        }
        return [];
      } catch (teamError) {
        return [];
      }
    });
    
    const results = await Promise.allSettled(promises);
    const teamNotebooks = results
      .filter(r => r.status === 'fulfilled')
      .flatMap(r => r.value);
    
    // Merge with existing cache (keep personal notebooks)
    const personalNotebooks = notebookCache?.filter(nb => nb._isPersonal) || [];
    notebookCache = [...personalNotebooks, ...teamNotebooks];
    cacheTimestamp = Date.now();
    saveNotebookCacheToDisk();
    
    console.error(`Background refresh complete: ${teamNotebooks.length} team notebooks loaded`);
  } catch (error) {
    console.error(`Background team notebook refresh failed: ${error.message}`);
  } finally {
    teamNotebooksLoading = false;
  }
}

// ============================================================================
// MCP TOOL DEFINITIONS
// ============================================================================

// --- Authentication Tools ---

server.tool(
  'authenticate',
  {
    // No input parameters expected for this tool
  },
  async () => {
    try {
      console.error('Starting device code authentication...');
      let deviceCodeInfo = null;
      const credential = new DeviceCodeCredential({
        clientId: clientId,
        userPromptCallback: (info) => {
          deviceCodeInfo = info;
          console.error(`\n=== AUTHENTICATION REQUIRED ===\n${info.message}\n================================\n`);
        }
      });

      const authPromise = credential.getToken(scopes);
      await new Promise(resolve => setTimeout(resolve, 2000)); // Allow time for userPromptCallback

      if (deviceCodeInfo) {
        const authMessage = `🔐 **AUTHENTICATION REQUIRED**

Please complete the following steps:
1. **Open this URL in your browser:** https://microsoft.com/devicelogin
2. **Enter this code:** ${deviceCodeInfo.userCode}
3. **Sign in with your Microsoft account that has OneNote access.**
4. **After completing authentication, use the 'saveAccessToken' tool.**

Token will be saved automatically upon successful browser authentication.`;

        authPromise.then(tokenResponse => {
          accessToken = tokenResponse.token;
          const tokenData = {
            token: accessToken,
            clientId: clientId,
            scopes: scopes,
            createdAt: new Date().toISOString(),
            expiresOn: tokenResponse.expiresOnTimestamp ? new Date(tokenResponse.expiresOnTimestamp).toISOString() : null
          };
          fs.writeFileSync(tokenFilePath, JSON.stringify(tokenData, null, 2));
          console.error('Token saved successfully!');
          initializeGraphClient();
        }).catch(error => {
          console.error(`Background authentication failed: ${error.message}`);
        });
        
        return { content: [{ type: 'text', text: authMessage }] };
      } else {
        return { isError: true, content: [{ type: 'text', text: 'Could not retrieve device code information. Please try again or check console logs.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Authentication failed: ${error.message}` }] };
    }
  }
);
// Note: For the above tool, the Zod schema `z.object({}).describe(...)` was simplified to `{}` as per the user's specific finding
// about the SDK's `server.tool(name, {param: z.type()}, handler)` signature.
// If the SDK *does* support a top-level describe on the Zod object itself, that would be:
// `z.object({}).describe('Start the authentication flow...')`

server.tool(
  'saveAccessToken',
  {
    // No input parameters
  },
  async () => {
    try {
      loadExistingToken();
      if (accessToken) {
        initializeGraphClient();
        const testResponse = await graphClient.api('/me').get();
        return {
          content: [{
            type: 'text',
            text: `✅ **Authentication Successful!**
Token loaded and verified.
**Account Info:**
- Name: ${testResponse.displayName || 'Unknown'}
- Email: ${testResponse.userPrincipalName || 'Unknown'}
🚀 You can now use OneNote tools!`
          }]
        };
      } else {
        return { isError: true, content: [{ type: 'text', text: `❌ **No Token Found.** Please run the 'authenticate' tool first.` }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to load or verify token: ${error.message}` }] };
    }
  }
);

server.tool(
  'checkTokenScopes',
  {
    // No input parameters
  },
  async () => {
    try {
      loadExistingToken();
      if (!accessToken) {
        return { isError: true, content: [{ type: 'text', text: `❌ **No Token Found.** Please run the 'authenticate' tool first.` }] };
      }

      // Decode JWT token (format: header.payload.signature)
      const parts = accessToken.split('.');
      if (parts.length !== 3) {
        return { isError: true, content: [{ type: 'text', text: `❌ **Invalid Token Format.** Token doesn't appear to be a valid JWT.` }] };
      }

      // Decode the payload (base64url encoded)
      const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString('utf-8'));
      
      // Extract scope information
      const scopes = payload.scp ? payload.scp.split(' ') : [];
      const audience = payload.aud || 'Unknown';
      const issuer = payload.iss || 'Unknown';
      const expiration = payload.exp ? new Date(payload.exp * 1000).toISOString() : 'Unknown';
      const appId = payload.appid || payload.azp || 'Unknown';
      
      // Check for required OneNote scopes
      const requiredScopes = ['Notes.Read.All', 'Notes.ReadWrite.All'];
      const hasRequiredScopes = requiredScopes.some(scope => scopes.includes(scope));
      const missingScopes = requiredScopes.filter(scope => !scopes.includes(scope));
      
      let message = `🔍 **Token Scope Analysis**

**Granted Scopes:**
${scopes.length > 0 ? scopes.map(s => `  • ${s}`).join('\n') : '  (None found)'}

**Audience:** ${audience}
**App ID:** ${appId}
**Expires:** ${expiration}

**OneNote Access Check:**`;

      if (hasRequiredScopes) {
        message += `\n✅ Token has sufficient OneNote permissions!`;
      } else {
        message += `\n❌ **ISSUE FOUND:** Token is missing required scopes!

**Missing Scopes:**
${missingScopes.map(s => `  • ${s}`).join('\n')}

**To Fix This:**
1. The Azure AD app (${appId}) needs these API permissions configured:
   - Microsoft Graph > Notes.Read.All (Delegated)
   - Microsoft Graph > Notes.ReadWrite.All (Delegated)
2. Admin consent may be required for .All scopes
3. After updating permissions, re-run the \`authenticate\` tool to get a new token

**Current Scopes Requested by MCP Server:**
${scopes.map(s => `  • ${s}`).join('\n')}

The app registration controls which scopes can actually be granted, regardless of what the code requests.`;
      }
      
      return { content: [{ type: 'text', text: message }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to decode token: ${error.message}\n\nThis might mean the token format is unexpected or corrupted.` }] };
    }
  }
);

// --- Page Reading Tools ---

server.tool(
  'listNotebooks',
  {
    includeTeamNotebooks: z.boolean().default(false).describe('Include notebooks from Microsoft Teams you have joined. Does not include personally shared notebooks from OneDrive (Microsoft API limitation).').optional(),
    refresh: z.boolean().default(false).describe('Force refresh of notebook cache.').optional()
  },
  async ({ includeTeamNotebooks = false, refresh = false }) => {
    try {
      await ensureGraphClient();
      
      // Use cache if available and not expired
      const cacheExpired = !cacheTimestamp || (Date.now() - cacheTimestamp > CACHE_TTL_MS);
      
      if (refresh || cacheExpired || !notebookCache) {
        // Progressive loading: for team notebooks, return personal immediately
        if (includeTeamNotebooks && !refresh) {
          // Load personal notebooks immediately, teams in background
          await refreshNotebookCache(true, true); // personalOnly=true
        } else {
          // Full refresh (may timeout with many teams)
          await refreshNotebookCache(includeTeamNotebooks);
        }
      }
      
      let allNotebooks = notebookCache || [];
      
      // Filter by team notebooks preference if cache includes both
      if (!includeTeamNotebooks) {
        allNotebooks = allNotebooks.filter(nb => nb._isPersonal);
      }
      
      if (allNotebooks.length > 0) {
        const notebookList = allNotebooks.map((nb, i) => formatPageInfo(nb, i)).join('\n\n');
        const teamInfo = includeTeamNotebooks && teamNotebooksLoading 
          ? '\n\n⏳ Team notebooks loading in background...' 
          : '';
        const cacheInfo = refresh ? '' : '\n\n💡 Use `refresh: true` to update the notebook list.';
        return { content: [{ type: 'text', text: `📚 **Your OneNote Notebooks** (${allNotebooks.length} found):\n\n${notebookList}${teamInfo}${cacheInfo}` }] };
      } else {
        return { content: [{ type: 'text', text: '📚 No OneNote notebooks found.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: error.message.includes('authenticate') ? '🔐 Authentication Required. Run `authenticate` tool.' : `Failed to list notebooks: ${error.message}` }] };
    }
  }
);

server.tool(
  'listSections',
  {
    notebookId: z.string().describe('The ID of the notebook to list sections from.')
  },
  async ({ notebookId }) => {
    try {
      await ensureGraphClient();
      console.error(`Fetching sections for notebook ID: ${notebookId}`);
      
      // Get correct API path for this notebook (personal or team)
      const apiPath = await getNotebookApiPath(notebookId, 'sections');
      const sections = await paginateGraphRequest(apiPath);
      
      if (sections.length > 0) {
        const sectionList = sections.map((section, i) => `${i + 1}. **${section.displayName}**\n   ID: ${section.id}\n   Created: ${new Date(section.createdDateTime).toLocaleDateString()}`).join('\n\n');
        return { content: [{ type: 'text', text: `📂 **Sections in Notebook** (${sections.length} found):\n\n${sectionList}` }] };
      } else {
        return { content: [{ type: 'text', text: '📂 No sections found in this notebook.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to list sections: ${error.message}` }] };
    }
  }
);

server.tool(
  'listPagesInSection',
  {
    sectionId: z.string().describe('The ID of the section to list pages from.'),
    top: z.number().min(1).max(100).default(20).describe('Number of pages to return per request (max 100).').optional(),
    orderBy: z.enum(['created', 'modified']).default('modified').describe('Sort by created or modified date.').optional()
  },
  async ({ sectionId, top, orderBy }) => {
    try {
      await ensureGraphClient();
      const orderField = orderBy === 'created' ? 'createdDateTime' : 'lastModifiedDateTime';
      console.error(`Fetching pages from section ID: ${sectionId} (top: ${top}, orderBy: ${orderField})`);
      
      const apiPath = `/me/onenote/sections/${sectionId}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=${top}&$orderby=${orderField} desc`;
      const pages = await paginateGraphRequest(apiPath);
      
      if (pages.length > 0) {
        const pageList = pages.slice(0, 50).map((page, i) => formatPageInfo(page, i)).join('\n\n');
        const morePages = pages.length > 50 ? `\n\n... and ${pages.length - 50} more pages. Total: ${pages.length}` : '';
        return { content: [{ type: 'text', text: `📄 **Pages in Section** (${pages.length} found):\n\n${pageList}${morePages}` }] };
      } else {
        return { content: [{ type: 'text', text: '📄 No pages found in this section.' }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to list pages in section: ${error.message}` }] };
    }
  }
);

server.tool(
  'searchPages',
  {
    query: z.string().describe('Search term for page titles.').optional(),
    notebookId: z.string().describe('Optional: limit search to specific notebook ID.').optional(),
    notebookName: z.string().describe('Optional: limit search to notebook by name (case-insensitive partial match).').optional(),
    top: z.number().min(1).max(100).default(100).describe('Max pages per section to fetch (default: 100, max: 100).').optional(),
    orderBy: z.enum(['created', 'modified', 'title']).default('modified').describe('Sort order: "created", "modified" (default), or "title".').optional(),
    maxResults: z.number().min(1).max(200).default(50).describe('Max total results to return (default: 50).').optional()
  },
  async ({ query, notebookId, notebookName, top = 100, orderBy = 'modified', maxResults = 50 }) => {
    try {
      await ensureGraphClient();
      console.error(`Searching pages (top: ${top}, orderBy: ${orderBy})...`);
      
      // Determine notebooks to search
      let notebooks = [];
      if (notebookId) {
        // Search specific notebook by ID
        try {
          const nb = await graphClient.api(`/me/onenote/notebooks/${notebookId}`).get();
          notebooks = [nb];
        } catch (error) {
          return { isError: true, content: [{ type: 'text', text: `Notebook ID "${notebookId}" not found. Use listNotebooks to get valid IDs.` }] };
        }
      } else if (notebookName) {
        // Search by name (partial match, case-insensitive)
        const allNotebooks = await graphClient.api('/me/onenote/notebooks').get();
        const searchName = notebookName.toLowerCase();
        notebooks = (allNotebooks.value || []).filter(nb => {
          const name = (nb.displayName || nb.name || '').toLowerCase();
          return name.includes(searchName);
        });
        
        if (notebooks.length === 0) {
          return { content: [{ type: 'text', text: `🔍 No notebooks found matching "${notebookName}".` }] };
        }
        console.error(`Found ${notebooks.length} notebook(s) matching "${notebookName}"`);
      } else {
        // Search all notebooks
        const notebooksResponse = await graphClient.api('/me/onenote/notebooks').get();
        notebooks = notebooksResponse.value || [];
      }
      
      if (notebooks.length === 0) {
        return { content: [{ type: 'text', text: '📚 No notebooks found.' }] };
      }
      
      // Map orderBy to API field
      const orderByField = orderBy === 'created' ? 'createdDateTime' : orderBy === 'modified' ? 'lastModifiedDateTime' : 'title';
      
      let allPages = [];
      let sectionsSearched = 0;
      
      // Iterate through notebooks and query sections with optimization
      for (const notebook of notebooks) {
        const sections = await paginateGraphRequest(`/me/onenote/notebooks/${notebook.id}/sections`);
        sectionsSearched += sections.length;
        
        const pages = await queryAllSectionsParallel(sections, async (section) => {
          // Use $orderby and $top for server-side optimization
          const apiPath = `/me/onenote/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=${top}&$orderby=${orderByField} desc`;
          return await paginateGraphRequest(apiPath);
        });
        
        // Add notebook name to each page for display
        pages.forEach(page => {
          page._notebook = notebook.displayName || notebook.name;
        });
        
        allPages = allPages.concat(pages);
      }
      
      console.error(`Fetched ${allPages.length} pages from ${sectionsSearched} sections`);
      
      // Filter by query if provided
      let filteredPages = allPages;
      if (query) {
        const searchTerm = query.toLowerCase();
        filteredPages = allPages.filter(page => page.title && page.title.toLowerCase().includes(searchTerm));
        console.error(`Filtered to ${filteredPages.length} pages matching "${query}"`);
      }
      
      // Sort by the requested field (already sorted per section, but may need global sort)
      filteredPages.sort((a, b) => {
        if (orderBy === 'created') {
          return new Date(b.createdDateTime) - new Date(a.createdDateTime);
        } else if (orderBy === 'modified') {
          return new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime);
        } else { // title
          return (a.title || '').localeCompare(b.title || '');
        }
      });
      
      // Limit to maxResults
      const displayPages = filteredPages.slice(0, maxResults);
      
      if (filteredPages.length > 0) {
        const pageList = displayPages.map((page, i) => {
          const webUrl = page.links?.oneNoteWebUrl?.href || '';
          const created = new Date(page.createdDateTime).toLocaleString();
          const modified = new Date(page.lastModifiedDateTime).toLocaleString();
          
          return `${i + 1}. **${page.title}**
   📚 Notebook: ${page._notebook}
   🔗 <${webUrl}>
   📅 Created: ${created}
   🔄 Modified: ${modified}`;
        }).join('\n\n');
        
        const morePages = filteredPages.length > maxResults ? `\n\n... and ${filteredPages.length - maxResults} more pages. Use 'maxResults' parameter to see more.` : '';
        const scopeInfo = notebookName ? ` in notebooks matching "${notebookName}"` : notebookId ? ` in notebook` : ` (${notebooks.length} notebook${notebooks.length > 1 ? 's' : ''})`;
        const queryInfo = query ? ` for "${query}"` : '';
        const perfInfo = `\n\n⚡ Performance: Fetched top ${top} pages per section, sorted by ${orderBy}`;
        
        return { 
          content: [{ 
            type: 'text', 
            text: `🔍 **Search Results**${queryInfo}${scopeInfo}\nFound ${filteredPages.length} page${filteredPages.length !== 1 ? 's' : ''}, showing ${displayPages.length}:\n\n${pageList}${morePages}${perfInfo}` 
          }] 
        };
      } else {
        const scopeInfo = notebookName ? ` in notebooks matching "${notebookName}"` : notebookId ? ` in that notebook` : '';
        return { content: [{ type: 'text', text: query ? `🔍 No pages found matching "${query}"${scopeInfo}.` : `📄 No pages found${scopeInfo}.` }] };
      }
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to search pages: ${error.message}` }] };
    }
  }
);

server.tool(
  'searchPagesByDate',
  {
    days: z.number().min(1).default(1).describe('Number of days back to search (1 = today only).').optional(),
    query: z.string().describe('Optional keyword to filter page titles.').optional(),
    dateField: z.enum(['created', 'modified', 'both']).default('both').describe('Filter by created, modified, or both dates.').optional(),
    includeContent: z.boolean().default(false).describe('Include page content preview (slower).').optional(),
    notebookName: z.string().describe('Optional: limit search to specific notebook name.').optional(),
    includeTeamNotebooks: z.boolean().default(false).describe('Include team notebooks in search.').optional()
  },
  async ({ days, query, dateField, includeContent, notebookName, includeTeamNotebooks }) => {
    try {
      await ensureGraphClient();
      const threshold = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
      threshold.setHours(0, 0, 0, 0); // Start at midnight
      console.error(`Searching pages from last ${days} day(s) (since ${threshold.toLocaleString()})...`);
      
      // Get all notebooks from cache (or refresh)
      if (!notebookCache) {
        await refreshNotebookCache(includeTeamNotebooks);
      }
      let notebooks = notebookCache || [];
      
      // Filter by personal/team preference
      if (!includeTeamNotebooks) {
        notebooks = notebooks.filter(nb => nb._isPersonal);
      }
      
      // Filter by notebook name if specified
      if (notebookName) {
        const searchTerm = notebookName.toLowerCase();
        notebooks = notebooks.filter(nb => 
          (nb.displayName || nb.name || '').toLowerCase().includes(searchTerm)
        );
        
        if (notebooks.length === 0) {
          return { content: [{ type: 'text', text: `📚 No notebook found matching "${notebookName}". ${!includeTeamNotebooks ? 'Try setting includeTeamNotebooks: true to search team notebooks.' : ''}` }] };
        }
      }
      
      if (notebooks.length === 0) {
        return { content: [{ type: 'text', text: '📚 No notebooks found.' }] };
      }
      
      console.error(`Found ${notebooks.length} notebook(s), fetching sections...`);
      let allMatchingPages = [];
      let totalSectionsChecked = 0;
      
      // Iterate through notebooks and query sections in parallel
      for (const notebook of notebooks) {
        // Get correct API path for this notebook (personal or team)
        const sectionsPath = notebook._groupId 
          ? `/groups/${notebook._groupId}/onenote/notebooks/${notebook.id}/sections`
          : `/me/onenote/notebooks/${notebook.id}/sections`;
        
        let sections = [];
        try {
          sections = await paginateGraphRequest(sectionsPath);
        } catch (error) {
          console.error(`Error fetching sections for notebook ${notebook.displayName}: ${error.message}`);
          continue; // Skip this notebook
        }
        
        totalSectionsChecked += sections.length;
        
        // Calculate reasonable top limit based on days (estimate ~10 pages per day, max 100)
        const topLimit = Math.min(100, Math.max(20, days * 10));
        
        // Query all sections in parallel with optimized query
        const pages = await queryAllSectionsParallel(sections, async (section) => {
          // Use $orderby and $top for efficient server-side filtering
          let sectionPages = [];
          try {
            const response = await retryWithBackoff(() => 
              graphClient
                .api(`/me/onenote/sections/${section.id}/pages`)
                .orderby('lastModifiedDateTime desc')
                .top(topLimit)
                .get()
            );
            sectionPages = response.value || [];
          } catch (error) {
            console.error(`Error fetching pages for section ${section.displayName}: ${error.message}`);
            return [];
          }
          
          // Filter by date and query
          const matchingPages = sectionPages.filter(page => {
            const created = new Date(page.createdDateTime);
            const modified = new Date(page.lastModifiedDateTime);
            
            let dateMatch = false;
            if (dateField === 'created') {
              dateMatch = created >= threshold;
            } else if (dateField === 'modified') {
              dateMatch = modified >= threshold;
            } else { // 'both'
              dateMatch = created >= threshold || modified >= threshold;
            }
            
            if (!dateMatch) return false;
            
            // Optional keyword filter
            if (query) {
              const searchTerm = query.toLowerCase();
              if (!page.title || !page.title.toLowerCase().includes(searchTerm)) {
                return false;
              }
            }
            
            return true;
          });
          
          // Add notebook and section context
          matchingPages.forEach(page => {
            page._notebook = notebook.displayName || notebook.name || 'Untitled';
            page._section = section.displayName || section.name || 'Untitled';
          });
          
          return matchingPages;
        });
        
        allMatchingPages = allMatchingPages.concat(pages);
        console.error(`Checked notebook "${notebook.displayName || notebook.name}", ${allMatchingPages.length} matches so far...`);
      }
      
      console.error(`Search complete: ${allMatchingPages.length} matches from ${totalSectionsChecked} sections`);
      
      if (allMatchingPages.length === 0) {
        // Add debug info to help diagnose
        const debugInfo = `\n\n🔍 **Debug Info:**\n- Threshold: ${threshold.toLocaleString()}\n- Days back: ${days}\n- Notebooks checked: ${notebooks.length}\n- Sections checked: ${totalSectionsChecked}\n- Date field: ${dateField}`;
        return { content: [{ type: 'text', text: (query ? `🔍 No pages found matching "${query}" in the last ${days} day(s).` : `📄 No pages found in the last ${days} day(s).`) + debugInfo }] };
      }
      
      // Sort by most recent first
      allMatchingPages.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
      
      // Format results
      let resultText = `🔍 **Date Search Results** (${allMatchingPages.length} found in last ${days} day(s)):\n`;
      if (query) resultText += `Keyword: "${query}"\n`;
      if (notebookName) resultText += `Notebook: "${notebookName}"\n`;
      resultText += `Checked: ${notebooks.length} notebooks, ${totalSectionsChecked} sections\n\n`;
      
      const displayPages = allMatchingPages.slice(0, 20);
      for (let i = 0; i < displayPages.length; i++) {
        const page = displayPages[i];
        const webUrl = page.links?.oneNoteWebUrl?.href || '';
        resultText += `${i + 1}. **${page.title}**\n`;
        resultText += `   ID: ${page.id}\n`;
        resultText += `   📚 ${page._notebook} / 📂 ${page._section}\n`;
        if (webUrl) resultText += `   🔗 <${webUrl}>\n`;
        resultText += `   Created: ${new Date(page.createdDateTime).toLocaleString()}\n`;
        resultText += `   Modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}\n`;
        
        if (includeContent) {
          try {
            const htmlContent = await fetchPageContentAdvanced(page.id, 'httpDirect');
            const preview = extractTextSummary(htmlContent, 200);
            resultText += `   Preview: ${preview}\n`;
          } catch (contentError) {
            resultText += `   Preview: (error loading content)\n`;
          }
        }
        
        resultText += '\n';
      }
      
      if (allMatchingPages.length > 20) {
        resultText += `... and ${allMatchingPages.length - 20} more pages.`;
      }
      
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to search pages by date: ${error.message}` }] };
    }
  }
);

server.tool(
  'searchPageContent',
  {
    query: z.string().describe('Text to search for within page content.'),
    days: z.number().min(1).describe('Optional: limit to pages from last N days.').optional(),
    notebookId: z.string().describe('Optional: limit to specific notebook.').optional(),
    maxPages: z.number().min(1).max(50).default(20).describe('Max number of pages to search (default 20).').optional()
  },
  async ({ query, days, notebookId, maxPages }) => {
    try {
      await ensureGraphClient();
      const searchTerm = query.toLowerCase();
      const threshold = days ? new Date(Date.now() - days * 24 * 60 * 60 * 1000) : null;
      console.error(`Searching page content for "${query}"...`);
      
      let notebooks = [];
      if (notebookId) {
        const nb = await graphClient.api(`/me/onenote/notebooks/${notebookId}`).get();
        notebooks = [nb];
      } else {
        const response = await graphClient.api('/me/onenote/notebooks').get();
        notebooks = response.value || [];
      }
      
      let matchingPages = [];
      let pagesSearched = 0;
      
      for (const notebook of notebooks) {
        const sections = await paginateGraphRequest(`/me/onenote/notebooks/${notebook.id}/sections`);
        
        for (const section of sections) {
          if (matchingPages.length >= maxPages) break;
          
          try {
            const apiPath = `/me/onenote/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=20`;
            const pages = await paginateGraphRequest(apiPath);
            
            for (const page of pages) {
              if (matchingPages.length >= maxPages) break;
              
              // Date filter if specified
              if (threshold) {
                const modified = new Date(page.lastModifiedDateTime);
                if (modified < threshold) continue;
              }
              
              pagesSearched++;
              
              // Fetch and search content
              try {
                const htmlContent = await fetchPageContentAdvanced(page.id, 'httpDirect');
                const textContent = extractReadableText(htmlContent).toLowerCase();
                
                if (textContent.includes(searchTerm)) {
                  page._notebook = notebook.displayName || notebook.name;
                  page._section = section.displayName || section.name;
                  page._snippet = extractSnippet(textContent, searchTerm, 150);
                  matchingPages.push(page);
                }
              } catch (contentError) {
                console.error(`Error reading page ${page.title}: ${contentError.message}`);
              }
            }
          } catch (sectionError) {
            console.error(`Error in section ${section.displayName}: ${sectionError.message}`);
          }
        }
        
        if (matchingPages.length >= maxPages) break;
      }
      
      console.error(`Content search complete: ${matchingPages.length} matches from ${pagesSearched} pages searched`);
      
      if (matchingPages.length === 0) {
        return { content: [{ type: 'text', text: `🔍 No pages found containing "${query}" in their content.\nSearched ${pagesSearched} pages.` }] };
      }
      
      let resultText = `🔍 **Content Search Results** (${matchingPages.length} matches for "${query}"):\n`;
      resultText += `Searched ${pagesSearched} pages\n\n`;
      
      for (let i = 0; i < matchingPages.length; i++) {
        const page = matchingPages[i];
        const webUrl = page.links?.oneNoteWebUrl?.href || '';
        resultText += `${i + 1}. **${page.title}**\n`;
        resultText += `   ID: ${page.id}\n`;
        resultText += `   📚 ${page._notebook} / 📂 ${page._section}\n`;
        if (webUrl) resultText += `   🔗 <${webUrl}>\n`;
        resultText += `   Modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}\n`;
        resultText += `   Snippet: "...${page._snippet}..."\n\n`;
      }
      
      if (matchingPages.length >= maxPages) {
        resultText += `\n💡 Showing first ${maxPages} matches. Use 'maxPages' parameter or narrow search with 'days' or 'notebookId'.`;
      }
      
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to search page content: ${error.message}` }] };
    }
  }
);

/**
 * Extracts a snippet around the search term.
 * @param {string} text - The full text to search.
 * @param {string} searchTerm - The term to find.
 * @param {number} maxLength - Maximum snippet length.
 * @returns {string} The snippet.
 */
function extractSnippet(text, searchTerm, maxLength = 150) {
  const index = text.indexOf(searchTerm);
  if (index === -1) return text.substring(0, maxLength);
  
  const start = Math.max(0, index - Math.floor(maxLength / 2));
  const end = Math.min(text.length, start + maxLength);
  let snippet = text.substring(start, end).trim();
  
  if (start > 0) snippet = '...' + snippet;
  if (end < text.length) snippet = snippet + '...';
  
  return snippet;
}

server.tool(
  'getMyRecentChanges',
  {
    sinceDate: z.string().describe('Start date in YYYY-MM-DD format, or "monday" for last Monday.').optional(),
    days: z.number().min(1).describe('Alternative: number of days back to search (e.g., 3 for last 3 days).').optional(),
    notebookId: z.string().describe('Optional: limit to specific notebook.').optional()
  },
  async ({ sinceDate, days, notebookId }) => {
    try {
      await ensureGraphClient();
      
      // Calculate threshold date
      let threshold;
      if (sinceDate === 'monday') {
        const today = new Date();
        const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, ...
        const daysBack = dayOfWeek === 0 ? 6 : dayOfWeek === 1 ? 0 : dayOfWeek - 1;
        threshold = new Date(today.getTime() - daysBack * 24 * 60 * 60 * 1000);
        threshold.setHours(0, 0, 0, 0);
      } else if (sinceDate) {
        threshold = new Date(sinceDate);
        threshold.setHours(0, 0, 0, 0);
      } else if (days) {
        threshold = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
        threshold.setHours(0, 0, 0, 0);
      } else {
        // Default to last 3 days
        threshold = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
        threshold.setHours(0, 0, 0, 0);
      }
      
      // Get notebooks
      let notebooks = [];
      if (notebookId) {
        // Get specific notebook from cache or API
        if (!notebookCache) {
          await refreshNotebookCache(true);
        }
        const cachedNotebook = notebookCache?.find(nb => nb.id === notebookId);
        if (cachedNotebook) {
          notebooks = [cachedNotebook];
        } else {
          // Try to fetch it directly
          try {
            const nb = await graphClient.api(`/me/onenote/notebooks/${notebookId}`).get();
            notebooks = [{ ...nb, _isPersonal: true, _groupId: null }];
          } catch (error) {
            return { isError: true, content: [{ type: 'text', text: `Notebook ${notebookId} not found. Please use listNotebooks to get valid IDs.` }] };
          }
        }
      } else {
        // Use cached notebooks to avoid timeout
        if (!notebookCache) {
          await refreshNotebookCache(true);
        }
        notebooks = notebookCache || [];
        
        if (notebooks.length > 30) {
          return { 
            isError: true, 
            content: [{ 
              type: 'text', 
              text: `⚠️ You have ${notebooks.length} notebooks. Scanning all would take too long.\n\n💡 **Use notebookId parameter:**\n1. Run \`listNotebooks\` to see your notebooks\n2. Call \`getMyRecentChanges\` with a specific \`notebookId\`\n\nExample: getMyRecentChanges(days: 7, notebookId: "your-notebook-id")` 
            }] 
          };
        }
      }
      
      let recentPages = [];
      
      for (const notebook of notebooks) {
        // Get correct API path for this notebook (personal or team)
        const apiPath = notebook._groupId 
          ? `/groups/${notebook._groupId}/onenote/notebooks/${notebook.id}/sections`
          : `/me/onenote/notebooks/${notebook.id}/sections`;
        
        let sections = [];
        try {
          sections = await paginateGraphRequest(apiPath);
        } catch (error) {
          console.error(`Error fetching sections for notebook ${notebook.displayName}: ${error.message}`);
          continue; // Skip this notebook if we can't access it
        }
        
        // Calculate reasonable top limit based on days (estimate ~10 pages per day, max 100)
        const daysCount = days || (sinceDate ? Math.ceil((Date.now() - threshold.getTime()) / (24 * 60 * 60 * 1000)) : 3);
        const topLimit = Math.min(100, Math.max(20, daysCount * 10));
        
        // Query all sections in parallel with optimized query
        const pages = await queryAllSectionsParallel(sections, async (section) => {
          // Use $orderby and $top for efficient server-side filtering
          // This fetches only the most recently modified pages instead of ALL pages
          let sectionPages = [];
          try {
            const response = await retryWithBackoff(() => 
              graphClient
                .api(`/me/onenote/sections/${section.id}/pages`)
                .orderby('lastModifiedDateTime desc')
                .top(topLimit)
                .get()
            );
            sectionPages = response.value || [];
          } catch (error) {
            console.error(`Error fetching pages for section ${section.displayName}: ${error.message}`);
            return [];
          }
          
          // Filter by modification date (most should pass since we're ordering by desc)
          const matchingPages = sectionPages.filter(page => {
            const modified = new Date(page.lastModifiedDateTime);
            return modified >= threshold;
          });
          
          // Add notebook and section context
          matchingPages.forEach(page => {
            page._notebook = notebook.displayName || notebook.name || 'Untitled';
            page._section = section.displayName || section.name || 'Untitled';
            page._isFromTeam = notebook._isFromTeam || false;
          });
          
          return matchingPages;
        });
        
        recentPages = recentPages.concat(pages);
      }
      
      
      if (recentPages.length === 0) {
        return { 
          content: [{ 
            type: 'text', 
            text: `📝 No pages modified in your notebooks since ${threshold.toLocaleDateString()}.` 
          }] 
        };
      }
      
      // Sort by most recent
      recentPages.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
      
      let resultText = `📝 **Recent Changes in Your Notebooks** (${recentPages.length} pages since ${threshold.toLocaleDateString()}):\n\n`;
      resultText += `_Note: Includes changes by you and collaborators in your notebooks._\n`;
      
      // Group by notebook
      const byNotebook = {};
      recentPages.forEach(page => {
        const nb = page._notebook;
        if (!byNotebook[nb]) byNotebook[nb] = [];
        byNotebook[nb].push(page);
      });
      
      for (const [notebookName, pages] of Object.entries(byNotebook)) {
        resultText += `\n📚 **${notebookName}** (${pages.length} pages)\n`;
        
        pages.forEach((page, i) => {
          const webUrl = page.links?.oneNoteWebUrl?.href || '';
          resultText += `\n${i + 1}. **${page.title}**\n`;
          resultText += `   📂 ${page._section}\n`;
          if (webUrl) resultText += `   🔗 <${webUrl}>\n`;
          resultText += `   Modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}\n`;
        });
      }
      
      resultText += `\n\n💡 **Standup Tip:** Use this list to fill out your "Async" section!`;
      
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to get recent changes: ${error.message}` }] };
    }
  }
);

server.tool(
  'createDailyNote',
  {
    notebookName: z.string().describe('Name of the notebook (e.g., "SQLNikon").'),
    sectionName: z.string().describe('Name of the section (e.g., "Elon").'),
    date: z.string().describe('Date in YYYY-MM-DD format, or "today" for current date.').optional(),
    title: z.string().describe('Optional custom title (defaults to M/D/YY format).').optional(),
    content: z.string().describe('Optional initial content for the note.').optional()
  },
  async ({ notebookName, sectionName, date, title, content }) => {
    try {
      await ensureGraphClient();
      
      // Determine the date
      let noteDate;
      if (date === 'today' || !date) {
        noteDate = new Date();
      } else {
        noteDate = new Date(date);
      }
      
      // Format title as M/D/YY (e.g., "2/20/26")
      const defaultTitle = `${noteDate.getMonth() + 1}/${noteDate.getDate()}/${noteDate.getFullYear().toString().slice(-2)}`;
      const pageTitle = title || defaultTitle;
      
      console.error(`Creating daily note "${pageTitle}" in ${notebookName}/${sectionName}...`);
      
      // Find the notebook
      const notebooks = await graphClient.api('/me/onenote/notebooks').get();
      const notebook = (notebooks.value || []).find(nb => 
        nb.displayName?.toLowerCase() === notebookName.toLowerCase()
      );
      
      if (!notebook) {
        return { isError: true, content: [{ type: 'text', text: `❌ Notebook "${notebookName}" not found. Available notebooks:\n${(notebooks.value || []).map(nb => `  - ${nb.displayName}`).join('\n')}` }] };
      }
      
      // Find the section
      const sections = await paginateGraphRequest(`/me/onenote/notebooks/${notebook.id}/sections`);
      const section = sections.find(s => 
        s.displayName?.toLowerCase() === sectionName.toLowerCase()
      );
      
      if (!section) {
        return { isError: true, content: [{ type: 'text', text: `❌ Section "${sectionName}" not found in notebook "${notebookName}". Available sections:\n${sections.map(s => `  - ${s.displayName}`).join('\n')}` }] };
      }
      
      // Check if page already exists
      const existingPages = await paginateGraphRequest(`/me/onenote/sections/${section.id}/pages?$select=id,title`);
      const existingPage = existingPages.find(p => p.title === pageTitle);
      
      if (existingPage) {
        return { content: [{ type: 'text', text: `ℹ️ **Note Already Exists**\n\nA page titled "${pageTitle}" already exists in ${notebookName}/${sectionName}.\n\nPage ID: ${existingPage.id}\n\nUse 'appendToPage' or 'updatePageContent' to modify it.` }] };
      }
      
      // Create the page
      const htmlContent = content ? textToHtml(content) : '<p>Notes for today...</p>';
      const pageHtml = `<!DOCTYPE html>
<html>
<head>
  <title>${textToHtml(pageTitle)}</title>
  <meta charset="utf-8">
</head>
<body>
  <h1>${textToHtml(pageTitle)}</h1>
  ${htmlContent}
  <hr>
  <p><em>Created via OneNote MCP on ${new Date().toLocaleString()}</em></p>
</body>
</html>`;
      
      const response = await graphClient
        .api(`/me/onenote/sections/${section.id}/pages`)
        .header('Content-Type', 'application/xhtml+xml')
        .post(pageHtml);
      
      const webUrl = response.links?.oneNoteWebUrl?.href || '';
      
      return {
        content: [{
          type: 'text',
          text: `✅ **Daily Note Created!**\n\n**Title:** ${response.title}\n**Notebook:** ${notebookName}\n**Section:** ${sectionName}\n**Page ID:** ${response.id}\n**Created:** ${new Date(response.createdDateTime).toLocaleString()}${webUrl ? `\n\n🔗 <${webUrl}>` : ''}`
        }]
      };
    } catch (error) {
      console.error(`CREATE DAILY NOTE ERROR: ${error.message}`, error.stack);
      return { isError: true, content: [{ type: 'text', text: `❌ **Error creating daily note:** ${error.message}` }] };
    }
  }
);

server.tool(
  'searchInNotebook',
  {
    notebookId: z.string().describe('The ID of the notebook to search within.'),
    query: z.string().describe('Optional keyword to filter page titles.').optional(),
    days: z.number().min(1).describe('Optional: search pages from last N days.').optional(),
    top: z.number().min(1).max(100).default(100).describe('Max pages per section.').optional()
  },
  async ({ notebookId, query, days, top }) => {
    try {
      await ensureGraphClient();
      console.error(`Searching in notebook ID: ${notebookId}`);
      
      // Get notebook info and correct API path
      let notebook = null;
      let apiPath = null;
      
      try {
        apiPath = await getNotebookApiPath(notebookId, 'sections');
        // Get notebook from cache
        if (!notebookCache) {
          await refreshNotebookCache(true);
        }
        notebook = notebookCache?.find(nb => nb.id === notebookId);
        
        if (!notebook) {
          // Try to fetch directly
          try {
            notebook = await graphClient.api(`/me/onenote/notebooks/${notebookId}`).get();
          } catch (e) {
            // Might be a team notebook
            return { isError: true, content: [{ type: 'text', text: `Notebook ${notebookId} not found. Please use listNotebooks to get valid IDs.` }] };
          }
        }
      } catch (error) {
        return { isError: true, content: [{ type: 'text', text: `Failed to access notebook: ${error.message}` }] };
      }
      
      const sections = await paginateGraphRequest(apiPath);
      
      if (sections.length === 0) {
        return { content: [{ type: 'text', text: `📂 No sections found in notebook "${notebook.displayName || notebook.name}".` }] };
      }
      
      console.error(`Found ${sections.length} section(s), searching pages...`);
      const threshold = days ? new Date(Date.now() - days * 24 * 60 * 60 * 1000) : null;
      
      // Query all sections in parallel
      const allMatchingPages = await queryAllSectionsParallel(sections, async (section) => {
        const apiPath = `/me/onenote/sections/${section.id}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=${top}`;
        const pages = await paginateGraphRequest(apiPath);
        
        const matchingPages = pages.filter(page => {
          // Date filter if specified
          if (threshold) {
            const created = new Date(page.createdDateTime);
            const modified = new Date(page.lastModifiedDateTime);
            if (created < threshold && modified < threshold) {
              return false;
            }
          }
          
          // Keyword filter if specified
          if (query) {
            const searchTerm = query.toLowerCase();
            return page.title && page.title.toLowerCase().includes(searchTerm);
          }
          
          return true;
        });
        
        matchingPages.forEach(page => {
          page._section = section.displayName || section.name;
        });
        
        return matchingPages;
      });
      
      console.error(`Search complete: ${allMatchingPages.length} matches`);
      
      if (allMatchingPages.length === 0) {
        return { content: [{ type: 'text', text: `🔍 No matching pages found in notebook "${notebook.displayName}".` }] };
      }
      
      // Sort by most recent
      allMatchingPages.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
      
      let resultText = `🔍 **Search Results in "${notebook.displayName}"** (${allMatchingPages.length} found):\n`;
      if (query) resultText += `Keyword: "${query}"\n`;
      if (days) resultText += `Time range: Last ${days} day(s)\n`;
      resultText += `Checked: ${sections.length} sections\n\n`;
      
      const displayPages = allMatchingPages.slice(0, 20);
      for (let i = 0; i < displayPages.length; i++) {
        const page = displayPages[i];
        const webUrl = page.links?.oneNoteWebUrl?.href || '';
        resultText += `${i + 1}. **${page.title}**\n`;
        resultText += `   📂 ${page._section}\n`;
        if (webUrl) resultText += `   🔗 <${webUrl}>\n`;
        resultText += `   Modified: ${new Date(page.lastModifiedDateTime).toLocaleString()}\n\n`;
      }
      
      if (allMatchingPages.length > 20) {
        resultText += `... and ${allMatchingPages.length - 20} more pages.`;
      }
      
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to search in notebook: ${error.message}` }] };
    }
  }
);

server.tool(
  'getPageContent',
  {
    pageId: z.string().describe('The ID of the page to retrieve content from.'),
    format: z.enum(['text', 'html', 'summary'])
      .default('text')
      .describe('Format of the content: text (readable), html (raw), or summary (brief).')
      .optional()
  },
  async ({ pageId, format }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      const htmlContent = await fetchPageContentAdvanced(pageId, 'httpDirect');
      const webUrl = pageInfo.links?.oneNoteWebUrl?.href || '';
      let resultText = '';

      if (format === 'html') {
        resultText = `📄 **${pageInfo.title}** (HTML Format)${webUrl ? `\n🔗 <${webUrl}>` : ''}\n\n${htmlContent}`;
      } else if (format === 'summary') {
        const summary = extractTextSummary(htmlContent, 300);
        resultText = `📄 **${pageInfo.title}** (Summary)${webUrl ? `\n🔗 <${webUrl}>` : ''}\n\n${summary}`;
      } else { // 'text'
        const textContent = extractReadableText(htmlContent);
        resultText = `📄 **${pageInfo.title}**\n📅 Modified: ${new Date(pageInfo.lastModifiedDateTime).toLocaleString()}${webUrl ? `\n🔗 <${webUrl}>` : ''}\n\n${textContent}`;
      }
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to get page content for ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'getPageByTitle',
  {
    title: z.string().describe('The title (or partial title) of the page to find.'),
    format: z.enum(['text', 'html', 'summary'])
      .default('text')
      .describe('Format of the content: text, html, or summary.')
      .optional()
  },
  async ({ title, format }) => {
    try {
      await ensureGraphClient();
      const searchTerm = title.toLowerCase();
      let matchingPage = null;

      // Strategy 1: Use flat /me/onenote/pages endpoint with OData filter (single API call)
      try {
        const filterQuery = `/me/onenote/pages?$filter=contains(tolower(title),'${searchTerm.replace(/'/g, "''")}')&$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=1&$orderby=lastModifiedDateTime desc`;
        const searchResponse = await graphClient.api(filterQuery).get();
        const results = searchResponse.value || [];
        if (results.length > 0) {
          matchingPage = results[0];
          console.error(`Found page via OData filter: "${matchingPage.title}"`);
        }
      } catch (filterError) {
        console.error(`OData filter not supported: ${filterError.message}`);
      }

      // Strategy 2: Paginate through /me/onenote/pages (flat list, no notebook/section iteration)
      if (!matchingPage) {
        console.error(`Falling back to flat pages list scan for "${title}"...`);
        let nextLink = `/me/onenote/pages?$select=id,title,createdDateTime,lastModifiedDateTime,links&$top=100&$orderby=lastModifiedDateTime desc`;
        let pagesScanned = 0;

        while (nextLink && !matchingPage) {
          const response = await retryWithBackoff(() => graphClient.api(nextLink).get());
          const pages = response.value || [];
          pagesScanned += pages.length;

          matchingPage = pages.find(p => p.title && p.title.toLowerCase().includes(searchTerm));

          if (!matchingPage) {
            nextLink = response['@odata.nextLink'] || null;
            // Cap at 500 pages to avoid excessive API calls
            if (pagesScanned >= 500) {
              console.error(`Scanned ${pagesScanned} pages without finding a match, stopping.`);
              break;
            }
          }
        }
        if (matchingPage) {
          console.error(`Found page via flat scan after ${pagesScanned} pages: "${matchingPage.title}"`);
        }
      }

      if (!matchingPage) {
        return { isError: true, content: [{ type: 'text', text: `❌ No page found with title containing "${title}".` }] };
      }

      const htmlContent = await fetchPageContentAdvanced(matchingPage.id, 'httpDirect');
      const webUrl = matchingPage.links?.oneNoteWebUrl?.href || '';
      let resultText = '';
      if (format === 'html') {
        resultText = `📄 **${matchingPage.title}** (HTML Format)${webUrl ? `\n🔗 <${webUrl}>` : ''}\n\n${htmlContent}`;
      } else if (format === 'summary') {
        const summary = extractTextSummary(htmlContent, 300);
        resultText = `📄 **${matchingPage.title}** (Summary)${webUrl ? `\n🔗 <${webUrl}>` : ''}\n\n${summary}`;
      } else { // 'text'
        const textContent = extractReadableText(htmlContent);
        resultText = `📄 **${matchingPage.title}**\n📅 Modified: ${new Date(matchingPage.lastModifiedDateTime).toLocaleString()}${webUrl ? `\n🔗 <${webUrl}>` : ''}\n\n${textContent}`;
      }
      return { content: [{ type: 'text', text: resultText }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `Failed to get page by title "${title}": ${error.message}` }] };
    }
  }
);

// --- Page Editing & Content Manipulation Tools ---

server.tool(
  'updatePageContent',
  {
    pageId: z.string().describe('The ID of the page to update.'),
    content: z.string().describe('New page content (HTML or markdown-style text).'),
    preserveTitle: z.boolean()
      .default(true)
      .describe('Keep the original title (default: true).')
      .optional()
  },
  async ({ pageId, content: newContent, preserveTitle }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error(`Updating content for page: "${pageInfo.title}" (ID: ${pageId})`);
      
      const htmlContentForUpdate = textToHtml(newContent);
      const finalHtml = `
        <div>
          ${preserveTitle ? `<h1>${pageInfo.title}</h1>` : ''}
          ${htmlContentForUpdate}
          <hr>
          <p><em>Updated via OneNote MCP on ${new Date().toLocaleString()}</em></p>
        </div>
      `;
      
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: 'replace', content: finalHtml }]),
        signal: AbortSignal.timeout(45000)
      });
      
      if (!response.ok) throw new Error(`Update failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Page Content Updated!**\nPage: ${pageInfo.title}\nUpdated: ${new Date().toLocaleString()}\nContent Length: ${newContent.length} chars.` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to update page content for ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'appendToPage',
  {
    pageId: z.string().describe('The ID of the page to append content to.'),
    content: z.string().describe('Content to append (HTML or markdown-style).'),
    addTimestamp: z.boolean().default(true).describe('Add a timestamp (default: true).').optional(),
    addSeparator: z.boolean().default(true).describe('Add a visual separator (default: true).').optional()
  },
  async ({ pageId, content: newContent, addTimestamp, addSeparator }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error(`Appending content to page: "${pageInfo.title}" (ID: ${pageId})`);
      
      const htmlContentToAppend = textToHtml(newContent);
      let appendHtml = '';
      if (addSeparator) appendHtml += '<hr>';
      if (addTimestamp) appendHtml += `<p><em>Added on ${new Date().toLocaleString()}</em></p>`;
      appendHtml += htmlContentToAppend;
      
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: 'append', content: appendHtml }]),
        signal: AbortSignal.timeout(45000)
      });
      
      if (!response.ok) throw new Error(`Append failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Content Appended!**\nPage: ${pageInfo.title}\nAppended: ${new Date().toLocaleString()}\nLength: ${newContent.length} chars.` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to append content to page ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'updatePageTitle',
  {
    pageId: z.string().describe('The ID of the page whose title is to be updated.'),
    newTitle: z.string().describe('The new title for the page.')
  },
  async ({ pageId, newTitle }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      const oldTitle = pageInfo.title;
      console.error(`Updating page title from "${oldTitle}" to "${newTitle}" for page ID "${pageId}"`);
      
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'title', action: 'replace', content: newTitle }]),
        signal: AbortSignal.timeout(45000)
      });
      
      if (!response.ok) throw new Error(`Title update failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Page Title Updated!**\nOld Title: ${oldTitle}\nNew Title: ${newTitle}\nUpdated: ${new Date().toLocaleString()}` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to update page title for ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'replaceTextInPage',
  {
    pageId: z.string().describe('The ID of the page to modify.'),
    findText: z.string().describe('The text to find and replace.'),
    replaceText: z.string().describe('The text to replace with.'),
    caseSensitive: z.boolean().default(false).describe('Case-sensitive search (default: false).').optional()
  },
  async ({ pageId, findText, replaceText, caseSensitive }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      const htmlContent = await fetchPageContentAdvanced(pageId, 'httpDirect');
      console.error(`Replacing text in page: "${pageInfo.title}" (ID: ${pageId})`);
      
      const flags = caseSensitive ? 'g' : 'gi';
      const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
      const matches = (htmlContent.match(regex) || []).length;
      
      if (matches === 0) {
        return { content: [{ type: 'text', text: `ℹ️ **No matches found** for "${findText}" in page: ${pageInfo.title}.` }] };
      }
      
      const updatedContent = htmlContent.replace(regex, replaceText);
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: 'replace', content: `<div>${updatedContent}</div>` }]),
        signal: AbortSignal.timeout(45000)
      });
      
      if (!response.ok) throw new Error(`Replace failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Text Replaced!**\nPage: ${pageInfo.title}\nFound: "${findText}" (${matches} occurrences)\nReplaced with: "${replaceText}".` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to replace text in page ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'addNoteToPage',
  {
    pageId: z.string().describe('The ID of the page to add a note to.'),
    note: z.string().describe('The note/comment content.'),
    noteType: z.enum(['note', 'todo', 'important', 'question'])
      .default('note')
      .describe('Type of note (note, todo, important, question).')
      .optional(),
    position: z.enum(['top', 'bottom'])
      .default('bottom')
      .describe('Position to add the note (top or bottom).')
      .optional()
  },
  async ({ pageId, note, noteType, position }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error(`Adding ${noteType} to page: "${pageInfo.title}" (ID: ${pageId}) at ${position}`);
      
      const icons = { note: '📝', todo: '✅', important: '🚨', question: '❓' };
      const colors = { note: '#e3f2fd', todo: '#e8f5e8', important: '#ffebee', question: '#fff3e0' };
      const noteHtml = `
        <div style="border-left: 4px solid #2196f3; background-color: ${colors[noteType]}; padding: 10px; margin: 10px 0;">
          <p><strong>${icons[noteType]} ${noteType.charAt(0).toUpperCase() + noteType.slice(1)}</strong> - <em>${new Date().toLocaleString()}</em></p>
          <p>${textToHtml(note)}</p>
        </div>`;
      
      const action = position === 'top' ? 'prepend' : 'append';
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: action, content: noteHtml }]),
        signal: AbortSignal.timeout(45000)
      });
      
      if (!response.ok) throw new Error(`Add note failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **${noteType.charAt(0).toUpperCase() + noteType.slice(1)} Added!**\nPage: ${pageInfo.title}\nPosition: ${position}.` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to add note to page ID "${pageId}": ${error.message}` }] };
    }
  }
);

server.tool(
  'addTableToPage',
  {
    pageId: z.string().describe('The ID of the page to add a table to.'),
    tableData: z.string().describe('Table data in CSV format (header row, then data rows).'),
    title: z.string().describe('Optional title for the table.').optional(),
    position: z.enum(['top', 'bottom'])
      .default('bottom')
      .describe('Position to add the table (top or bottom).')
      .optional()
  },
  async ({ pageId, tableData, title, position }) => {
    try {
      await ensureGraphClient();
      const pageInfo = await graphClient.api(`/me/onenote/pages/${pageId}`).get();
      console.error(`Adding table to page: "${pageInfo.title}" (ID: ${pageId}) at ${position}`);
      
      const rows = tableData.trim().split('\n').map(row => row.split(',').map(cell => cell.trim()));
      if (rows.length < 2) throw new Error('Table data must have at least a header row and one data row.');
      
      const headerRow = rows[0];
      const dataRows = rows.slice(1);
      let tableHtml = title ? `<h3>📊 ${textToHtml(title)}</h3>` : '';
      tableHtml += `<table style="border-collapse: collapse; width: 100%; margin: 10px 0;"><thead><tr style="background-color: #f5f5f5;">${headerRow.map(cell => `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${textToHtml(cell)}</th>`).join('')}</tr></thead><tbody>${dataRows.map(row => `<tr>${row.map(cell => `<td style="border: 1px solid #ddd; padding: 8px;">${textToHtml(cell)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
      
      const action = position === 'top' ? 'prepend' : 'append';
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        method: 'PATCH',
        headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify([{ target: 'body', action: action, content: tableHtml }]),
        signal: AbortSignal.timeout(45000)
      });
      
      if (!response.ok) throw new Error(`Add table failed: ${response.status} ${response.statusText}`);
      
      return { content: [{ type: 'text', text: `✅ **Table Added!**\nPage: ${pageInfo.title}\nTitle: ${title || 'Untitled'}\nPosition: ${position}.` }] };
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: `❌ Failed to add table to page ID "${pageId}": ${error.message}` }] };
    }
  }
);

// --- Page Creation Tool ---
server.tool(
  'createPage',
  {
    title: z.string().min(1, { message: "Title cannot be empty." }).describe('The title for the new page.'),
    content: z.string().min(1, { message: "Content cannot be empty." }).describe('The content for the new page (HTML or markdown-style).')
  },
  async ({ title, content }) => {
    try {
      await ensureGraphClient();
      console.error(`Attempting to create page with title: "${title}"`);
      
      const sectionsResponse = await graphClient.api('/me/onenote/sections').get();
      if (!sectionsResponse.value || sectionsResponse.value.length === 0) {
        throw new Error('No sections found in your OneNote. Cannot create a page.');
      }
      const targetSectionId = sectionsResponse.value[0].id;
      const targetSectionName = sectionsResponse.value[0].displayName;
      
      const htmlContent = textToHtml(content);
      const pageHtml = `<!DOCTYPE html>
<html>
<head>
  <title>${textToHtml(title)}</title>
  <meta charset="utf-8">
</head>
<body>
  <h1>${textToHtml(title)}</h1>
  ${htmlContent}
  <hr>
  <p><em>Created via OneNote MCP on ${new Date().toLocaleString()}</em></p>
</body>
</html>`;
      
      const response = await graphClient
        .api(`/me/onenote/sections/${targetSectionId}/pages`)
        .header('Content-Type', 'application/xhtml+xml')
        .post(pageHtml);
      
      const webUrl = response.links?.oneNoteWebUrl?.href || '';
      
      return {
        content: [{
          type: 'text',
          text: `✅ **Page Created Successfully!**
**Title:** ${response.title}
**Page ID:** ${response.id}
**In Section:** ${targetSectionName}
**Created:** ${new Date(response.createdDateTime).toLocaleString()}${webUrl ? `\n\n🔗 <${webUrl}>` : ''}`
        }]
      };
    } catch (error) {
      console.error(`CREATE PAGE ERROR: ${error.message}`, error.stack);
      return { isError: true, content: [{ type: 'text', text: `❌ **Error creating page:** ${error.message}` }] };
    }
  }
);



// ============================================================================
// SERVER STARTUP
// ============================================================================

/**
 * Main function to initialize and start the MCP server.
 */
async function main() {
  loadExistingToken(); // Attempt to load token at startup
  if (accessToken) {
    initializeGraphClient(); // Initialize client if token was loaded
    
    // Load notebook cache from disk if available
    const cacheLoaded = loadNotebookCacheFromDisk();
    if (cacheLoaded) {
      console.error('📦 Notebook cache loaded from disk');
    } else {
      console.error('💾 No valid cache found, will load on first use');
    }
  }

  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    console.error('🚀✨ OneNote Ultimate MCP Server is now LIVE! ✨🚀');
    console.error(`   Client ID: ${clientId.substring(0, 8)}... (Using ${process.env.AZURE_CLIENT_ID ? 'environment variable' : 'default'})`);
    console.error('   Ready to manage your OneNote like never before!');
    console.error('--- Available Tool Categories ---');
    console.error('   🔐 Auth: authenticate, saveAccessToken, checkTokenScopes');
    console.error('   📚 Read: listNotebooks, listSections, listPagesInSection, searchPages, searchPagesByDate, searchPageContent, searchInNotebook, getPageContent, getPageByTitle');
    console.error('   📝 Productivity: getMyRecentChanges, createDailyNote');
    console.error('   ✏️ Edit: updatePageContent, appendToPage, updatePageTitle, replaceTextInPage, addNoteToPage, addTableToPage');
    console.error('   ➕ Create: createPage');
    console.error('---------------------------------');
    
    process.on('SIGINT', () => {
      console.error('\n🔌 OneNote MCP Server shutting down gracefully...');
      process.exit(0);
    });
    process.on('SIGTERM', () => {
      console.error('\n🔌 OneNote MCP Server terminated...');
      process.exit(0);
    });

  } catch (error) {
    console.error(`💀 Critical error starting server: ${error.message}`, error.stack);
    process.exit(1);
  }
}

main();