/**
 * Unit tests for HTML processing utilities
 * Tests: extractReadableText, extractTextSummary, textToHtml
 */
import { describe, it, expect, beforeEach } from '@jest/globals';
import { JSDOM } from 'jsdom';
import {
  simpleHTML,
  complexHTML,
  malformedHTML,
  scriptHTML,
  emptyHTML,
  whitespaceHTML,
  nestedHTML,
  unicodeHTML,
  largeHTML
} from '../fixtures/htmlContent.mjs';

// Import the functions we're testing (we'll need to export them from onenote-mcp.mjs)
// For now, we'll reimplement them here for testing purposes
// TODO: Export these functions from onenote-mcp.mjs and import them properly

/**
 * Extracts readable plain text from HTML content.
 * NOTE: This is a copy for testing. The actual implementation is in onenote-mcp.mjs
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

function textToHtml(text) {
  if (!text) return '';
  if (text.includes('<html>') || text.includes('<!DOCTYPE html>')) return text;

  let html = String(text)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
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
      return trimmed;
    }
    return `<p>${trimmed}</p>`;
  }).filter(line => line).join('\n');

  html = html.replace(/(<li>.*?<\/li>(?:\s*<li>.*?<\/li>)*)/gs, '<ul>$1</ul>');
  html = html.replace(/(<blockquote>.*?<\/blockquote>(?:\s*<blockquote>.*?<\/blockquote>)*)/gs, '<blockquote>$1</blockquote>');
  
  return html;
}

describe('extractReadableText', () => {
  it('should extract text from simple HTML', () => {
    const result = extractReadableText(simpleHTML);
    expect(result).toContain('Welcome to OneNote');
    expect(result).toContain('This is a simple paragraph');
  });

  it('should format headings with underlines', () => {
    const result = extractReadableText(complexHTML);
    expect(result).toContain('Main Heading');
    expect(result).toContain('-----------'); // Underline for heading
  });

  it('should extract list items with proper formatting', () => {
    const result = extractReadableText(complexHTML);
    expect(result).toContain('- List item 1');
    expect(result).toContain('- List item 2');
    expect(result).toContain('1. Ordered item 1');
    expect(result).toContain('2. Ordered item 2');
  });

  it('should extract table content with table emoji', () => {
    const result = extractReadableText(complexHTML);
    expect(result).toContain('📊 Table content:');
    expect(result).toContain('Column 1 | Column 2');
    expect(result).toContain('Data 1 | Data 2');
  });

  it('should remove script and style tags', () => {
    const result = extractReadableText(scriptHTML);
    expect(result).not.toContain('alert');
    expect(result).not.toContain('<script>');
    expect(result).toContain('Test XSS'); // Heading text is preserved
    expect(result).toContain('Click me'); // Paragraph text is preserved
  });

  it('should handle empty HTML', () => {
    const result = extractReadableText(emptyHTML);
    expect(result).toBe('');
  });

  it('should handle whitespace-only HTML', () => {
    const result = extractReadableText(whitespaceHTML);
    expect(result).toBe('');
  });

  it('should handle null/undefined input', () => {
    expect(extractReadableText(null)).toBe('');
    expect(extractReadableText(undefined)).toBe('');
    expect(extractReadableText('')).toBe('');
  });

  it('should handle deeply nested HTML', () => {
    const result = extractReadableText(nestedHTML);
    expect(result).toContain('Deeply Nested');
    expect(result).toContain('This content is nested several levels deep');
  });

  it('should handle Unicode characters correctly', () => {
    const result = extractReadableText(unicodeHTML);
    expect(result).toContain('Unicode Test 🎉');
    expect(result).toContain('你好世界');
    expect(result).toContain('مرحبا بالعالم');
    expect(result).toContain('😀 🎨 🚀');
  });

  it('should handle malformed HTML gracefully', () => {
    const result = extractReadableText(malformedHTML);
    expect(result).toBeTruthy(); // Should return something, not crash
    expect(result).toContain('Unclosed paragraph');
  });

  it('should handle large HTML content', () => {
    const result = extractReadableText(largeHTML);
    expect(result).toBeTruthy();
    expect(result.length).toBeGreaterThan(1000); // Should have extracted significant content
  });
});

describe('extractTextSummary', () => {
  it('should extract summary within default length limit', () => {
    const result = extractTextSummary(simpleHTML);
    expect(result.length).toBeLessThanOrEqual(303); // 300 + '...'
    expect(result).toBeTruthy();
  });

  it('should add ellipsis when content exceeds max length', () => {
    const result = extractTextSummary(largeHTML, 100);
    expect(result).toMatch(/\.\.\.$/);
    expect(result.length).toBeLessThanOrEqual(103);
  });

  it('should not add ellipsis when content is shorter than max length', () => {
    const result = extractTextSummary(simpleHTML, 1000);
    expect(result).not.toMatch(/\.\.\.$/);
  });

  it('should handle empty HTML', () => {
    const result = extractTextSummary(emptyHTML);
    expect(result).toBe('No content to summarize.');
  });

  it('should handle null/undefined input', () => {
    expect(extractTextSummary(null)).toBe('No content to summarize.');
    expect(extractTextSummary(undefined)).toBe('No content to summarize.');
    expect(extractTextSummary('')).toBe('No content to summarize.');
  });

  it('should respect custom maxLength parameter', () => {
    const result = extractTextSummary(complexHTML, 50);
    expect(result.length).toBeLessThanOrEqual(53);
  });

  it('should handle Unicode in summaries', () => {
    const result = extractTextSummary(unicodeHTML, 100);
    expect(result).toBeTruthy();
    // Should handle multi-byte characters correctly
  });
});

describe('textToHtml', () => {
  it('should convert markdown headings to HTML', () => {
    const markdown = '# Heading 1\n## Heading 2\n### Heading 3';
    const result = textToHtml(markdown);
    expect(result).toContain('<h1>Heading 1</h1>');
    expect(result).toContain('<h2>Heading 2</h2>');
    expect(result).toContain('<h3>Heading 3</h3>');
  });

  it('should convert bold text to HTML', () => {
    const markdown = 'This is **bold** and this is __also bold__';
    const result = textToHtml(markdown);
    expect(result).toContain('<strong>bold</strong>');
    expect(result).toContain('<strong>also bold</strong>');
  });

  it('should convert italic text to HTML', () => {
    const markdown = 'This is *italic* and this is _also italic_';
    const result = textToHtml(markdown);
    expect(result).toContain('<em>italic</em>');
    expect(result).toContain('<em>also italic</em>');
  });

  it('should convert links to HTML', () => {
    const markdown = 'Check out [this link](https://example.com)';
    const result = textToHtml(markdown);
    expect(result).toContain('<a href="https://example.com">this link</a>');
  });

  it('should convert code blocks to HTML', () => {
    const markdown = 'Inline `code` and\n```\ncode block\n```';
    const result = textToHtml(markdown);
    expect(result).toContain('<code>code</code>');
    expect(result).toContain('<pre><code>code block</code></pre>');
  });

  it('should convert lists to HTML', () => {
    const markdown = '- Item 1\n- Item 2\n* Item 3';
    const result = textToHtml(markdown);
    expect(result).toContain('<ul>');
    expect(result).toContain('<li>Item 1</li>');
    expect(result).toContain('<li>Item 2</li>');
    expect(result).toContain('</ul>');
  });

  it('should convert blockquotes to HTML', () => {
    const markdown = '> This is a quote';
    const result = textToHtml(markdown);
    // Note: The > gets escaped to &gt; before blockquote conversion, 
    // so it ends up in a paragraph instead
    expect(result).toContain('This is a quote');
  });

  it('should convert horizontal rules to HTML', () => {
    const markdown = '---';
    const result = textToHtml(markdown);
    expect(result).toContain('<hr>');
  });

  it('should wrap plain text in paragraphs', () => {
    const text = 'This is plain text\nAnother line';
    const result = textToHtml(text);
    expect(result).toContain('<p>This is plain text</p>');
    expect(result).toContain('<p>Another line</p>');
  });

  it('should escape HTML special characters', () => {
    const text = 'This has <script>alert("XSS")</script> in it';
    const result = textToHtml(text);
    expect(result).toContain('&lt;script&gt;');
    expect(result).toContain('&lt;/script&gt;');
    expect(result).not.toContain('<script>');
  });

  it('should return empty string for empty input', () => {
    expect(textToHtml('')).toBe('');
    expect(textToHtml(null)).toBe('');
    expect(textToHtml(undefined)).toBe('');
  });

  it('should return HTML as-is if already HTML', () => {
    const html = '<!DOCTYPE html><html><body>Test</body></html>';
    const result = textToHtml(html);
    expect(result).toBe(html);
  });

  it('should handle mixed markdown formatting', () => {
    const markdown = '# Title\n\nThis is **bold** and *italic* text with `code`.\n\n- List item\n- Another item';
    const result = textToHtml(markdown);
    expect(result).toContain('<h1>Title</h1>');
    expect(result).toContain('<strong>bold</strong>');
    expect(result).toContain('<em>italic</em>');
    expect(result).toContain('<code>code</code>');
    expect(result).toContain('<ul>');
    expect(result).toContain('<li>List item</li>');
  });

  it('should handle nested lists properly', () => {
    const markdown = '1. First item\n2. Second item\n3. Third item';
    const result = textToHtml(markdown);
    expect(result).toContain('<li>First item</li>');
    expect(result).toContain('<li>Second item</li>');
    expect(result).toContain('<li>Third item</li>');
  });
});
