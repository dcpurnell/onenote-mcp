/**
 * Test fixtures for HTML content used in testing HTML processing functions
 */

export const simpleHTML = `
<!DOCTYPE html>
<html>
<head><title>Test Page</title></head>
<body>
  <h1>Welcome to OneNote</h1>
  <p>This is a simple paragraph.</p>
</body>
</html>
`;

export const complexHTML = `
<!DOCTYPE html>
<html>
<body>
  <h1>Main Heading</h1>
  <h2>Subheading</h2>
  <p>This is a paragraph with <strong>bold</strong> and <em>italic</em> text.</p>
  <ul>
    <li>List item 1</li>
    <li>List item 2</li>
    <li>List item 3</li>
  </ul>
  <ol>
    <li>Ordered item 1</li>
    <li>Ordered item 2</li>
  </ol>
  <table>
    <tr>
      <th>Column 1</th>
      <th>Column 2</th>
    </tr>
    <tr>
      <td>Data 1</td>
      <td>Data 2</td>
    </tr>
  </table>
</body>
</html>
`;

export const malformedHTML = `
<div>
  <p>Unclosed paragraph
  <strong>Unclosed strong tag
  <ul>
    <li>Item 1
    <li>Item 2
  </div>
`;

export const scriptHTML = `
<!DOCTYPE html>
<html>
<body>
  <h1>Test XSS</h1>
  <script>alert('XSS');</script>
  <p onclick="alert('click')">Click me</p>
  <a href="javascript:alert('XSS')">Link</a>
  <img src="x" onerror="alert('error')">
</body>
</html>
`;

export const emptyHTML = ``;

export const whitespaceHTML = `
   
   
   
`;

export const nestedHTML = `
<!DOCTYPE html>
<html>
<body>
  <div>
    <div>
      <div>
        <h1>Deeply Nested</h1>
        <p>This content is nested several levels deep.</p>
        <ul>
          <li>
            <ul>
              <li>Nested list item</li>
            </ul>
          </li>
        </ul>
      </div>
    </div>
  </div>
</body>
</html>
`;

export const unicodeHTML = `
<!DOCTYPE html>
<html>
<body>
  <h1>Unicode Test 🎉</h1>
  <p>Chinese: 你好世界</p>
  <p>Arabic: مرحبا بالعالم</p>
  <p>Emoji: 😀 🎨 🚀</p>
  <p>Special: © ® ™ € ¥</p>
</body>
</html>
`;

export const largeHTML = `
<!DOCTYPE html>
<html>
<body>
  ${Array.from({ length: 100 }, (_, i) => `
    <h2>Section ${i + 1}</h2>
    <p>This is paragraph ${i + 1} with some content to make it longer and test performance.</p>
    <ul>
      <li>Item 1</li>
      <li>Item 2</li>
      <li>Item 3</li>
    </ul>
  `).join('\n')}
</body>
</html>
`;
