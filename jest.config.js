export default {
  testEnvironment: 'node',
  moduleNameMapper: {
    '^(\\.{1,2}/.*)\\.js$': '$1',
  },
  transform: {},
  collectCoverageFrom: [
    'onenote-mcp.mjs',
    '!**/node_modules/**',
    '!**/tests/**',
  ],
  coverageThreshold: {
    global: {
      branches: 70,
      functions: 70,
      lines: 80,
      statements: 80,
    },
  },
  testMatch: [
    '**/tests/**/*.test.mjs',
    '**/tests/**/*.test.js',
  ],
  verbose: true,
  testTimeout: 10000,
};
