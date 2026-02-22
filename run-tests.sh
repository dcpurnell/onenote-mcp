#!/bin/bash

# Quick Test Commands for OneNote MCP Server
# Run these commands whenever you make changes to verify everything works

echo "🧪 OneNote MCP Server - Test Suite"
echo "===================================="
echo ""

# 1. Run unit tests (fastest, always works)
echo "1️⃣  Running Unit Tests..."
npm run test:unit
UNIT_EXIT=$?

echo ""
echo "=================================="
echo ""

# 2. Run authentication integration tests
echo "2️⃣  Running Authentication Tests..."
npm run test:integration -- authentication.test.mjs
AUTH_EXIT=$?

echo ""
echo "=================================="
echo ""

# 3. Show summary
echo "📊 Test Summary:"
echo "=================================="
if [ $UNIT_EXIT -eq 0 ]; then
  echo "✅ Unit Tests: PASSED (33 tests)"
else
  echo "❌ Unit Tests: FAILED"
fi

if [ $AUTH_EXIT -eq 0 ]; then
  echo "✅ Authentication Tests: PASSED (20 tests)"
else
  echo "❌ Authentication Tests: FAILED"
fi

echo ""
echo "💡 Quick Commands:"
echo "  npm run test:unit       - Run unit tests only"
echo "  npm run test:watch      - Watch mode (auto-run on changes)"
echo "  npm run test:coverage   - Generate coverage report"
echo "  npm test                - Run all tests"
echo ""

# Exit with error if any tests failed
if [ $UNIT_EXIT -ne 0 ] || [ $AUTH_EXIT -ne 0 ]; then
  exit 1
fi

echo "✨ All tests passed! Safe to commit."
exit 0
