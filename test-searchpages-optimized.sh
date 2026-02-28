#!/bin/bash

# Test script for optimized searchPages tool
# Run after authenticating with the MCP server

echo "=== Testing Optimized searchPages Tool ==="
echo ""

echo "Test 1: Basic search with query"
node test-mcp-tool.mjs searchPages '{"query": "meeting"}'
echo ""

echo "Test 2: Search with top limit and orderBy"
node test-mcp-tool.mjs searchPages '{"top": 20, "orderBy": "modified", "maxResults": 10}'
echo ""

echo "Test 3: Search specific notebook by name"
node test-mcp-tool.mjs searchPages '{"notebookName": "Work", "query": "project"}'
echo ""

echo "Test 4: Alphabetical search"
node test-mcp-tool.mjs searchPages '{"orderBy": "title", "maxResults": 20}'
echo ""

echo "Test 5: Oldest pages first"
node test-mcp-tool.mjs searchPages '{"orderBy": "created", "maxResults": 15}'
echo ""

echo "=== Tests Complete ==="
