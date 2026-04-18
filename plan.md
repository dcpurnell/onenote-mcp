# OneNote MCP Server — `getPageContent` Timeout Fix

## Your Role

You are a coding agent working on the source code of a **Model Context Protocol (MCP) server** that wraps the Microsoft Graph OneNote API. A downstream client (Claude) uses this server to fetch OneNote page content for a weekly review pipeline. Two of the server's tools consistently time out. Your job is to diagnose the root cause and ship a fix.

Work diagnostic-first. Do not make code changes until you have a concrete hypothesis confirmed by a reproduction or by reading the relevant handler code.

---

## Symptom

Two tools time out at 60 seconds on **every** call, across many different pages:

| Tool | Behavior | Result |
|------|----------|--------|
| `getPageContent(pageId, format)` | All formats (`text`, `html`, `summary`) | ❌ 60s timeout |
| `getPageByTitle(title, format)` | All formats | ❌ 60s timeout |

Meanwhile these tools work normally and return in under ~2 seconds:

| Tool | Returns | Result |
|------|---------|--------|
| `listNotebooks(refresh: true, includeTeamNotebooks: true)` | 69 notebooks | ✅ |
| `searchPagesByDate(days: 7)` | 22 pages with titles, URLs, dates | ✅ |
| `searchPageContent(query: "SQL", days: 1, maxPages: 3)` | 2 matches with **snippet excerpts** | ✅ |
| `checkTokenScopes()` | Full scope list | ✅ |

Auth is fine. Token is fresh, has `Notes.Read`, `Notes.Read.All`, `Notes.ReadWrite`, `Notes.ReadWrite.All`. The failure is specific to the full-page content retrieval endpoints.

Key tell: `searchPageContent` returns **snippets** (short excerpts) quickly. `getPageContent` returns **full page HTML** and hangs. This points at payload size or a blocking pattern in the full-content code path — not at auth, not at Graph availability, not at page-specific issues.

---

## Hypothesis

The `getPageContent` and `getPageByTitle` handlers almost certainly call:

```
GET https://graph.microsoft.com/v1.0/me/onenote/pages/{id}/content
```

This endpoint returns the full page as HTML. Likely failure modes, in rough order of probability:

1. **HTTP client default timeout is below 60s** — the request is actually in flight when the MCP server's outer 60s timeout fires and kills it. The Graph call might succeed on its own given enough time.
2. **Response body is large and being buffered synchronously** — pages with images, tables, or long meeting notes produce big HTML. Buffering to a string before returning blocks the event loop or the async task.
3. **Retry logic is amplifying the latency** — a 429 throttle on first attempt triggers backoff + retry, and the retries stack up inside the 60s window.
4. **`getPageByTitle` is doing a title search + full-content fetch in sequence** — if the first step is already slow and the second adds more, total time blows past the timeout.
5. **Token refresh inside the handler** — if auth refresh is triggered lazily inside the content call (but not the search calls), that adds seconds.

The last three are less likely given that `searchPageContent` uses the same auth and works, so rule those out only after checking timing.

---

## Investigation Steps

Do these in order. Stop once you have enough evidence to pick a fix.

### Step 1: Locate the handlers

Find the tool implementations:

- `getPageContent`
- `getPageByTitle`
- `searchPageContent` (this one works — useful as a control)
- `searchPagesByDate` (works — also a control)

Search the repo for the tool names, or for `onenote/pages` URL fragments. Note the HTTP client being used (axios, fetch, got, node-fetch, python requests, etc.) and any shared request wrapper.

### Step 2: Read the working vs failing code side by side

Compare `searchPageContent` to `getPageContent`:

- Different Graph endpoint? (`/search` vs `/pages/{id}/content`)
- Different response handling? (JSON snippet list vs raw HTML body)
- Different timeout config?
- Different error handling / retry behavior?
- Different auth helper?

Write down the deltas. The root cause lives in one of them.

### Step 3: Check the HTTP client timeout configuration

Find where the HTTP client is instantiated. Look for `timeout`, `requestTimeout`, `httpAgent`, or similar. If it's set below 60s, that's likely the problem — the Graph response is larger and slower than the client allows, and it aborts before the outer MCP timeout would.

Note: the MCP protocol has its own 60s request timeout on the **client** side (Claude). If the MCP server takes longer than 60s to respond, the client gives up. So even if the server's HTTP client has no timeout, a too-slow Graph call still causes a client-visible timeout. **This means the fix must make the handler respond in well under 60s**, not just eventually succeed.

### Step 4: Instrument one failing call

Add temporary logging to `getPageContent`:

```
- timestamp before Graph request
- timestamp after Graph response headers arrive
- response Content-Length header if present
- timestamp after body fully read
- any retry attempts and their delays
- total handler duration
```

Run the tool against a small page (a fresh daily shutdown page is a good test — they are < 2KB). If this still times out at 60s on a tiny page, the issue is not payload size — it is auth refresh, throttling, or a bug in the handler itself.

If the small page succeeds but larger pages fail, the issue is payload-related and fixes 3 or 4 from the "Likely Fixes" list below are appropriate.

### Step 5: Sanity check Graph directly

Use Postman / Graph Explorer / curl with a fresh token to hit:

```
GET https://graph.microsoft.com/v1.0/me/onenote/pages/{pageId}/content
Authorization: Bearer <token>
```

Time the raw call. If Graph itself takes > 30s, the problem is upstream (Graph throttling, tenant load, etc.) and the MCP server can only mitigate — not fix.

---

## Likely Fixes (in order of simplicity)

Pick the smallest fix that addresses what you find. Do not apply them all — pick one or two.

### Fix A — Raise the HTTP client timeout

If the client has a timeout below 60s, raise it to ~45s (leaving headroom under the MCP 60s ceiling). This is one line of config.

### Fix B — Strip heavy elements before returning

OneNote HTML includes embedded image data URLs and resource references that can balloon page size. For the `text` and `summary` formats, strip these before building the response. Keep `html` format faithful if needed, but consider stripping `<img>` `src` attributes (images aren't useful for the downstream summarization anyway).

A second option: when `format=summary` is requested, truncate to the first ~2000 characters of visible text. `summary` is meant to be brief.

### Fix C — Stream the response body

If the HTTP client buffers the full body into a string before the handler returns, switch to a streaming read and parse incrementally. Most of the visible text comes from the first few KB of HTML in practice.

### Fix D — Graceful degradation on timeout

Wrap the Graph call in a shorter internal timeout (say, 45s) and return a structured error response instead of hanging until the MCP timeout fires. Callers can handle a fast error better than a slow timeout.

### Fix E — Honor `Retry-After` but cap total latency

If throttling retries are adding up, ensure each retry's delay is bounded and total retries × delay < 40s. If throttling persists, return the throttle error to the caller rather than retrying indefinitely.

---

## Success Criteria

After the fix:

- [ ] `getPageContent(pageId, format: "text")` returns in < 30s for pages up to ~100KB
- [ ] `getPageContent(pageId, format: "summary")` returns in < 10s for any page
- [ ] `getPageByTitle(title, format: "summary")` returns in < 15s
- [ ] 20 consecutive `getPageContent` calls on different pages all succeed
- [ ] No regression in `searchPagesByDate`, `searchPageContent`, `listNotebooks`, `listPagesInSection`
- [ ] Existing tool schemas are unchanged (downstream client depends on them)

---

## Test Procedure

1. Start the MCP server locally with logging enabled.
2. Use the MCP inspector (or a small script that speaks MCP stdio) to call:
   - `getPageContent` on a short page (e.g., a one-paragraph daily shutdown)
   - `getPageContent` on a long page (e.g., a long-running stand-up page)
   - `getPageByTitle` with a unique title
   - `searchPageContent` (regression check — should still work)
   - `searchPagesByDate` (regression check)
3. Record timing for each. Compare before/after.
4. Run all four calls back-to-back to confirm no cumulative slowdown.

Test page IDs can be obtained by calling `searchPagesByDate(days: 7, includeTeamNotebooks: true)` and inspecting the returned metadata. If the tool doesn't currently surface IDs in its output, fix that too — downstream callers need IDs to chain into `getPageContent`.

---

## Out of Scope

Don't take on the following in this PR:

- Refactoring auth / token storage
- Changing the tool schemas (adding, removing, renaming tools or parameters)
- Search or list tool behavior (those work)
- Device-code flow or authentication UX
- Rebuilding the MCP transport layer

Keep the diff surgical. One or two handlers, one timeout/streaming change, logging to verify.

---

## Deliverables

1. **Root cause summary** (1–2 paragraphs): what was actually causing the timeout
2. **Code changes**: a PR or diff with clear commit messages
3. **Before/after timing table**: the same calls, timed on the same pages, before the change and after
4. **Any new config**: env vars, defaults changed, documented in the README
5. **Notes on any secondary issues spotted** but not fixed — for follow-up

---

## Context the downstream client needs (reference)

The downstream Claude-based client:

- Calls `searchPagesByDate(days: 7)` first to get a list of recently touched pages
- Then expects to call `getPageContent(pageId)` on each to build summaries
- Currently falls back to a metadata-only feed because `getPageContent` times out
- Does not currently receive page IDs from `searchPagesByDate` — the tool returns titles and SharePoint URLs but not Graph API page IDs. If `getPageContent` requires an ID and the search results don't include one, that is also a bug worth confirming and fixing.

If you find that `searchPagesByDate` suppresses page IDs in its response, add them to the output. That's a one-line fix and unblocks the downstream use case even before the timeout fix lands.
