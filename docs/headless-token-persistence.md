# OneNote Token Cache Retention Investigation — 2026-08-19

Subagent investigation: why the OneNote access-token/refresh-token cache was "not retaining" across
OpenClaw gateway restarts, with reference to the non-atomic-write / corruption context.

## Verdict
Two sequential root causes; both addressed on 2026-08-19. Retention now VERIFIED WORKING (live check
23:2x MDT: `listNotebooks` → 9 notebooks via gateway MCP after the 22:50:56 gateway restart).

## Root cause #1 — refresh token was never persisted at all (pre 12:28)
- Original `onenote-mcp.mjs` used keyring-backed `@azure/identity-cache-persistence`.
- Headless box: no DBus/libsecret → keyring init crashed the MCP child ~2s after first tool call.
- Persistence disabled → only `.access-token.txt` (1h AT) hit disk. RT lived in MCP process memory only.
- Gateway restart → fresh MCP process → empty cache → `trySilentToken()` fails → device-code re-auth.
- Fix: commit `ead4031` — file-backed MSAL cache (`@azure/msal-node-extensions`
  `FilePersistence` + `PersistenceCachePlugin` via custom `useIdentityPlugin`), cache →
  `.msal-cache-onenote-mcp-token-cache.json` (0600), `tokenCachePersistenceOptions.enabled: true`.
- Proof: 12:28 device-code login; fresh process with `.access-token.txt` moved aside silently
  renewed from the persisted cache.

## Root cause #2 — non-atomic saves corrupted the file on restart / concurrent runs (13:58–21:1x)
- Confirmed in `node_modules/@azure/msal-node-extensions/lib/msal-node-extensions.cjs` L1158–1180:
  `FilePersistence.save()` = chmod + plain `fs.promises.writeFile(path, contents)` + chmod.
  `writeFile` opens with O_TRUNC and is **not atomic**, no fsync, no cross-call serialization at the
  file level.
- Kill window: the MCP SIGTERM handler (`onenote-mcp.mjs` ~L2632) does `process.exit(0)` immediately —
  no flush/grace. Gateway restart SIGTERMs MCP children; a kill landing mid-save tears the file.
  Evidence: cache found at **0 bytes** before the 21:11 re-auth (truncate done, write killed).
- Interleave window: two corruption events produced a **1504-char binary blob** where the JWT should
  be. Both cache versions are near-identical JSON, so byte interleaving (or partial write in the
  truncated file) yields *parseable* JSON with mangled token values. First event: gateway had spawned
  **two** MCP instances (PIDs 27390/27435, same parent) and load-balanced tool calls between them
  (16:44 session). Second event (17:04) was nominally single-writer — consistent with kill-mid-write
  on the already-truncated file.
- `PersistenceCachePlugin` does have a cross-process lockfile (`.lockfile`) serializing
  before/afterCacheAccess, but it does not (a) stop per-instance device-code state divergence
  (code completed in instance A, next call hit unauthenticated instance B), nor (b) a kill inside
  the lock section before rename semantics existed.

## Fix #2 (commit 50a7dc3, 18:53, loaded at the 19:4x gateway restart) — hardened save
`onenote-mcp.mjs` L39–108:
1. `msalAtomicWrite()`: write to `<file>.<pid>.tmp` + `rename` → atomic on POSIX (same fs); a killed
   save leaves the previous file intact.
2. Validated `.lastgood` snapshot: only written when the just-saved file passes validation.
3. `msalEnsureCacheHealthy()`: auto-restore from `.lastgood` at process startup AND inside the
   persistence factory (before first silent renewal).

## Live state verified (23:2x MDT)
- Cache file: 5999 B, 0600, valid JSON, all 5 MSAL sections present (Account/IdToken/AccessToken/
  **RefreshToken** 1 entry/AppMetadata); `.lastgood` byte-identical (`cmp` clean); no stray tmp/lock
  files.
- Gateway restarted 22:50:56 (PID 35057); cache + `.access-token.txt` updated 22:54 (silent renewal
  succeeded after restart).
- All MCP instances started after 22:56 → all running the hardened code (file mtime 18:52).
- `onenote__listNotebooks` → 9 notebooks ✅.

## Residual risks / recommended follow-ups (priority order)
1. **Multi-writer condition is LIVE right now**: gateway PID 35057 has ≥3 concurrent OneNote MCP
   children (35295@22:56, 35374@22:59, 36687@23:19; 36102@23:10 exited). Atomic write + lockfile
   prevent byte tears, but each instance holds its own in-memory MSAL cache:
   - device-code re-auth confusion (code lands in one instance; others keep demanding fresh codes);
   - RT-rotation last-writer-wins: a stale instance can save an older RT. Personal Microsoft
     accounts tolerate this (rolling RT grace ~90d); org accounts with RT retirement would brick.
   - Proper fix is gateway-side: one shared MCP instance across sessions (config:
     `~/.openclaw/openclaw.json` → `mcp.servers.onenote`). Until then keep the TOOLS.md rule: never
     run a second MCP instance/auth driver against the same dir.
2. **Validation too shallow**: `msalValidateCacheFile` only does `JSON.parse`. The observed binary
   blob was parseable JSON with a mangled value → it would pass validation and be snapshotted to
   `.lastgood` (self-heal restores the corruption). Strengthen: require
   `data.RefreshToken && Object.keys(data.RefreshToken).length > 0`, ideally also that AT values look
   like JWTs (3 dot-separated segments).
3. **No fsync before rename**: safe for process kills (what restarts are), not for power loss/kernel
   panic. Use fd-level open/write/fsync/close + dir fsync if that matters.
4. **SIGTERM exits without draining in-flight saves**: benign now (worst case a seconds-stale RT on
   disk; rolling RT grace covers it). Optional: drain pending saves ≤2s before `process.exit(0)`.
5. **Orphan window**: the 22:54 cache write postdates the 22:50:56 new-gateway start → an MCP child
   of the dying gateway was still alive at 22:54. Harmless with atomic writes, but explains
   "stale in-memory auth state" re-auth confusion (21:11 session): orphaned children can outlive a
   gateway restart.

## Key evidence locations
- MCP code: `/home/mintvm/.openclaw/mcp-servers/onenote-mcp/onenote-mcp.mjs` (hardened block L39–108;
  SIGTERM handler ~L2632)
- Non-atomic save: `.../node_modules/@azure/msal-node-extensions/lib/msal-node-extensions.cjs`
  L1158–1180 (`FilePersistence.save`); L271+ (`PersistenceCachePlugin`, lockfile + save path)
- Commits: `ead4031` (file-backed cache), `50a7dc3` (hardening)
- Corrupt quarantines: `.corrupt-auth-backup-20260819{,b,c}/` in the MCP dir
- Session notes: `memory/2026-08-19.md`, `2026-08-19-1242.md`, `-1616.md`, `-1644.md`, `-1941.md`,
  `-2111.md`; rules in `TOOLS.md`
