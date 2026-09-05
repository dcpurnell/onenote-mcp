// One-off interactive device-code login for the OneNote MCP server.
// Run:  node ~/tools/onenote-mcp/auth-cli.mjs
// Shares client ID, scopes, cache name and auth-record path with onenote-mcp.mjs,
// so the MCP server can then acquire tokens silently.
import { DeviceCodeCredential, useIdentityPlugin, serializeAuthenticationRecord } from '@azure/identity';
import { cachePersistencePlugin } from '@azure/identity-cache-persistence';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const authRecordFilePath = path.join(__dirname, '.auth-record.json');
const tokenFilePath = path.join(__dirname, '.access-token.txt');
const clientId = process.env.AZURE_CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const scopes = ['Notes.Read', 'Notes.ReadWrite', 'Notes.Read.All', 'Notes.ReadWrite.All', 'Notes.Create', 'User.Read'];

useIdentityPlugin(cachePersistencePlugin);

const credential = new DeviceCodeCredential({
  clientId,
  tenantId: process.env.AZURE_TENANT_ID || 'organizations',
  tokenCachePersistenceOptions: { enabled: true, name: 'onenote-mcp-token-cache', unsafeAllowUnencryptedStorage: true },
  userPromptCallback: (info) => console.log(`\n${info.message}\n`),
});

const record = await credential.authenticate(scopes);
fs.writeFileSync(authRecordFilePath, serializeAuthenticationRecord(record));
const tok = await credential.getToken(scopes);
fs.writeFileSync(tokenFilePath, JSON.stringify({ token: tok.token, clientId, scopes, createdAt: new Date().toISOString(), expiresOn: new Date(tok.expiresOnTimestamp).toISOString() }, null, 2));
const claims = JSON.parse(Buffer.from(tok.token.split('.')[1], 'base64url').toString());
console.log(`Signed in as ${claims.upn || claims.preferred_username || claims.unique_name}`);
console.log(`Granted scopes: ${claims.scp}`);
console.log(`Saved ${authRecordFilePath} - the MCP server will now renew tokens silently.`);
