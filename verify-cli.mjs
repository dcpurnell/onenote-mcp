// Non-interactive check: can we get a Graph token silently from the persisted cache?
// Run:  node ~/tools/onenote-mcp/verify-cli.mjs
import { DeviceCodeCredential, useIdentityPlugin, deserializeAuthenticationRecord } from '@azure/identity';
import { cachePersistencePlugin } from '@azure/identity-cache-persistence';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const authRecordFilePath = path.join(__dirname, '.auth-record.json');
const clientId = process.env.AZURE_CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const scopes = ['Notes.Read', 'Notes.ReadWrite', 'Notes.Read.All', 'Notes.ReadWrite.All', 'Notes.Create', 'User.Read'];

useIdentityPlugin(cachePersistencePlugin);

const record = fs.existsSync(authRecordFilePath) ? deserializeAuthenticationRecord(fs.readFileSync(authRecordFilePath, 'utf8')) : undefined;
console.log('auth record:', record ? `${record.username} tenant=${record.tenantId}` : 'MISSING');

const credential = new DeviceCodeCredential({
  clientId,
  tokenCachePersistenceOptions: { enabled: true, name: 'onenote-mcp-token-cache', unsafeAllowUnencryptedStorage: true },
  disableAutomaticAuthentication: true,
  authenticationRecord: record,
});

try {
  const tok = await credential.getToken(scopes);
  const claims = JSON.parse(Buffer.from(tok.token.split('.')[1], 'base64url').toString());
  console.log('silent token OK for', claims.upn || claims.preferred_username);
  console.log('scopes:', claims.scp);
  const r = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks?$select=displayName', { headers: { Authorization: `Bearer ${tok.token}` } });
  console.log('Graph /me/onenote/notebooks:', r.status);
  if (r.ok) console.log((await r.json()).value.map(n => ' - ' + n.displayName).join('\n'));
  else console.log(await r.text());
} catch (e) {
  console.log('silent token FAILED:', e.message);
}
