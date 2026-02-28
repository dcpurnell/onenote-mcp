# Authentication & Permissions Troubleshooting Guide

## Quick Diagnosis: Check Your Token Scopes

If you're authenticated but `listNotebooks` returns empty results, the most common cause is **missing API permissions** in your Azure AD app registration.

### Step 1: Run the Diagnostic Tool

Use the new `checkTokenScopes` tool to see what permissions your current token actually has:

```
Tool: checkTokenScopes
```

This will show you:
- ✅ All scopes currently granted to your token
- ❌ Any missing required scopes
- 📋 Token metadata (app ID, expiration, etc.)

### What You Should See

**✅ Good Token (has required scopes):**
```
Granted Scopes:
  • Notes.Read.All
  • Notes.ReadWrite.All
  • User.Read
  • ...
```

**❌ Problem Token (missing required scopes):**
```
Granted Scopes:
  • User.Read
  
ISSUE FOUND: Token is missing required scopes!

Missing Scopes:
  • Notes.Read.All
  • Notes.ReadWrite.All
```

---

## Problem: Missing Required Scopes

If `checkTokenScopes` shows you're missing `Notes.Read.All` or `Notes.ReadWrite.All`, this means your Azure AD app registration doesn't have the right API permissions configured.

### Why This Happens

The MCP server code requests these scopes:
```javascript
const scopes = [
  'Notes.Read', 
  'Notes.ReadWrite', 
  'Notes.Read.All',      // ← Required for ALL notebooks including shared/team
  'Notes.ReadWrite.All', // ← Required for ALL notebooks including shared/team
  'Notes.Create', 
  'User.Read'
];
```

**BUT** Azure AD will only grant scopes that are:
1. **Configured** in the app registration's API permissions
2. **Consented** by you (or an admin for `.All` scopes)

Even if your code asks for a scope, if the app registration doesn't allow it, you won't get it in your token.

---

## Solution: Fix Your Azure AD App Registration

### Default App (Microsoft Graph Explorer)

If you're using the **default client ID** (`14d82eec-204b-4c2f-b7e8-296a70dab67e`), this is Microsoft's Graph Explorer app, which may have limited permissions.

**✅ Recommended:** Create your own Azure AD app registration with full control over permissions.

### Option 1: Create Your Own Azure AD App (Recommended)

#### 1. Go to Azure Portal
Navigate to: https://portal.azure.com → **Azure Active Directory** → **App registrations** → **New registration**

#### 2. Register the Application
- **Name:** `OneNote MCP Server` (or any name you prefer)
- **Supported account types:** 
  - Choose "Accounts in this organizational directory only" for work/school accounts
  - OR "Accounts in any organizational directory and personal Microsoft accounts" for broader access
- **Redirect URI:** Leave blank (we're using device code flow)
- Click **Register**

#### 3. Note Your Application (client) ID
- Copy the **Application (client) ID** from the overview page
- Save this - you'll need it for the MCP server

#### 4. Add API Permissions
Go to **API permissions** → **Add a permission**:

**a) Add Microsoft Graph Delegated Permissions:**
- Click **Microsoft Graph** → **Delegated permissions**
- Search for and add:
  - ✅ `Notes.Read` - Read user OneNote notebooks
  - ✅ `Notes.ReadWrite` - Read and write user OneNote notebooks  
  - ✅ `Notes.Read.All` - **Read all OneNote notebooks that user can access**
  - ✅ `Notes.ReadWrite.All` - **Read and write all OneNote notebooks**
  - ✅ `Notes.Create` - Create pages in user notebooks
  - ✅ `User.Read` - Sign in and read user profile

**b) Grant Admin Consent (if required):**
- The `.All` scopes typically require admin consent
- If you're an admin: Click **Grant admin consent for [Your Organization]**
- If not: Ask your IT admin to grant consent, OR use personal Microsoft account which doesn't require admin consent

#### 5. Enable Public Client Flow
- Go to **Authentication** → **Advanced settings**
- Set **Allow public client flows** to **Yes**
- Click **Save**

#### 6. Update Your MCP Server Configuration

**Option A: Environment Variable (recommended)**
```bash
export AZURE_CLIENT_ID="your-app-id-here"
```

**Option B: Edit the Code**
Edit `onenote-mcp.mjs` line ~16:
```javascript
const clientId = 'your-app-id-here';
```

#### 7. Re-authenticate
- Delete the old token: `rm .access-token.txt`
- Run the `authenticate` tool again
- Complete the device code flow
- Run `saveAccessToken` to verify
- Run `checkTokenScopes` to confirm you now have all required scopes

---

### Option 2: Verify Existing App Registration

If you already have an Azure AD app, verify it has the correct permissions:

#### 1. Find Your App
Azure Portal → **Azure Active Directory** → **App registrations** → Search for your app

#### 2. Check API Permissions
Go to **API permissions** and verify you have:
- Microsoft Graph → Delegated:
  - Notes.Read
  - Notes.ReadWrite
  - **Notes.Read.All** ← Critical for shared/team notebooks
  - **Notes.ReadWrite.All** ← Critical for shared/team notebooks
  - Notes.Create
  - User.Read

#### 3. Check Consent Status
- Look at the **Status** column
- Should say "Granted for [Organization]" or "Granted"
- If it says "Not granted", click **Grant admin consent**

#### 4. Verify Public Client Flow
**Authentication** → **Advanced settings** → **Allow public client flows** = **Yes**

---

## After Fixing App Registration

### Complete Re-authentication Flow

1. **Delete old token:**
   ```bash
   rm .access-token.txt
   ```

2. **Authenticate with new app:**
   - Run `authenticate` tool
   - Follow device code flow
   - Sign in and grant consent to all requested permissions

3. **Verify token:**
   - Run `saveAccessToken`
   - Run `checkTokenScopes`
   - Confirm you see `Notes.Read.All` and `Notes.ReadWrite.All`

4. **Test notebook access:**
   - Run `listNotebooks`
   - Should now see your notebooks!

---

## Understanding Scope Differences

| Scope | What It Does | Required For |
|-------|-------------|--------------|
| `Notes.Read` | Read notebooks **owned by the user** | Basic personal notebook access |
| `Notes.ReadWrite` | Read/write notebooks **owned by the user** | Editing personal notebooks |
| `Notes.Read.All` | Read **all notebooks user can access** | Shared notebooks, Team notebooks |
| `Notes.ReadWrite.All` | Read/write **all notebooks user can access** | Editing shared/team notebooks |
| `Notes.Create` | Create new pages/sections | Creating content |

**Key Insight:** Without `.All` scopes, you can only access notebooks you directly own, NOT:
- Notebooks shared with you from OneDrive
- Notebooks in Microsoft Teams
- Notebooks in SharePoint sites

---

## Common Issues & Solutions

### Issue: "No notebooks found" even after re-authentication

**Possible causes:**
1. ✅ Check scopes with `checkTokenScopes` - verify `.All` scopes are present
2. 🔄 Token cached - try deleting `.access-token.txt` and re-authenticate
3. 📋 Actually no notebooks - verify in OneNote web/app that you have notebooks
4. 🏢 Organizational policy - some enterprises restrict `.All` permissions

### Issue: "Admin consent required"

**Solution:**
- Use a personal Microsoft account (no admin needed), OR
- Request IT admin to grant consent, OR  
- If you're admin: Grant consent in Azure Portal

### Issue: Device code authentication times out

**Solution:**
- Browser must stay open during sign-in
- Complete within 15 minutes
- Check firewall/proxy isn't blocking Microsoft login
- Try different browser or incognito mode

### Issue: Graph API returns 403 Forbidden

**Causes:**
- Missing scopes (check with `checkTokenScopes`)
- Token expired (tokens last ~1 hour, refresh by re-authenticating)
- Organizational policy blocking API access

---

## Verification Checklist

After completing setup, verify everything works:

- [ ] `checkTokenScopes` shows Notes.Read.All and Notes.ReadWrite.All
- [ ] `listNotebooks` returns your notebooks
- [ ] `listSections` works for at least one notebook
- [ ] `getPageContent` can read a page
- [ ] Team notebooks appear when using `includeTeamNotebooks: true`

---

## Still Having Issues?

### Enable Debug Logging

The MCP server logs to stderr. When running via MCP inspector or Claude Desktop, check the logs for:
- Token loading messages
- Graph API request URLs
- Error responses with details

### Manual API Test

Test the Graph API directly with your token:

```bash
# Get your token
TOKEN=$(cat .access-token.txt | jq -r '.token')

# Test notebooks endpoint
curl -H "Authorization: Bearer $TOKEN" \
  https://graph.microsoft.com/v1.0/me/onenote/notebooks

# Should return JSON with notebooks array
```

If this works but MCP tool doesn't, it's a server issue. If this fails too, it's an authentication/permission issue.

---

## Security Notes

- **Token Storage:** Tokens are stored in `.access-token.txt` - keep this file secure
- **Token Expiration:** Tokens expire after ~1 hour; re-authenticate when needed
- **Scope Principle:** Only request permissions your app actually needs
- **Production Use:** For production, consider using a proper secret management service and refresh tokens

---

## Quick Reference: Required Files & Permissions

### Azure AD App Configuration
```
✅ API Permissions (Delegated):
   - Notes.Read
   - Notes.ReadWrite  
   - Notes.Read.All        ← Critical
   - Notes.ReadWrite.All   ← Critical
   - Notes.Create
   - User.Read

✅ Authentication:
   - Public client flows: Enabled

✅ Consent:
   - Admin consent granted (for .All scopes)
```

### MCP Server Configuration
```javascript
// In onenote-mcp.mjs
const clientId = process.env.AZURE_CLIENT_ID || 'your-app-id';
const scopes = [
  'Notes.Read', 
  'Notes.ReadWrite', 
  'Notes.Read.All',      // ← Must match app registration
  'Notes.ReadWrite.All', // ← Must match app registration
  'Notes.Create', 
  'User.Read'
];
```

---

**Last Updated:** February 28, 2026
