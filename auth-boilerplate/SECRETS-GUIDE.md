# Azure Authentication Secrets Setup Guide

This guide explains how to obtain and configure the required secrets for Microsoft Azure authentication.

## Required Secrets

You'll need these secrets for your authentication system:

### 1. AZURE_CLIENT_ID (Application ID)
- **What it is**: Unique identifier for your Azure app registration
- **Where to find it**: Azure Portal > App registrations > Your app > Overview page
- **Example**: `dad12e09-3e7f-42fa-86e8-a0378bdd2699`
- **Security**: Safe to include in frontend code

### 2. AZURE_TENANT_ID (Directory ID)  
- **What it is**: Unique identifier for your Azure AD tenant/organization
- **Where to find it**: Azure Portal > App registrations > Your app > Overview page
- **Example**: `9c9325a5-65f3-4a30-9c57-e857ca0dd71a`
- **Security**: Safe to include in frontend code

### 3. AZURE_CLIENT_SECRET
- **What it is**: Secret key for server-side authentication
- **Where to find it**: Azure Portal > App registrations > Your app > Certificates & secrets
- **Example**: `dJa8Q~GxYzA1B2C3D4E5F6G7H8I9J0K1L2M3N4`
- **Security**: ⚠️ CRITICAL - Never expose in frontend, only use server-side

## Step-by-Step Secret Retrieval

### Getting Client ID and Tenant ID

1. **Go to Azure Portal**
   - Visit [portal.azure.com](https://portal.azure.com)
   - Sign in with your Azure account

2. **Navigate to App Registrations**
   - Click "Azure Active Directory" in the left sidebar
   - Click "App registrations"
   - Find and click your application

3. **Copy the IDs**
   - On the Overview page, you'll see:
     - **Application (client) ID** → Copy this for `AZURE_CLIENT_ID`
     - **Directory (tenant) ID** → Copy this for `AZURE_TENANT_ID`

### Getting Client Secret

1. **Go to Certificates & Secrets**
   - In your app registration, click "Certificates & secrets" in the left menu

2. **Create New Secret**
   - Click "New client secret"
   - Add description: "Auth Boilerplate Secret"
   - Choose expiration (recommend 12-24 months)
   - Click "Add"

3. **Copy Secret Value**
   - ⚠️ **IMPORTANT**: Copy the secret VALUE immediately
   - You won't be able to see it again after leaving this page
   - Store it securely for `AZURE_CLIENT_SECRET`

## Environment Variable Setup

### Development (.env file)
```bash
# Create .env file in your project root
VITE_AZURE_CLIENT_ID=your-client-id-here
VITE_AZURE_TENANT_ID=your-tenant-id-here  
AZURE_CLIENT_SECRET=your-client-secret-here
VITE_AZURE_REDIRECT_URI=http://localhost:3000
```

### Production (Replit Secrets)
For production deployment on Replit:

1. **Go to Replit Secrets**
   - Open your Replit project
   - Click the lock icon in the sidebar (Secrets)

2. **Add Each Secret**
   ```
   Key: VITE_AZURE_CLIENT_ID
   Value: your-client-id-here
   
   Key: VITE_AZURE_TENANT_ID  
   Value: your-tenant-id-here
   
   Key: AZURE_CLIENT_SECRET
   Value: your-client-secret-here
   
   Key: VITE_AZURE_REDIRECT_URI
   Value: https://your-replit-url.replit.app
   ```

## Security Best Practices

### Frontend Secrets (VITE_ prefixed)
- ✅ Safe to expose in browser
- Used for client-side authentication initialization
- Users can see these in browser dev tools (this is normal)

### Backend Secrets (no VITE_ prefix)
- ⚠️ Must stay server-side only
- Never include in frontend builds
- Used for server-to-server authentication

### Secret Rotation
- **Client secrets expire** - set calendar reminders
- **Rotate secrets regularly** (every 12-24 months)
- **Update all environments** when rotating

### Common Mistakes to Avoid
- ❌ Don't put AZURE_CLIENT_SECRET in frontend code
- ❌ Don't commit secrets to version control
- ❌ Don't share secrets in chat/email
- ❌ Don't use the same secret across environments

## Redirect URI Configuration

### Development URLs
Add these redirect URIs in Azure App Registration > Authentication:
```
http://localhost:3000
http://localhost:5000
http://127.0.0.1:3000
```

### Production URLs
Add your production URLs:
```
https://your-domain.com
https://your-replit-url.replit.app
```

### Important Notes
- URLs must match exactly (including https/http)
- No trailing slashes
- Ports must match what your dev server uses

## Testing Your Secrets

### Quick Test Script
Create a test file to verify your secrets work:

```javascript
// test-auth.js
const { PublicClientApplication } = require('@azure/msal-node');

const config = {
  auth: {
    clientId: process.env.VITE_AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.VITE_AZURE_TENANT_ID}`
  }
};

const pca = new PublicClientApplication(config);
console.log('MSAL configured successfully!');
console.log('Client ID:', process.env.VITE_AZURE_CLIENT_ID);
console.log('Tenant ID:', process.env.VITE_AZURE_TENANT_ID);
```

Run with: `node test-auth.js`

### Manual Token Test
Test your client secret manually:
```bash
curl -X POST \
  https://login.microsoftonline.com/YOUR_TENANT_ID/oauth2/v2.0/token \
  -H 'Content-Type: application/x-www-form-urlencoded' \
  -d 'client_id=YOUR_CLIENT_ID&client_secret=YOUR_CLIENT_SECRET&scope=https://graph.microsoft.com/.default&grant_type=client_credentials'
```

## Troubleshooting

### "Invalid Client" Error
- Check CLIENT_ID is correct
- Verify app registration exists
- Ensure app is not disabled

### "Invalid Secret" Error  
- Check CLIENT_SECRET is correct
- Verify secret hasn't expired
- Ensure no extra spaces/characters

### "Redirect URI Mismatch"
- Check redirect URI matches exactly
- Verify protocol (http vs https)
- Check for trailing slashes

### "Insufficient Privileges"
- Check API permissions are granted
- Verify admin consent if required
- Ensure user has necessary roles

## SharePoint Integration Secrets

If using SharePoint, you may need additional configuration:

### SharePoint App Permissions
1. **Go to SharePoint Admin Center**
   - Visit [admin.microsoft.com](https://admin.microsoft.com)
   - Navigate to SharePoint admin center

2. **App Permissions**
   - Go to Advanced > App permissions
   - Grant your app necessary permissions

### SharePoint-Specific Scopes
Update your scopes to include:
```typescript
export const sharePointRequest = {
  scopes: [
    'https://your-tenant.sharepoint.com/AllSites.FullControl'
  ],
};
```

## Getting Help

### Azure Support Resources
- [Azure AD Documentation](https://docs.microsoft.com/en-us/azure/active-directory/)
- [MSAL.js Documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-overview)
- [Microsoft Graph Documentation](https://docs.microsoft.com/en-us/graph/)

### Common Support Channels
- Azure Portal - Support tickets
- Microsoft Q&A forums
- Stack Overflow (tag: azure-active-directory)

Remember: Keep your secrets secure and never share them publicly!