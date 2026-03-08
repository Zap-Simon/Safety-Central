# Microsoft Azure Authentication Setup Guide

This guide will help you set up Microsoft Azure authentication in your new project using this boilerplate.

## Prerequisites

- Azure account with permissions to create App Registrations
- Node.js and npm installed
- Basic knowledge of React and Express

## Step 1: Azure App Registration

1. **Go to Azure Portal**
   - Visit [Azure Portal](https://portal.azure.com)
   - Navigate to "Azure Active Directory" > "App registrations"

2. **Create New Registration**
   - Click "New registration"
   - Name: `Your App Name`
   - Supported account types: "Accounts in this organizational directory only"
   - Redirect URI: `Single-page application (SPA)` -> `https://your-domain.com`

3. **Configure Authentication**
   - Go to "Authentication" in your app registration
   - Add redirect URIs for development: `http://localhost:3000`, `http://localhost:5000`
   - Enable "Access tokens" and "ID tokens" under Implicit grant
   - Set logout URL: `https://your-domain.com`

4. **API Permissions**
   - Go to "API permissions"
   - Add permissions:
     - Microsoft Graph: `User.Read` (delegated)
     - Microsoft Graph: `Sites.Read.All` (delegated) - if using SharePoint
     - SharePoint: `AllSites.FullControl` (delegated) - if using SharePoint

5. **Certificates & Secrets** (for server-side operations)
   - Go to "Certificates & secrets"
   - Create a new client secret
   - Copy the secret value (you won't see it again!)

6. **Copy Important Values**
   - Application (client) ID
   - Directory (tenant) ID
   - Client secret (from step 5)

## Step 2: Project Setup

1. **Copy Boilerplate Files**
   ```bash
   # Copy the entire auth-boilerplate folder to your project
   cp -r auth-boilerplate/* your-new-project/
   ```

2. **Install Dependencies**
   ```bash
   cd your-new-project
   npm install @azure/msal-browser @azure/msal-node @types/react @types/react-dom react react-dom typescript lucide-react express @types/express @types/node tsx
   ```

3. **Environment Variables**
   ```bash
   # Copy the environment template
   cp .env.example .env
   
   # Edit .env with your values
   VITE_AZURE_CLIENT_ID=your-client-id-here
   VITE_AZURE_TENANT_ID=your-tenant-id-here
   AZURE_CLIENT_SECRET=your-client-secret-here
   VITE_AZURE_REDIRECT_URI=https://your-domain.com
   ```

## Step 3: Frontend Integration

1. **Update msalConfig.ts**
   - The config uses environment variables, so just ensure your .env is correct
   - Update SharePoint domain in `sharePointRequest.scopes` if needed

2. **Add to Your App.tsx**
   ```tsx
   import { useEffect } from 'react';
   import { authService } from './auth/authService';
   import LoginComponent from './auth/LoginComponent';

   function App() {
     useEffect(() => {
       // Initialize authentication on app startup
       authService.initialize().catch(console.error);
     }, []);

     return (
       <div className="App">
         <header>
           <LoginComponent />
         </header>
         {/* Your app content */}
       </div>
     );
   }
   ```

3. **Protected Routes Example**
   ```tsx
   import { useState, useEffect } from 'react';
   import { authService } from './auth/authService';

   function ProtectedPage() {
     const [isAuthenticated, setIsAuthenticated] = useState(false);

     useEffect(() => {
       setIsAuthenticated(authService.isAuthenticated());
     }, []);

     if (!isAuthenticated) {
       return <div>Please sign in to access this page.</div>;
     }

     return <div>Protected content here</div>;
   }
   ```

## Step 4: Backend Integration

1. **Add to your Express server**
   ```typescript
   import express from 'express';
   import { authenticateToken } from './auth/auth-middleware';
   import protectedRoutes from './auth/protected-routes-examples';

   const app = express();

   // Use the protected routes
   app.use(protectedRoutes);

   // Example protected endpoint
   app.get('/api/protected', authenticateToken, (req, res) => {
     res.json({ 
       message: 'This is protected!', 
       user: req.user 
     });
   });
   ```

2. **Frontend API Calls**
   ```typescript
   // Example of calling protected API from frontend
   const callProtectedAPI = async () => {
     try {
       const token = await authService.getAccessToken();
       
       const response = await fetch('/api/protected', {
         headers: {
           'Authorization': `Bearer ${token}`,
           'Content-Type': 'application/json'
         }
       });
       
       const data = await response.json();
       console.log(data);
     } catch (error) {
       console.error('API call failed:', error);
     }
   };
   ```

## Step 5: Testing

1. **Start your application**
   ```bash
   npm run dev
   ```

2. **Test Authentication Flow**
   - Click "Sign In" button
   - Should redirect to Microsoft login
   - After login, you should see user info
   - Try accessing protected endpoints

3. **Check Browser Console**
   - Should see MSAL logs
   - No authentication errors
   - Tokens being acquired successfully

## Step 6: Production Deployment

1. **Update Redirect URIs**
   - In Azure App Registration, add your production URL
   - Update .env with production values

2. **Security Considerations**
   - Use HTTPS in production
   - Keep client secrets secure
   - Regularly rotate client secrets
   - Monitor authentication logs

## Troubleshooting

### Common Issues

1. **CORS Errors**
   - Ensure redirect URIs match exactly in Azure
   - Check popup blockers

2. **Token Validation Fails**
   - Verify client ID and tenant ID
   - Check API permissions are granted

3. **SharePoint Access Denied**
   - Ensure user has access to SharePoint
   - Check SharePoint-specific permissions

4. **Silent Token Refresh Fails**
   - Normal behavior - fallback to popup works
   - Check token expiry times

### Debug Steps

1. **Enable MSAL Logging**
   - Already configured in msalConfig.ts
   - Check browser console for detailed logs

2. **Test Token Manually**
   ```bash
   # Test token validity
   curl -H "Authorization: Bearer YOUR_TOKEN" https://graph.microsoft.com/v1.0/me
   ```

3. **Check Network Tab**
   - Look for 401/403 responses
   - Verify Authorization headers

## Customization

### Adding Custom Scopes
Update `msalConfig.ts`:
```typescript
export const customRequest = {
  scopes: [
    'User.Read',
    'your.custom.scope'
  ],
};
```

### Role-Based Access
Use the `requireRole` middleware:
```typescript
app.get('/admin', authenticateToken, requireRole(['admin']), (req, res) => {
  // Admin only content
});
```

### Custom User Properties
Extend the middleware to fetch additional user info from your database or external services.

## Support

For issues specific to this boilerplate, check:
- Azure documentation
- MSAL.js documentation  
- Microsoft Graph API documentation

Remember to keep your client secrets secure and never commit them to version control!