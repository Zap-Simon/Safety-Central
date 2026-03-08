import { Configuration, PublicClientApplication } from '@azure/msal-browser';

// MSAL 2.0 configuration for SPA authentication
const msalConfig: Configuration = {
  auth: {
    clientId: process.env.VITE_AZURE_CLIENT_ID || 'your-client-id-here',
    authority: `https://login.microsoftonline.com/${process.env.VITE_AZURE_TENANT_ID || 'your-tenant-id-here'}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
    navigateToLoginRequestUrl: false, // MSAL 2.0 improvement for SPA
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
    secureCookies: true, // MSAL 2.0 security enhancement
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        console.log(`[MSAL] ${message}`);
      },
    },
    allowNativeBroker: false, // Disable native broker for web SPA
  },
};

// Production logging removed for cleaner output

// Create the MSAL instance
export const msalInstance = new PublicClientApplication(msalConfig);

// Graph API scopes for user authentication and basic access
export const loginRequest = {
  scopes: [
    'User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
  ],
};

// Separate SharePoint scopes for list access (update with your SharePoint domain)
export const sharePointRequest = {
  scopes: [
    'https://your-sharepoint-domain.sharepoint.com/AllSites.FullControl'
  ],
};

// Initialize MSAL with MSAL 2.0 pattern
export const initializeMsal = async () => {
  try {
    await msalInstance.initialize();
    // MSAL initialized successfully
  } catch (error) {
    console.error('❌ MSAL initialization failed:', error);
    throw error;
  }
};