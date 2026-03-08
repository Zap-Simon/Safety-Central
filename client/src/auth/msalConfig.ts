import { Configuration, PublicClientApplication } from '@azure/msal-browser';

// MSAL 2.0 configuration for SPA authentication
const msalConfig: Configuration = {
  auth: {
    clientId: 'dad12e09-3e7f-42fa-86e8-a0378bdd2699',
    authority: 'https://login.microsoftonline.com/9c9325a5-65f3-4a30-9c57-e857ca0dd71a',
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
        if (import.meta.env.MODE === 'development') {
          console.log(`[MSAL] ${message}`);
        }
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

// Separate SharePoint scopes for list access
export const sharePointRequest = {
  scopes: [
    'https://cranfieldglass.sharepoint.com/AllSites.FullControl'
  ],
};

// Initialize MSAL with MSAL 2.0 pattern
export const initializeMsal = async () => {
  try {
    await msalInstance.initialize();
    // MSAL initialized successfully
  } catch (error) {
    console.error('MSAL initialization failed:', error);
    throw error;
  }
};