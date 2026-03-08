# New Replit Project Setup Guide - Financial Pricing System

## Authentication Setup (Microsoft 365 Integration)

### Required Environment Variables
Add these to your new Replit project's Secrets:

```bash
# Azure/Microsoft Authentication
AZURE_TENANT_ID=9c9325a5-65f3-4a30-9c57-e857ca0dd71a
AZURE_CLIENT_ID=dad12e09-3e7f-42fa-86e8-a0378bdd2699
AZURE_CLIENT_SECRET=[Ask user for this secret]
AZURE_REDIRECT_URI=https://your-new-replit-url.replit.dev

# Database (if needed)
DATABASE_URL=[Your PostgreSQL connection string]

# OpenAI (if needed)
OPENAI_API_KEY=[Your OpenAI API key]
```

### Frontend Authentication Files to Copy

**1. Create: `/client/src/auth/msalConfig.ts`**
```typescript
import { Configuration } from '@azure/msal-browser';

// MSAL 2.0 configuration for SPA authentication
const msalConfig: Configuration = {
  auth: {
    clientId: 'dad12e09-3e7f-42fa-86e8-a0378bdd2699',
    authority: 'https://login.microsoftonline.com/9c9325a5-65f3-4a30-9c57-e857ca0dd71a',
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
    secureCookies: true,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        console.log(`[MSAL] ${message}`);
      },
    },
    allowNativeBroker: false,
  },
};

export default msalConfig;
```

**2. Create: `/client/src/auth/authService.ts`**
```typescript
import { PublicClientApplication, AuthenticationResult } from '@azure/msal-browser';
import msalConfig from './msalConfig';

const loginRequest = {
  scopes: ['https://graph.microsoft.com/.default'],
  prompt: 'select_account',
};

class AuthService {
  private msalInstance: PublicClientApplication;

  constructor() {
    this.msalInstance = new PublicClientApplication(msalConfig);
  }

  async initialize(): Promise<void> {
    try {
      await this.msalInstance.initialize();
      
      const response = await this.msalInstance.handleRedirectPromise();
      if (response) {
        this.handleAuthResult(response);
      }
      
      const accounts = this.msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        await this.acquireTokenSilently();
      }
    } catch (error) {
      console.error('MSAL initialization failed:', error);
    }
  }

  async signIn(): Promise<void> {
    try {
      const response = await this.msalInstance.loginPopup(loginRequest);
      if (response) {
        this.handleAuthResult(response);
      }
    } catch (error) {
      console.error('Sign in failed:', error);
      throw error;
    }
  }

  async signOut(): Promise<void> {
    try {
      await this.msalInstance.logoutPopup();
    } catch (error) {
      console.error('Sign out failed:', error);
      throw error;
    }
  }

  private handleAuthResult(response: AuthenticationResult): void {
    // Store token and user info
    localStorage.setItem('authToken', response.accessToken);
    localStorage.setItem('userInfo', JSON.stringify(response.account));
  }

  private async acquireTokenSilently(): Promise<string | null> {
    try {
      const accounts = this.msalInstance.getAllAccounts();
      if (accounts.length === 0) return null;

      const response = await this.msalInstance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });

      localStorage.setItem('authToken', response.accessToken);
      return response.accessToken;
    } catch (error) {
      console.error('Silent token acquisition failed:', error);
      return null;
    }
  }

  isAuthenticated(): boolean {
    return !!localStorage.getItem('authToken');
  }

  getToken(): string | null {
    return localStorage.getItem('authToken');
  }

  getUserInfo(): any {
    const userInfo = localStorage.getItem('userInfo');
    return userInfo ? JSON.parse(userInfo) : null;
  }
}

export default new AuthService();
```

### Backend Authentication Files to Copy

**1. Create: `/server/sharepoint-app-auth.ts`**
```typescript
interface AzureAppCredentials {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

export class SharePointAppAuthService {
  private accessToken: string | null = null;
  private tokenExpiry: number = 0;
  private credentials: AzureAppCredentials;

  constructor() {
    this.credentials = {
      tenantId: process.env.AZURE_TENANT_ID!,
      clientId: process.env.AZURE_CLIENT_ID!,
      clientSecret: process.env.AZURE_CLIENT_SECRET!
    };
  }

  async getSharePointAccessToken(): Promise<string> {
    if (this.accessToken && Date.now() < this.tokenExpiry) {
      return this.accessToken;
    }

    const tokenUrl = `https://login.microsoftonline.com/${this.credentials.tenantId}/oauth2/v2.0/token`;
    
    const params = new URLSearchParams();
    params.append('client_id', this.credentials.clientId);
    params.append('client_secret', this.credentials.clientSecret);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');

    try {
      const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params,
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Token request failed: ${response.status} ${response.statusText} - ${errorText}`);
      }

      const tokenData = await response.json();
      
      if (!tokenData.access_token) {
        throw new Error('No access token received from Azure');
      }
      
      this.accessToken = tokenData.access_token;
      this.tokenExpiry = Date.now() + (tokenData.expires_in - 300) * 1000;
      
      console.log('✅ SharePoint application token acquired successfully');
      return this.accessToken!;
    } catch (error) {
      console.error('❌ Failed to acquire SharePoint token:', error);
      throw error;
    }
  }

  isConfigured(): boolean {
    return !!(this.credentials.tenantId && this.credentials.clientId && this.credentials.clientSecret);
  }
}
```

## Complete Styling Setup

### 1. Install Required Packages
```bash
npm install @azure/msal-browser @azure/msal-node tailwindcss @tailwindcss/typography lucide-react clsx tailwind-merge class-variance-authority
```

### 2. Create: `/client/src/index.css`
```css
@tailwind base;
@tailwind components;
@tailwind utilities;

:root {
  --background: hsl(0, 0%, 100%);
  --foreground: hsl(20, 14.3%, 4.1%);
  --muted: hsl(60, 4.8%, 95.9%);
  --muted-foreground: hsl(25, 5.3%, 44.7%);
  --popover: hsl(0, 0%, 100%);
  --popover-foreground: hsl(20, 14.3%, 4.1%);
  --card: hsl(0, 0%, 100%);
  --card-foreground: hsl(20, 14.3%, 4.1%);
  --border: hsl(20, 5.9%, 90%);
  --input: hsl(20, 5.9%, 90%);
  --primary: hsl(207, 90%, 54%);
  --primary-foreground: hsl(211, 100%, 99%);
  --secondary: hsl(60, 4.8%, 95.9%);
  --secondary-foreground: hsl(24, 9.8%, 10%);
  --accent: hsl(60, 4.8%, 95.9%);
  --accent-foreground: hsl(24, 9.8%, 10%);
  --destructive: hsl(0, 84.2%, 60.2%);
  --destructive-foreground: hsl(60, 9.1%, 97.8%);
  --ring: hsl(20, 14.3%, 4.1%);
  --radius: 0.5rem;
  
  /* Microsoft colors */
  --ms-blue: hsl(207, 100%, 41%);
  --ms-blue-dark: hsl(208, 84%, 40%);
  --ms-blue-light: hsl(191, 100%, 48%);
  --ms-green: hsl(120, 86%, 25%);
  --ms-amber: hsl(45, 100%, 50%);
  --ms-red: hsl(356, 70%, 51%);
  --ms-gray-900: hsl(25, 3%, 19%);
  --ms-gray-700: hsl(25, 5%, 38%);
  --ms-gray-500: hsl(25, 4%, 54%);
  --ms-gray-100: hsl(40, 14%, 95%);
}

.dark {
  --background: hsl(240, 10%, 3.9%);
  --foreground: hsl(0, 0%, 98%);
  --muted: hsl(240, 3.7%, 15.9%);
  --muted-foreground: hsl(240, 5%, 64.9%);
  --popover: hsl(240, 10%, 3.9%);
  --popover-foreground: hsl(0, 0%, 98%);
  --card: hsl(240, 10%, 3.9%);
  --card-foreground: hsl(0, 0%, 98%);
  --border: hsl(240, 3.7%, 15.9%);
  --input: hsl(240, 3.7%, 15.9%);
  --primary: hsl(207, 90%, 54%);
  --primary-foreground: hsl(211, 100%, 99%);
  --secondary: hsl(240, 3.7%, 15.9%);
  --secondary-foreground: hsl(0, 0%, 98%);
  --accent: hsl(240, 3.7%, 15.9%);
  --accent-foreground: hsl(0, 0%, 98%);
  --destructive: hsl(0, 62.8%, 30.6%);
  --destructive-foreground: hsl(0, 0%, 98%);
  --ring: hsl(240, 4.9%, 83.9%);
  --radius: 0.5rem;
}

@layer base {
  * {
    @apply border-border;
  }

  body {
    @apply font-sans antialiased bg-background text-foreground;
    font-family: 'Inter', sans-serif;
  }
}

/* Mobile-friendly utilities */
@layer utilities {
  .line-clamp-2 {
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
  }
  
  .line-clamp-1 {
    display: -webkit-box;
    -webkit-line-clamp: 1;
    -webkit-box-orient: vertical;
    overflow: hidden;
  }
  
  .btn-mobile {
    min-height: 44px;
    padding: 12px 16px;
  }
  
  input[type="text"],
  input[type="email"],
  input[type="password"],
  textarea,
  select {
    font-size: 16px;
  }
  
  @media (min-width: 640px) {
    input[type="text"],
    input[type="email"], 
    input[type="password"],
    textarea,
    select {
      font-size: 0.875rem;
    }
  }
}

/* Microsoft theme colors */
.ms-blue {
  background-color: var(--ms-blue);
}

.ms-blue-dark {
  background-color: var(--ms-blue-dark);
}

.ms-green {
  background-color: var(--ms-green);
}

.ms-amber {
  background-color: var(--ms-amber);
}

.ms-red {
  background-color: var(--ms-red);
}
```

### 3. Create: `/tailwind.config.ts`
```typescript
import type { Config } from 'tailwindcss';

const config: Config = {
  darkMode: ["class"],
  content: [
    './client/src/**/*.{js,ts,jsx,tsx,mdx}',
    './client/index.html',
  ],
  theme: {
    extend: {
      borderRadius: {
        lg: 'var(--radius)',
        md: 'calc(var(--radius) - 2px)',
        sm: 'calc(var(--radius) - 4px)',
      },
      colors: {
        background: 'hsl(var(--background))',
        foreground: 'hsl(var(--foreground))',
        card: {
          DEFAULT: 'hsl(var(--card))',
          foreground: 'hsl(var(--card-foreground))',
        },
        popover: {
          DEFAULT: 'hsl(var(--popover))',
          foreground: 'hsl(var(--popover-foreground))',
        },
        primary: {
          DEFAULT: 'hsl(var(--primary))',
          foreground: 'hsl(var(--primary-foreground))',
        },
        secondary: {
          DEFAULT: 'hsl(var(--secondary))',
          foreground: 'hsl(var(--secondary-foreground))',
        },
        muted: {
          DEFAULT: 'hsl(var(--muted))',
          foreground: 'hsl(var(--muted-foreground))',
        },
        accent: {
          DEFAULT: 'hsl(var(--accent))',
          foreground: 'hsl(var(--accent-foreground))',
        },
        destructive: {
          DEFAULT: 'hsl(var(--destructive))',
          foreground: 'hsl(var(--destructive-foreground))',
        },
        border: 'hsl(var(--border))',
        input: 'hsl(var(--input))',
        ring: 'hsl(var(--ring))',
      },
    },
  },
  plugins: [],
};

export default config;
```

## Package.json Dependencies
Add these to your new project:

```json
{
  "dependencies": {
    "@azure/msal-browser": "^3.28.1",
    "@azure/msal-node": "^2.15.0",
    "react": "^18.2.0",
    "react-dom": "^18.2.0",
    "tailwindcss": "^3.4.0",
    "lucide-react": "^0.446.0",
    "clsx": "^2.1.0",
    "tailwind-merge": "^2.2.0",
    "class-variance-authority": "^0.7.0"
  }
}
```

## Quick Start Steps

1. **Create new Replit project**
2. **Add all environment variables to Secrets**
3. **Copy all authentication files above**
4. **Copy styling files above**
5. **Install dependencies**
6. **Initialize auth in your main App component:**

```typescript
import { useEffect } from 'react';
import authService from './auth/authService';

function App() {
  useEffect(() => {
    authService.initialize();
  }, []);

  return (
    <div className="min-h-screen bg-background">
      {/* Your financial pricing app here */}
    </div>
  );
}
```

## Missing Secret You'll Need

You'll need to get the `AZURE_CLIENT_SECRET` from the original project owner since it's not visible in the code for security reasons.

This setup will give you:
- ✅ Same Microsoft 365 authentication
- ✅ Exact same styling and color scheme
- ✅ Mobile-optimized interface
- ✅ Professional corporate appearance
- ✅ Same authentication flow and tokens

The financial pricing system will look and authenticate exactly like the main health & safety system!