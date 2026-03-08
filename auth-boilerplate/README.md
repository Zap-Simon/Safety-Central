# Microsoft Azure Authentication Boilerplate

This boilerplate contains all the code and configuration needed to implement Microsoft Azure authentication (MSAL 2.0) in a React application.

## What's Included

### Frontend Authentication Files
- `msalConfig.ts` - Azure MSAL configuration with client ID, tenant ID, and scopes
- `authService.ts` - Complete authentication service with login, logout, and token management
- `LoginPage.tsx` - Ready-to-use login component with Microsoft branding

### Backend Authentication Files
- `auth-middleware.ts` - Server-side middleware for token validation
- `protected-routes-examples.ts` - Examples of how to protect API endpoints

### Configuration Files
- `.env.example` - Environment variables template
- `package-dependencies.txt` - Required npm packages

### Setup Instructions
- `SETUP.md` - Step-by-step setup guide

## Quick Start

1. Copy the files to your new project
2. Install the required dependencies
3. Set up your Azure App Registration
4. Configure environment variables
5. Update the redirect URLs

See SETUP.md for detailed instructions.

## Features

- MSAL 2.0 with popup authentication
- Silent token refresh
- Secure token storage in session storage
- SharePoint API integration ready
- Server-side route protection
- Error handling and logging
- TypeScript support

## Security Features

- Secure cookie settings
- Token expiry validation
- Silent token acquisition
- Proper logout handling
- CSRF protection ready