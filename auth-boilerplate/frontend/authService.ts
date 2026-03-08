import { msalInstance, loginRequest, sharePointRequest } from './msalConfig';
import { AuthenticationResult, SilentRequest, InteractionRequiredAuthError } from '@azure/msal-browser';

export class AuthService {
  private static instance: AuthService;
  private accessToken: string | null = null;
  private tokenExpiry: number = 0;

  private constructor() {}

  static getInstance(): AuthService {
    if (!AuthService.instance) {
      AuthService.instance = new AuthService();
    }
    return AuthService.instance;
  }

  async initialize(): Promise<void> {
    try {
      await msalInstance.initialize();
      
      // Handle redirect response if coming back from authentication
      const response = await msalInstance.handleRedirectPromise();
      if (response) {
        this.handleAuthResult(response);
      }
      
      // Check for existing accounts and acquire token silently
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        await this.acquireTokenSilently();
      }
    } catch (error) {
      console.error('MSAL initialization failed:', error);
    }
  }

  async signIn(): Promise<void> {
    try {
      const response = await msalInstance.loginPopup(loginRequest);
      if (response) {
        this.handleAuthResult(response);
      }
    } catch (error) {
      console.error('Sign in failed:', error);
      throw error;
    }
  }

  private async acquireTokenSilently(): Promise<void> {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return;

    try {
      const silentRequest: SilentRequest = {
        scopes: loginRequest.scopes,
        account: accounts[0],
      };

      const response = await msalInstance.acquireTokenSilent(silentRequest);
      this.handleAuthResult(response);
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        // Silent token acquisition requires user interaction
        console.log('Silent token acquisition failed, interaction required');
      } else {
        console.error('Silent token acquisition failed:', error);
      }
    }
  }

  async getAccessToken(): Promise<string> {
    // Ensure MSAL is initialized
    if (!msalInstance.getConfiguration()) {
      await msalInstance.initialize();
    }

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      throw new Error('No authenticated accounts found');
    }

    // Check if we have a valid cached token
    if (this.accessToken && Date.now() < this.tokenExpiry) {
      return this.accessToken;
    }

    try {
      const silentRequest: SilentRequest = {
        scopes: loginRequest.scopes,
        account: accounts[0],
      };

      const response = await msalInstance.acquireTokenSilent(silentRequest);
      this.handleAuthResult(response);
      return this.accessToken!;
    } catch (error) {
      // Handle InteractionRequiredAuthError specifically
      if (error instanceof InteractionRequiredAuthError) {
        try {
          const response = await msalInstance.acquireTokenPopup(loginRequest);
          this.handleAuthResult(response);
          return this.accessToken!;
        } catch (interactiveError) {
          console.error('Interactive token acquisition failed:', interactiveError);
          throw interactiveError;
        }
      } else {
        console.error('Token acquisition failed:', error);
        throw error;
      }
    }
  }

  isAuthenticated(): boolean {
    const accounts = msalInstance.getAllAccounts();
    return accounts.length > 0 && this.accessToken !== null && Date.now() < this.tokenExpiry;
  }

  async signOut(): Promise<void> {
    try {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        // Use popup logout instead of redirect to avoid iframe issues
        await msalInstance.logoutPopup({
          account: accounts[0],
          mainWindowRedirectUri: window.location.origin
        });
      }
    } catch (error) {
      console.error('Sign out failed:', error);
      // Clear local state even if popup logout fails
    }
    this.accessToken = null;
    this.tokenExpiry = 0;
  }

  /**
   * Get SharePoint-specific access token for list operations
   * Uses different scopes than the main login token
   */
  async getSharePointToken(): Promise<string> {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      throw new Error('No authenticated accounts found');
    }

    try {
      // Try to get SharePoint token silently
      const silentRequest: SilentRequest = {
        scopes: sharePointRequest.scopes,
        account: accounts[0],
      };

      const response = await msalInstance.acquireTokenSilent(silentRequest);
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        // Need user interaction for SharePoint permissions
        const response = await msalInstance.acquireTokenPopup(sharePointRequest);
        return response.accessToken;
      } else {
        console.error('SharePoint token acquisition failed:', error);
        throw error;
      }
    }
  }

  private handleAuthResult(result: AuthenticationResult): void {
    this.accessToken = result.accessToken;
    this.tokenExpiry = result.expiresOn ? result.expiresOn.getTime() : Date.now() + 3600000;
  }

  getCurrentUser(): any {
    const accounts = msalInstance.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  }

  /**
   * Navigate to an authenticated protected resource
   * This example shows how to pass auth tokens to server routes
   */
  async openProtectedResource(path: string): Promise<void> {
    try {
      const token = await this.getAccessToken();
      
      // Create a form with the auth token and submit to open in new tab
      const form = document.createElement('form');
      form.method = 'GET';
      form.action = path;
      form.target = '_blank';
      
      // Add authorization header via hidden input (handled server-side)
      const tokenInput = document.createElement('input');
      tokenInput.type = 'hidden';
      tokenInput.name = 'access_token';
      tokenInput.value = token;
      form.appendChild(tokenInput);
      
      document.body.appendChild(form);
      form.submit();
      document.body.removeChild(form);
    } catch (error) {
      console.error('Failed to open protected resource:', error);
      throw error;
    }
  }
}

export const authService = AuthService.getInstance();