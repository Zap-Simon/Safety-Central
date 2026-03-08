import { Request, Response, NextFunction } from 'express';

// Extend Request interface to include user info
declare global {
  namespace Express {
    interface Request {
      user?: any;
      accessToken?: string;
    }
  }
}

/**
 * Middleware to authenticate requests using Bearer tokens
 * Extracts token from Authorization header and validates it
 */
export const authenticateToken = async (req: Request, res: Response, next: NextFunction) => {
  try {
    // Extract Bearer token from Authorization header
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({
        success: false,
        error: 'Authentication required. Please provide a valid Bearer token.',
        authenticated: false
      });
    }

    const accessToken = authHeader.substring(7); // Remove 'Bearer ' prefix
    
    if (!accessToken) {
      return res.status(401).json({
        success: false,
        error: 'No access token provided.',
        authenticated: false
      });
    }

    // Optional: Validate token with Microsoft Graph API
    // This step verifies the token is valid and gets user info
    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        return res.status(401).json({
          success: false,
          error: 'Invalid or expired access token.',
          authenticated: false
        });
      }

      const userInfo = await response.json();
      
      // Attach user info and token to request object for use in routes
      req.user = userInfo;
      req.accessToken = accessToken;
      
      next();
    } catch (tokenValidationError) {
      console.error('Token validation failed:', tokenValidationError);
      return res.status(401).json({
        success: false,
        error: 'Token validation failed.',
        authenticated: false
      });
    }

  } catch (error) {
    console.error('Authentication middleware error:', error);
    return res.status(500).json({
      success: false,
      error: 'Authentication service error.',
      authenticated: false
    });
  }
};

/**
 * Alternative middleware for form-based token authentication
 * Useful for protected document/file routes
 */
export const authenticateTokenFromForm = async (req: Request, res: Response, next: NextFunction) => {
  try {
    // Extract token from Authorization header or form data
    const authHeader = req.headers.authorization;
    const formToken = req.query.access_token || req.body.access_token;
    
    let accessToken = null;
    if (authHeader && authHeader.startsWith('Bearer ')) {
      accessToken = authHeader.substring(7);
    } else if (formToken) {
      accessToken = formToken;
    }

    if (!accessToken) {
      return res.status(401).send(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>Authentication Required</title>
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <style>
            body { 
              font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
              margin: 0; padding: 40px; background: #f8fafc; 
            }
            .container { 
              max-width: 500px; margin: 0 auto; background: white; 
              padding: 40px; border-radius: 12px; 
              box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; 
            }
            .error { color: #dc2626; font-weight: 600; margin-bottom: 20px; }
            .description { color: #64748b; margin-bottom: 30px; line-height: 1.6; }
            .button { 
              background: #0066cc; color: white; padding: 12px 24px; 
              border: none; border-radius: 8px; text-decoration: none; 
              display: inline-block; font-weight: 600; 
            }
          </style>
        </head>
        <body>
          <div class="container">
            <div class="error">Authentication Required</div>
            <p class="description">
              You need to be signed in to access this resource. 
              Please return to the main application and sign in with your Microsoft account.
            </p>
            <a href="/" class="button">Return to Application</a>
          </div>
        </body>
        </html>
      `);
    }

    // Validate token and get user info
    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error('Invalid token');
      }

      const userInfo = await response.json();
      req.user = userInfo;
      req.accessToken = accessToken;
      
      next();
    } catch (error) {
      return res.status(401).send(`
        <!DOCTYPE html>
        <html>
        <head>
          <title>Authentication Failed</title>
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <style>
            body { 
              font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
              margin: 0; padding: 40px; background: #f8fafc; 
            }
            .container { 
              max-width: 500px; margin: 0 auto; background: white; 
              padding: 40px; border-radius: 12px; 
              box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; 
            }
            .error { color: #dc2626; font-weight: 600; margin-bottom: 20px; }
            .description { color: #64748b; margin-bottom: 30px; line-height: 1.6; }
            .button { 
              background: #0066cc; color: white; padding: 12px 24px; 
              border: none; border-radius: 8px; text-decoration: none; 
              display: inline-block; font-weight: 600; 
            }
          </style>
        </head>
        <body>
          <div class="container">
            <div class="error">Authentication Failed</div>
            <p class="description">
              Your authentication token is invalid or has expired. 
              Please return to the main application and sign in again.
            </p>
            <a href="/" class="button">Return to Application</a>
          </div>
        </body>
        </html>
      `);
    }

  } catch (error) {
    console.error('Form authentication error:', error);
    return res.status(500).send('Authentication service error');
  }
};

/**
 * Middleware to check if user has specific permissions or roles
 * This is a template - customize based on your needs
 */
export const requireRole = (allowedRoles: string[]) => {
  return (req: Request, res: Response, next: NextFunction) => {
    if (!req.user) {
      return res.status(401).json({
        success: false,
        error: 'Authentication required.'
      });
    }

    // Example: Check user roles (customize based on your user object structure)
    const userRoles = req.user.roles || [];
    const hasRequiredRole = allowedRoles.some(role => userRoles.includes(role));

    if (!hasRequiredRole) {
      return res.status(403).json({
        success: false,
        error: `Access denied. Required roles: ${allowedRoles.join(', ')}`
      });
    }

    next();
  };
};