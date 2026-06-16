import express, { type Request, Response, NextFunction } from "express";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import compression from "compression";
import pino from "pino";
import pinoHttp from "pino-http";
import crypto from "crypto";
import { registerRoutes } from "./routes";
import { setupVite, serveStatic, log } from "./vite";
import { printSetupInstructions } from "./secrets-helper";

// Production logger
const logger = pino({
  level: process.env.LOG_LEVEL || (process.env.NODE_ENV === 'production' ? 'info' : 'debug'),
  transport: process.env.NODE_ENV === 'development' ? {
    target: 'pino-pretty'
  } : undefined,
  redact: {
    paths: ['req.headers.authorization', 'req.headers.cookie'],
    censor: '[REDACTED]'
  }
});

const app = express();

// Trust proxy for correct IP detection in production
app.set('trust proxy', 1);

// Request logging (early for comprehensive coverage)
app.use(pinoHttp({ 
  logger,
  genReqId: () => crypto.randomUUID(),
  customSuccessMessage: (req, res) => `request completed`,
  autoLogging: {
    ignore: (req) => req.url.startsWith('/assets') || req.url.startsWith('/static')
  }
}));

// Set request ID in response headers for traceability
app.use((req, res, next) => {
  res.setHeader('x-request-id', req.id as string);
  next();
});

// Cache buster — HTML page must never be cached so browsers always get the latest bundle
app.use((req, res, next) => {
  const isHtmlRequest = req.headers.accept?.includes('text/html') && !req.path.startsWith('/api') && !req.path.startsWith('/assets');
  if (isHtmlRequest) {
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
  }
  next();
});

// Security headers with environment-specific CSP
const isProduction = process.env.NODE_ENV === 'production';
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      baseUri: ["'self'"],
      // Allow Teams, SharePoint, and Microsoft auth domains to embed the tab
      frameAncestors: ["'self'", "https://login.microsoftonline.com", "https://*.microsoftonline.com", "https://teams.microsoft.com", "https://*.teams.microsoft.com", "https://*.sharepoint.com", "https://*.office.com"],
      // UI libraries (Radix/shadcn) inject inline styles at runtime, so 'unsafe-inline'
      // is required for styles in production as well as dev.
      styleSrc: ["'self'", "'unsafe-inline'", "https:"],
      // Allow our own scripts plus the CDN-hosted mermaid library and the Replit dev banner.
      // In dev we additionally need 'unsafe-inline'/'unsafe-eval' for Vite HMR.
      scriptSrc: isProduction
        ? ["'self'", "https://unpkg.com", "https://replit.com"]
        : ["'self'", "'unsafe-inline'", "'unsafe-eval'", "https://unpkg.com", "https://replit.com"],
      imgSrc: ["'self'", "data:", "https:"],
      connectSrc: isProduction
        ? ["'self'", "https:", "https://login.microsoftonline.com", "https://graph.microsoft.com"]
        : ["'self'", "wss:", "https:", "https://login.microsoftonline.com", "https://graph.microsoft.com"], // Allow WebSocket for Vite HMR in dev and Microsoft APIs
      fontSrc: ["'self'", "https:"],
      // Allow MSAL ssoSilent() to open a hidden iframe to Microsoft login
      frameSrc: ["'self'", "https://login.microsoftonline.com", "https://*.microsoftonline.com"],
      // Allow forms to be submitted to Microsoft auth endpoints
      formAction: ["'self'", "https://login.microsoftonline.com", "https://*.microsoftonline.com"],
    },
  },
  // Disable X-Frame-Options so Teams can embed the /teams-tab page in an iframe.
  // CSP frame-ancestors above provides the actual security boundary.
  frameguard: false,
  // Disable cross-origin policies that interfere with Microsoft auth popups
  crossOriginOpenerPolicy: false,
  crossOriginResourcePolicy: false,
}));

// Compression
app.use(compression());

// Rate limiting (after logging for observability)
// Authenticated users are keyed by their token so each user gets an independent limit.
// Unauthenticated requests fall back to IP-based limiting.
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 500, // limit per key per windowMs
  message: { error: 'Too many requests, please try again later.' },
  standardHeaders: true,
  legacyHeaders: false,
  keyGenerator: (req) => {
    const auth = req.headers.authorization;
    if (auth && auth.startsWith('Bearer ')) {
      // Use the token itself as the key so each authenticated user has their own counter
      return auth.slice(7);
    }
    // Fall back to IP for unauthenticated requests
    return (req.ip || req.socket?.remoteAddress || 'unknown');
  },
});
app.use('/api', limiter);

// Body parsing
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: false, limit: '10mb' }));

(async () => {
  const server = await registerRoutes(app);

  app.use((err: any, req: Request, res: Response, _next: NextFunction) => {
    const status = err.status || err.statusCode || 500;
    
    // Log error with structured logging
    req.log.error({
      err,
      req: {
        method: req.method,
        url: req.url,
        headers: req.headers,
      },
      status,
    }, 'Request error');

    // Return safe error message in production
    const message = process.env.NODE_ENV === 'production' 
      ? (status < 500 ? err.message : 'Internal Server Error')
      : err.message || "Internal Server Error";

    res.status(status).json({ error: message });
  });

  // importantly only setup vite in development and after
  // setting up all the other routes so the catch-all route
  // doesn't interfere with the other routes
  if (app.get("env") === "development") {
    await setupVite(app, server);
  } else {
    serveStatic(app);
  }

  // ALWAYS serve the app on port 5000
  // this serves both the API and the client.
  // It is the only port that is not firewalled.
  const port = 5000;
  server.listen({
    port,
    host: "0.0.0.0",
    reusePort: true,
  }, () => {
    log(`serving on port ${port}`);
    
    // Show SharePoint setup status on startup
    if (process.env.NODE_ENV === 'development') {
      printSetupInstructions();
    }
  });
})();
