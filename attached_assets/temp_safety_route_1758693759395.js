  // Clean safety glazing standards proxy endpoint
  app.get('/api/safety-standards', async (req, res) => {
    try {
      const targetUrl = 'https://www.wganz.org.nz/safety-glazing-standards/';
      
      // Clean up expired cache entries
      await pageCacheStorage.clearExpiredCache();
      
      // Check for cached version first
      const cachedPage = await pageCacheStorage.getCachedPage(targetUrl);
      let html;
      
      if (cachedPage) {
        console.log('Serving from cache - instant load');
        html = cachedPage.content;
      } else {
        console.log('Fetching fresh content');
        
        const response = await fetch(targetUrl, {
          headers: {
            'User-Agent': req.get('User-Agent') || 'Mozilla/5.0 (compatible; GlazierMate/1.0)'
          }
        });
        
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        html = await response.text();
        
        // Strip slow external content
        html = html.replace(/<script[^>]*src=[^>]*google[^>]*><\/script>/gi, '');
        html = html.replace(/<script[^>]*src=[^>]*facebook[^>]*><\/script>/gi, '');
        html = html.replace(/<script[^>]*src=[^>]*twitter[^>]*><\/script>/gi, '');
        html = html.replace(/<link[^>]*href=[^>]*fonts\.googleapis[^>]*>/gi, '');
        
        // Cache the optimized version
        await pageCacheStorage.setCachedPage(targetUrl, html, 'text/html', 6);
      }
      
      // Add simple navigation header
      const headerHTML = `
        <div style="position: fixed; top: 0; left: 0; right: 0; z-index: 999999; background: hsl(215, 25%, 27%); padding: 16px; display: flex; align-items: center; gap: 16px; border-bottom: 1px solid rgba(255, 255, 255, 0.15);">
          <button onclick="sessionStorage.setItem('returning-from-safety-standards', 'true'); window.location.href='/';" style="display: flex; align-items: center; justify-content: center; width: 44px; height: 44px; background: rgba(255, 255, 255, 0.15); border: 1px solid rgba(255, 255, 255, 0.25); border-radius: 12px; color: white; font-size: 20px; font-weight: 700; cursor: pointer; outline: none;">
            ←
          </button>
          <div style="flex: 1;">
            <h1 style="margin: 0 0 2px 0; color: white; font-size: 20px; font-weight: 700; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;">Safety Glazing Standards</h1>
            <p style="margin: 0; color: rgba(255, 255, 255, 0.85); font-size: 14px; font-weight: 500; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;">Professional Glazier Calculator</p>
          </div>
        </div>
        <style>
          body { padding-top: 76px !important; margin: 0 !important; }
          @media (max-width: 768px) {
            body { padding-top: 76px !important; }
          }
        </style>
      `;
      
      // Inject header after body tag
      const bodyIndex = html.indexOf('<body');
      if (bodyIndex > -1) {
        const bodyTagEnd = html.indexOf('>', bodyIndex) + 1;
        html = html.substring(0, bodyTagEnd) + headerHTML + html.substring(bodyTagEnd);
      }
      
      res.setHeader('Content-Type', 'text/html');
      res.setHeader('Cache-Control', 'no-cache');
      res.send(html);
      
      console.log('Header provides navigation back to calculator');
      
    } catch (error) {
      console.error('Safety standards proxy error:', error.message);
      
      const offlineHTML = `
        <html lang="en">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Safety Standards - Offline</title>
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
              margin: 0;
              padding: 20px;
              background: hsl(215, 25%, 27%);
              min-height: 100vh;
              display: flex;
              align-items: center;
              justify-content: center;
              color: white;
            }
            .container {
              background: rgba(255, 255, 255, 0.1);
              backdrop-filter: blur(10px);
              border-radius: 20px;
              padding: 40px;
              text-align: center;
              max-width: 400px;
              border: 1px solid rgba(255, 255, 255, 0.2);
            }
            h1 { margin-top: 0; font-size: 24px; }
            p { opacity: 0.9; line-height: 1.6; }
            .retry-btn {
              background: #3B82F6;
              color: white;
              border: none;
              padding: 12px 24px;
              border-radius: 8px;
              font-size: 16px;
              cursor: pointer;
              margin: 10px;
              transition: all 0.2s ease;
            }
            .retry-btn:hover { background: #2563EB; transform: translateY(-1px); }
            .back-btn {
              background: rgba(255, 255, 255, 0.2);
              color: white;
              border: 1px solid rgba(255, 255, 255, 0.3);
              padding: 12px 24px;
              border-radius: 8px;
              text-decoration: none;
              display: inline-block;
              margin: 10px;
              transition: all 0.2s ease;
            }
            .back-btn:hover { background: rgba(255, 255, 255, 0.3); transform: translateY(-1px); }
          </style>
        </head>
        <body>
          <div class="container">
            <h1>🛡️ Safety Standards</h1>
            <p>Unable to load safety standards at the moment. Please check your internet connection and try again.</p>
            <p style="font-size: 14px; opacity: 0.7;">Error: ${error.message}</p>
            <div>
              <button class="retry-btn" onclick="window.location.reload()">Try Again</button>
              <a href="/" class="back-btn">← Back to Calculator</a>
            </div>
          </div>
        </body>
        </html>
      `;
      
      res.status(503).send(offlineHTML);
    }
  });