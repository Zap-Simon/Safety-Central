import express from 'express';
import { authenticateToken, authenticateTokenFromForm } from './auth-middleware';

const router = express.Router();

/**
 * Example: Protected API endpoint that requires authentication
 * Returns user data from the authenticated request
 */
router.get('/api/user/profile', authenticateToken, async (req, res) => {
  try {
    // req.user is available from the authentication middleware
    const user = req.user;
    
    res.json({
      success: true,
      user: {
        id: user.id,
        name: user.displayName || user.name,
        email: user.mail || user.userPrincipalName,
        jobTitle: user.jobTitle,
        department: user.department
      }
    });
  } catch (error) {
    console.error('Error fetching user profile:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to fetch user profile'
    });
  }
});

/**
 * Example: Protected data endpoint 
 * Fetches data using the authenticated user's token
 */
router.get('/api/data/secure', authenticateToken, async (req, res) => {
  try {
    const accessToken = req.accessToken;
    
    // Example: Use the token to call Microsoft Graph or SharePoint
    const response = await fetch('https://graph.microsoft.com/v1.0/me/drive/items', {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`API call failed: ${response.status}`);
    }

    const data = await response.json();
    
    res.json({
      success: true,
      data: data.value || [],
      message: 'Data retrieved successfully'
    });
    
  } catch (error) {
    console.error('Error fetching secure data:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to fetch data',
      message: error instanceof Error ? error.message : 'Unknown error'
    });
  }
});

/**
 * Example: Protected document/file route
 * Useful for serving protected PDFs, documents, etc.
 */
router.get('/documents/protected/:filename', authenticateTokenFromForm, async (req, res) => {
  try {
    const { filename } = req.params;
    const user = req.user;
    
    // Example: Check if user has access to this document
    // You could check permissions, user roles, etc.
    
    // For demonstration, we'll just serve a response
    // In reality, you'd serve actual files or proxy to SharePoint/OneDrive
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
    
    // Example response - in reality you'd stream the actual file
    res.send(`Protected document access granted for ${user.displayName}`);
    
  } catch (error) {
    console.error('Error serving protected document:', error);
    res.status(500).send('Error serving document');
  }
});

/**
 * Example: POST endpoint for creating data
 * Requires authentication and validates user permissions
 */
router.post('/api/data/create', authenticateToken, async (req, res) => {
  try {
    const user = req.user;
    const { title, description, category } = req.body;
    
    // Validate required fields
    if (!title || !description) {
      return res.status(400).json({
        success: false,
        error: 'Title and description are required'
      });
    }
    
    // Example: Create data with user context
    const newItem = {
      id: Date.now().toString(), // In reality, use proper ID generation
      title,
      description,
      category: category || 'General',
      createdBy: user.displayName || user.name,
      createdAt: new Date().toISOString(),
      userEmail: user.mail || user.userPrincipalName
    };
    
    // Here you would save to your database or external service
    // For this example, we'll just return the created item
    
    res.json({
      success: true,
      data: newItem,
      message: 'Item created successfully'
    });
    
  } catch (error) {
    console.error('Error creating item:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to create item'
    });
  }
});

/**
 * Example: SharePoint-specific endpoint
 * Uses the token to interact with SharePoint lists
 */
router.get('/api/sharepoint/lists/:listName', authenticateToken, async (req, res) => {
  try {
    const { listName } = req.params;
    const accessToken = req.accessToken;
    
    // Example SharePoint site URL - update with your domain
    const sharePointSiteUrl = 'https://your-domain.sharepoint.com';
    const listUrl = `${sharePointSiteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    
    const response = await fetch(listUrl, {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json;odata=verbose'
      }
    });
    
    if (!response.ok) {
      throw new Error(`SharePoint API error: ${response.status}`);
    }
    
    const data = await response.json();
    
    res.json({
      success: true,
      listName,
      items: data.d?.results || [],
      count: data.d?.results?.length || 0
    });
    
  } catch (error) {
    console.error('Error fetching SharePoint list:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to fetch SharePoint list',
      message: error instanceof Error ? error.message : 'Unknown error'
    });
  }
});

/**
 * Example: Bulk operations endpoint
 * Demonstrates handling multiple operations with authentication
 */
router.post('/api/data/bulk-update', authenticateToken, async (req, res) => {
  try {
    const user = req.user;
    const { items } = req.body;
    
    if (!Array.isArray(items) || items.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Items array is required and must not be empty'
      });
    }
    
    const results = [];
    let successCount = 0;
    let errorCount = 0;
    
    // Process each item
    for (const item of items) {
      try {
        // Validate item structure
        if (!item.id || !item.action) {
          results.push({
            id: item.id || 'unknown',
            success: false,
            error: 'Item must have id and action fields'
          });
          errorCount++;
          continue;
        }
        
        // Simulate processing (replace with actual logic)
        // Here you would update your database or call external APIs
        const processedItem = {
          ...item,
          updatedBy: user.displayName || user.name,
          updatedAt: new Date().toISOString()
        };
        
        results.push({
          id: item.id,
          success: true,
          data: processedItem
        });
        successCount++;
        
        // Add delay between operations to avoid rate limits
        await new Promise(resolve => setTimeout(resolve, 100));
        
      } catch (itemError) {
        console.error(`Error processing item ${item.id}:`, itemError);
        results.push({
          id: item.id,
          success: false,
          error: itemError instanceof Error ? itemError.message : 'Processing failed'
        });
        errorCount++;
      }
    }
    
    res.json({
      success: errorCount === 0,
      processed: items.length,
      successful: successCount,
      failed: errorCount,
      results
    });
    
  } catch (error) {
    console.error('Error in bulk update:', error);
    res.status(500).json({
      success: false,
      error: 'Bulk update operation failed'
    });
  }
});

export default router;