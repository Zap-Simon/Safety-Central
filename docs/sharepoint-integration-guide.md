# SharePoint Integration Guide

This guide documents how to integrate SharePoint document libraries into your application tabs, based on the working implementations for Health & Safety Policy documents and Equipment Test & Tag records.

## Overview

The SharePoint integration uses Microsoft Graph API to fetch documents from SharePoint sites and display them in a user-friendly interface with preview and download capabilities.

## Architecture

### Backend Components
- **API Endpoints**: Express routes that authenticate and fetch documents from SharePoint
- **Authentication**: Uses Azure AD access tokens passed from the frontend
- **Folder Discovery**: Dynamically explores SharePoint folder structures
- **Error Handling**: Includes retry logic for temporary SharePoint issues

### Frontend Components
- **Document Display**: Card-based UI showing document details
- **Authentication**: Uses MSAL for Azure AD authentication
- **Actions**: Preview (opens in SharePoint) and download functionality

## Working Examples

### 1. Health & Safety Policy Documents

**SharePoint Site**: `cranfieldglass.sharepoint.com/sites/HealthSafetyAdministration`

**Folder Structure**:
```
Root/
└── Health & Safety Policy/
    ├── H&S Policy Cranfield Glass - Review 31-03-2025.pdf
    └── H&S Policy Cranfield Glass.docx
```

**API Endpoint**: `/api/policy-documents`

**Implementation Notes**:
- Folder is directly in the root drive
- Explores root to find "Health & Safety Policy" folder
- Works with both PDF and Word documents

### 2. Environment Policy Documents

**SharePoint Site**: `cranfieldglass.sharepoint.com/sites/EnvironmentAdministration`

**Folder Structure**:
```
Root/
└── Environment Policy/
    ├── Environment Policy Document.pdf
    └── Environmental Management Procedures.docx
```

**API Endpoint**: `/api/environment-policy-documents`

**Implementation Notes**:
- Folder is directly in the root drive
- Explores root to find folders containing "environment" keyword
- Works with both PDF and Word documents
- Uses green color scheme in the UI (green-600/green-700 gradients)

### 3. Quality Policy Documents

**SharePoint Site**: `cranfieldglass.sharepoint.com/sites/QualityAdministration`

**Folder Structure**:
```
Root/
└── Quality Policy/
    ├── Quality Management System.pdf
    └── Quality Standards Document.docx
```

**API Endpoint**: `/api/quality-policy-documents`

**Implementation Notes**:
- Folder is directly in the root drive
- Explores root to find folders containing "quality" keyword
- Works with both PDF and Word documents
- Uses blue color scheme in the UI (blue-600/blue-700 gradients)

### 4. Equipment Test & Tag Records

**SharePoint Site**: `cranfieldglass.sharepoint.com/sites/EquipmentMaintenance`

**Folder Structure**:
```
Root/
└── Equipment Test - Tag/
    └── CRANFIELD GLASS 11TH SEPT.pdf
```

**API Endpoint**: `/api/equipment-test-tag-documents`

**Implementation Notes**:
- Folder is directly in the root drive
- Uses flexible folder name matching for "test" and "tag" keywords
- Includes retry logic for temporary SharePoint issues

## Step-by-Step Implementation Guide

### Step 1: Create Backend API Endpoint

Add a new endpoint in `server/routes.ts`:

```javascript
app.get('/api/your-new-documents', async (req, res) => {
  try {
    // 1. Authentication check
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({
        success: false,
        error: 'Authentication required'
      });
    }

    const accessToken = authHeader.substring(7);
    
    // 2. Test basic Graph API access
    const testResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json'
      }
    });
    
    if (!testResponse.ok) {
      throw new Error(`Graph API authentication failed: ${testResponse.status}`);
    }

    // 3. Get SharePoint site
    const siteUrl = "cranfieldglass.sharepoint.com:/sites/YourSiteName:";
    const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteUrl}`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json'
      }
    });

    if (!siteResponse.ok) {
      const siteErrorText = await siteResponse.text();
      throw new Error(`Site lookup error: ${siteResponse.status} - ${siteErrorText}`);
    }

    const siteData = await siteResponse.json();
    
    // 4. Explore folder structure (see detailed folder discovery logic below)
    // ... folder discovery implementation
    
    // 5. Fetch files with retry logic (see retry implementation below)
    // ... file fetching with retry logic
    
    res.json({
      success: true,
      documents
    });

  } catch (error) {
    console.error('Error fetching documents:', error);
    res.status(500).json({
      success: false,
      error: error instanceof Error ? error.message : 'Failed to fetch documents'
    });
  }
});
```

### Step 2: Folder Discovery Logic

Choose the appropriate folder discovery pattern based on your SharePoint structure:

#### Pattern A: Direct Root Folder (like Equipment Test & Tag)

```javascript
// Explore root to find your folder directly
const rootResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root/children?$select=name,folder,webUrl`, {
  headers: {
    'Authorization': `Bearer ${accessToken}`,
    'Accept': 'application/json'
  }
});

const rootData = await rootResponse.json();

// Look for your target folder directly in root
const targetFolder = rootData.value.find((item: any) => 
  item.folder && item.name.toLowerCase().includes('your-folder-keyword')
);

if (targetFolder) {
  folderPath = `/${targetFolder.name}`;
}
```

#### Pattern B: Nested in Shared Documents (if needed)

```javascript
// First find Shared Documents, then look inside it
const sharedDocsFolder = rootData.value.find((item: any) => 
  item.folder && item.name.toLowerCase().includes('shared documents')
);

if (sharedDocsFolder) {
  // Explore inside Shared Documents
  const sharedDocsResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root:/${encodeURIComponent(sharedDocsFolder.name)}:/children?$select=name,folder,webUrl`, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json'
    }
  });
  
  const sharedDocsData = await sharedDocsResponse.json();
  
  const targetFolder = sharedDocsData.value.find((item: any) => 
    item.folder && item.name.toLowerCase().includes('your-folder-keyword')
  );
  
  if (targetFolder) {
    folderPath = `/${sharedDocsFolder.name}/${targetFolder.name}`;
  }
}
```

### Step 3: File Fetching with Retry Logic

```javascript
// Retry logic for temporary SharePoint issues
let response;
let attempts = 0;
const maxAttempts = 3;

while (attempts < maxAttempts) {
  attempts++;
  
  response = await fetch(graphApiUrl, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json'
    }
  });

  if (response.ok) {
    break;
  }
  
  // If it's a 503 (Service Unavailable) or 429 (Too Many Requests), wait and retry
  if ((response.status === 503 || response.status === 429) && attempts < maxAttempts) {
    console.log(`SharePoint temporarily unavailable (${response.status}), retrying in ${attempts * 2} seconds...`);
    await new Promise(resolve => setTimeout(resolve, attempts * 2000));
    continue;
  }
  
  break;
}

if (!response.ok) {
  const errorText = await response.text();
  
  // Provide user-friendly message for temporary SharePoint issues
  if (response.status === 503) {
    throw new Error('SharePoint is temporarily unavailable. Please try again in a few minutes.');
  } else if (response.status === 429) {
    throw new Error('SharePoint request limit reached. Please try again shortly.');
  }
  
  throw new Error(`SharePoint API error: ${response.status} - ${errorText}`);
}
```

### Step 4: Create Frontend Integration

Add to your page component:

```typescript
// 1. Add interface for document type
interface YourDocument {
  name: string;
  url: string;
  type: 'word' | 'pdf';
  size: string;
  modified: string;
}

// 2. Add query for fetching documents
const { data: documents, isLoading, error } = useQuery({
  queryKey: ['/api/your-new-documents'],
  queryFn: async () => {
    try {
      const token = await authService.getAccessToken();
      const response = await fetch('/api/your-new-documents', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json'
        },
        credentials: 'include'
      });
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to fetch documents: ${response.status} - ${errorText}`);
      }
      
      return await response.json();
    } catch (error) {
      console.error('Error fetching documents:', error);
      throw error;
    }
  },
  enabled: true,
  retry: false
});

// 3. Add handlers for preview and download
const handleDocumentPreview = (document: YourDocument) => {
  window.open(document.url, '_blank');
};

const handleDocumentDownload = async (document: YourDocument) => {
  try {
    const token = await authService.getAccessToken();
    const downloadUrl = `/api/policy-document-proxy?url=${encodeURIComponent(document.url)}&access_token=${encodeURIComponent(token)}&download=true`;
    
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = document.name;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  } catch (error) {
    console.error('Failed to download document:', error);
  }
};
```

### Step 5: Add UI Components

Use the card-based UI pattern:

```jsx
{isLoading ? (
  <div className="flex items-center justify-center py-8">
    <div className="text-gray-500">Loading documents...</div>
  </div>
) : documents && documents.documents?.length > 0 ? (
  <div className="grid gap-4 md:grid-cols-2">
    {documents.documents.map((document: YourDocument, index: number) => (
      <div
        key={index}
        className="border border-gray-200 rounded-lg p-4 hover:border-gray-300 transition-colors flex flex-col"
        data-testid={`card-document-${index}`}
      >
        <div className="flex items-start justify-between mb-3">
          <div className="flex items-center space-x-3 flex-1 min-w-0">
            {document.type === 'word' ? (
              <FileText className="text-blue-600 flex-shrink-0" size={20} />
            ) : (
              <File className="text-red-600 flex-shrink-0" size={20} />
            )}
            <div className="min-w-0 flex-1">
              <h3 className="font-medium text-gray-900 truncate">{document.name}</h3>
              <div className="flex items-center space-x-2 mt-1">
                <Badge variant="secondary" className="text-xs">
                  {document.type.toUpperCase()}
                </Badge>
                <span className="text-xs text-gray-500">{document.size}</span>
              </div>
            </div>
          </div>
        </div>
        
        <div className="text-xs text-gray-500 mb-3">
          Last modified: {document.modified}
        </div>
        
        <div className="flex space-x-2 mt-auto">
          <Button 
            variant="outline" 
            size="sm" 
            onClick={() => handleDocumentPreview(document)}
            className="flex-1"
          >
            <ExternalLink size={14} className="mr-1" />
            Preview
          </Button>
          <Button 
            variant="outline" 
            size="sm" 
            onClick={() => handleDocumentDownload(document)}
            className="flex-1"
          >
            <Download size={14} className="mr-1" />
            Download
          </Button>
        </div>
      </div>
    ))}
  </div>
) : error ? (
  <div className="text-center py-8">
    <div className="text-red-600 mb-2">Failed to load documents</div>
    <div className="text-sm text-gray-500">Please check your SharePoint access or try again later</div>
  </div>
) : (
  <div className="text-center py-8">
    <div className="text-gray-500 mb-2">No documents found</div>
    <div className="text-sm text-gray-400">Documents will appear here once available in SharePoint</div>
  </div>
)}
```

## Common Patterns and Best Practices

### Error Handling
- Always include retry logic for 503 and 429 errors
- Provide user-friendly error messages
- Log detailed errors for debugging

### Folder Discovery
- Use flexible keyword matching for folder names
- Handle both direct root and nested folder structures
- Always explore the actual folder structure rather than assuming paths

### Authentication
- Validate tokens before making SharePoint requests
- Handle token refresh automatically via MSAL
- Use the existing policy-document-proxy for downloads

### Performance
- Use TanStack Query for caching and loading states
- Implement proper loading and error states
- Add data-testid attributes for testing

## Troubleshooting

### Common Issues

1. **403 Forbidden**: Check SharePoint site permissions
2. **404 Not Found**: Verify site URL and folder paths
3. **503 Service Unavailable**: Temporary SharePoint issue - retry logic handles this
4. **401 Unauthorized**: Token expired or invalid - MSAL should refresh automatically

### Debugging Steps

1. Check the browser console for detailed error messages
2. Verify the SharePoint site URL is correct
3. Test the site access in a browser while logged in
4. Check the server logs for detailed Graph API responses

## Example SharePoint Sites to Integrate

Based on the existing pattern, you could add integrations for:

- Safety incident reports
- Equipment manuals
- Training materials
- Quality control documents
- Maintenance schedules
- Compliance certificates

Each would follow the same pattern but with different site URLs and folder structures.