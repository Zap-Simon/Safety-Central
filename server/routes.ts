import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import { SharePointListsService } from "./sharepoint-lists-service";
import { generateMeetingWordDoc } from "./word-generator";
import { OpenAIService } from "./openai-service";
import { MarkdownMeetingGenerator } from "./markdown-generator";
import { db } from "./db";
import { 
  cardOrdering, 
  insertCardOrderingSchema,
  insertTrainingClassificationSchema,
  insertTrainingModuleSchema,
  insertModulePrerequisiteSchema,
  insertModuleToolSchema,
  insertModuleMaterialSchema,
  insertStaffModuleProgressSchema,
  COMPETENCY_LEVELS,
  isAbleToUse
} from "@shared/schema";
import { eq, desc, sql } from "drizzle-orm";
import { getExcelData, listExcelFiles } from "./sharepoint-excel-service.js";
import { PAGEDJS_BASE64 } from "./assets/pagedjs";

// Decode the bundled Paged.js polyfill once. Inlined into the meeting-minutes
// export so page-number footers render reliably and offline in any browser.
// Escape any literal "</script>" so it can't prematurely close the inline tag.
const PAGEDJS_SCRIPT = Buffer.from(PAGEDJS_BASE64, "base64")
  .toString("utf8")
  .replace(/<\/script>/gi, "<\\/script>");

export async function registerRoutes(app: Express): Promise<Server> {


  // SECURITY: Remove CORS OPTIONS handler - not needed for same-origin requests

  // Get policy documents from SharePoint folder
  app.get('/api/policy-documents', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      
      // First, test basic Graph API access
      console.log('Testing basic Graph API access...');
      const testResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });
      
      if (!testResponse.ok) {
        console.error('Basic Graph API test failed:', testResponse.status);
        throw new Error(`Graph API authentication failed: ${testResponse.status}`);
      } else {
        console.log('Basic Graph API test passed');
      }

      // Fetch files from the Health & Safety Policy folder using Microsoft Graph API
      // First get the site ID using the correct format
      const siteUrl = "cranfieldglass.sharepoint.com:/sites/HealthSafetyAdministration:";
      console.log('Getting site information for:', siteUrl);
      
      const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteUrl}`, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      if (!siteResponse.ok) {
        const siteErrorText = await siteResponse.text();
        console.error('Site lookup error:', siteErrorText);
        throw new Error(`Site lookup error: ${siteResponse.status} - ${siteErrorText}`);
      }

      const siteData = await siteResponse.json();
      console.log('Site data retrieved successfully, site ID:', siteData.id);
      
      // First, let's explore the drive structure to find the correct folder
      console.log('Exploring drive root to find folders...');
      const rootUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root/children?$select=name,folder,webUrl`;
      
      const rootResponse = await fetch(rootUrl, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      if (!rootResponse.ok) {
        const rootErrorText = await rootResponse.text();
        console.error('Root exploration error:', rootErrorText);
        throw new Error(`Root exploration error: ${rootResponse.status} - ${rootErrorText}`);
      }

      const rootData = await rootResponse.json();
      console.log('Root folder contents:', rootData.value.map((item: any) => ({ name: item.name, isFolder: !!item.folder })));
      
      // Look for Health & Safety Policy folder directly in root
      const policyFolder = rootData.value.find((item: any) => 
        item.folder && item.name.toLowerCase().includes('health') && item.name.toLowerCase().includes('safety')
      );
      
      if (!policyFolder) {
        throw new Error('Could not find Health & Safety Policy folder. Available folders: ' + 
          rootData.value.filter((item: any) => item.folder).map((item: any) => item.name).join(', '));
      }
      
      console.log('Found policy folder directly in root:', policyFolder.name);
      
      // Now fetch files from the policy folder (it's directly in root)
      const folderPath = `/${policyFolder.name}`;
      const graphApiUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root:${encodeURIComponent(folderPath)}:/children?$select=name,size,lastModifiedDateTime,webUrl,file&$orderby=lastModifiedDateTime desc`;
      
      console.log('Calling Graph API URL:', graphApiUrl);
      
      const response = await fetch(graphApiUrl, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      console.log('Graph API response status:', response.status);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error('Graph API error details:', errorText);
        throw new Error(`SharePoint API error: ${response.status} - ${errorText}`);
      }

      const data = await response.json();
      const files = data.value || [];
      
      console.log('Total files found:', files.length);
      console.log('Files details:', files.map((f: any) => ({ name: f.name, hasFile: !!f.file, size: f.size })));
      
      // Process files and identify types - only process actual files, not folders
      const documents = files
        .filter((file: any) => file.file) // Only include files, not folders
        .map((file: any) => {
          const name = file.name;
          const extension = name.split('.').pop()?.toLowerCase();
          const type = extension === 'docx' || extension === 'doc' ? 'word' : 
                      extension === 'pdf' ? 'pdf' : 'unknown';
          
          // Convert Graph API date format
          const modified = new Date(file.lastModifiedDateTime).toLocaleDateString();
          
          // Format file size
          const sizeInKB = Math.round(file.size / 1024);
          const size = sizeInKB > 1024 ? `${Math.round(sizeInKB / 1024)} MB` : `${sizeInKB} KB`;
          
          // Use the webUrl provided by Graph API
          const documentUrl = file.webUrl;
          
          return {
            name,
            url: documentUrl,
            type,
            size,
            modified
          };
        })
        .filter((doc: any) => doc.type !== 'unknown'); // Only include Word and PDF files
      
      console.log('Final documents to return:', documents.length);
      console.log('Document names:', documents.map((d: any) => d.name));

      res.json({
        success: true,
        documents
      });

    } catch (error) {
      console.error('Error fetching policy documents:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch policy documents'
      });
    }
  });

  // Get environment policy documents from SharePoint folder
  app.get('/api/environment-policy-documents', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      
      // First, test basic Graph API access
      console.log('Testing basic Graph API access for environment policy...');
      const testResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });
      
      if (!testResponse.ok) {
        console.error('Basic Graph API test failed:', testResponse.status);
        throw new Error(`Graph API authentication failed: ${testResponse.status}`);
      } else {
        console.log('Basic Graph API test passed');
      }

      // Fetch files from the Environment Policy folder using Microsoft Graph API
      // First get the site ID using the correct format
      const siteUrl = "cranfieldglass.sharepoint.com:/sites/CranfieldGlass-EnvironmentalManagementSystemEMS:";
      console.log('Getting site information for:', siteUrl);
      
      const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteUrl}`, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      if (!siteResponse.ok) {
        const siteErrorText = await siteResponse.text();
        console.error('Site lookup error:', siteErrorText);
        throw new Error(`Site lookup error: ${siteResponse.status} - ${siteErrorText}`);
      }

      const siteData = await siteResponse.json();
      console.log('Site data retrieved successfully, site ID:', siteData.id);
      
      // First, let's explore the drive structure to find the correct folder
      console.log('Exploring drive root to find folders...');
      const rootUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root/children?$select=name,folder,webUrl`;
      
      const rootResponse = await fetch(rootUrl, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      if (!rootResponse.ok) {
        const rootErrorText = await rootResponse.text();
        console.error('Root exploration error:', rootErrorText);
        throw new Error(`Root exploration error: ${rootResponse.status} - ${rootErrorText}`);
      }

      const rootData = await rootResponse.json();
      console.log('Root folder contents:', rootData.value.map((item: any) => ({ name: item.name, isFolder: !!item.folder })));
      
      // Look for Current Policy folder directly in root
      const policyFolder = rootData.value.find((item: any) => 
        item.folder && item.name.toLowerCase().includes('current policy')
      );
      
      if (!policyFolder) {
        throw new Error('Could not find Current Policy folder. Available folders: ' + 
          rootData.value.filter((item: any) => item.folder).map((item: any) => item.name).join(', '));
      }
      
      console.log('Found environment policy folder (Current Policy) in root:', policyFolder.name);
      
      // Now fetch files from the policy folder (it's directly in root)
      const folderPath = `/${policyFolder.name}`;
      const graphApiUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root:${encodeURIComponent(folderPath)}:/children?$select=name,size,lastModifiedDateTime,webUrl,file&$orderby=lastModifiedDateTime desc`;
      
      console.log('Calling Graph API URL for environment policy:', graphApiUrl);
      
      const response = await fetch(graphApiUrl, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      console.log('Graph API response status:', response.status);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error('Graph API error details:', errorText);
        throw new Error(`SharePoint API error: ${response.status} - ${errorText}`);
      }

      const data = await response.json();
      const files = data.value || [];
      
      console.log('Total environment policy files found:', files.length);
      console.log('Files details:', files.map((f: any) => ({ name: f.name, hasFile: !!f.file, size: f.size })));
      
      // Process files and identify types - only process actual files, not folders
      const documents = files
        .filter((file: any) => file.file) // Only include files, not folders
        .map((file: any) => {
          const name = file.name;
          const extension = name.split('.').pop()?.toLowerCase();
          const type = extension === 'docx' || extension === 'doc' ? 'word' : 
                      extension === 'pdf' ? 'pdf' : 'unknown';
          
          // Convert Graph API date format
          const modified = new Date(file.lastModifiedDateTime).toLocaleDateString();
          
          // Format file size
          const sizeInKB = Math.round(file.size / 1024);
          const size = sizeInKB > 1024 ? `${Math.round(sizeInKB / 1024)} MB` : `${sizeInKB} KB`;
          
          // Use the webUrl provided by Graph API
          const documentUrl = file.webUrl;
          
          return {
            name,
            url: documentUrl,
            type,
            size,
            modified
          };
        })
        .filter((doc: any) => doc.type !== 'unknown'); // Only include Word and PDF files
      
      console.log('Final environment policy documents to return:', documents.length);
      console.log('Document names:', documents.map((d: any) => d.name));

      res.json({
        success: true,
        documents
      });

    } catch (error) {
      console.error('Error fetching environment policy documents:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch environment policy documents'
      });
    }
  });

  // Get quality policy documents from SharePoint folder
  app.get('/api/quality-policy-documents', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      
      // First, test basic Graph API access
      console.log('Testing basic Graph API access for quality policy...');
      const testResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });
      
      if (!testResponse.ok) {
        console.error('Basic Graph API test failed:', testResponse.status);
        throw new Error(`Graph API authentication failed: ${testResponse.status}`);
      } else {
        console.log('Basic Graph API test passed');
      }

      // Fetch files from the Quality Policy folder using Microsoft Graph API
      // First get the site ID using the correct format
      const siteUrl = "cranfieldglass.sharepoint.com:/sites/CranfieldGlassQuality:";
      console.log('Getting site information for:', siteUrl);
      
      const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteUrl}`, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      if (!siteResponse.ok) {
        const siteErrorText = await siteResponse.text();
        console.error('Site lookup error:', siteErrorText);
        throw new Error(`Site lookup error: ${siteResponse.status} - ${siteErrorText}`);
      }

      const siteData = await siteResponse.json();
      console.log('Site data retrieved successfully, site ID:', siteData.id);
      
      // First, let's explore the drive structure to find the correct folder
      console.log('Exploring drive root to find folders...');
      const rootUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root/children?$select=name,folder,webUrl`;
      
      const rootResponse = await fetch(rootUrl, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      if (!rootResponse.ok) {
        const rootErrorText = await rootResponse.text();
        console.error('Root exploration error:', rootErrorText);
        throw new Error(`Root exploration error: ${rootResponse.status} - ${rootErrorText}`);
      }

      const rootData = await rootResponse.json();
      console.log('Root folder contents:', rootData.value.map((item: any) => ({ name: item.name, isFolder: !!item.folder })));
      
      // Look for Current Policy folder directly in root
      const policyFolder = rootData.value.find((item: any) => 
        item.folder && item.name.toLowerCase().includes('current policy')
      );
      
      if (!policyFolder) {
        throw new Error('Could not find Current Policy folder. Available folders: ' + 
          rootData.value.filter((item: any) => item.folder).map((item: any) => item.name).join(', '));
      }
      
      console.log('Found quality policy folder (Current Policy) in root:', policyFolder.name);
      
      // Now fetch files from the policy folder (it's directly in root)
      const folderPath = `/${policyFolder.name}`;
      const graphApiUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root:${encodeURIComponent(folderPath)}:/children?$select=name,size,lastModifiedDateTime,webUrl,file&$orderby=lastModifiedDateTime desc`;
      
      console.log('Calling Graph API URL for quality policy:', graphApiUrl);
      
      const response = await fetch(graphApiUrl, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      console.log('Graph API response status:', response.status);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error('Graph API error details:', errorText);
        throw new Error(`SharePoint API error: ${response.status} - ${errorText}`);
      }

      const data = await response.json();
      const files = data.value || [];
      
      console.log('Total quality policy files found:', files.length);
      console.log('Files details:', files.map((f: any) => ({ name: f.name, hasFile: !!f.file, size: f.size })));
      
      // Process files and identify types - only process actual files, not folders
      const documents = files
        .filter((file: any) => file.file) // Only include files, not folders
        .map((file: any) => {
          const name = file.name;
          const extension = name.split('.').pop()?.toLowerCase();
          const type = extension === 'docx' || extension === 'doc' ? 'word' : 
                      extension === 'pdf' ? 'pdf' : 'unknown';
          
          // Convert Graph API date format
          const modified = new Date(file.lastModifiedDateTime).toLocaleDateString();
          
          // Format file size
          const sizeInKB = Math.round(file.size / 1024);
          const size = sizeInKB > 1024 ? `${Math.round(sizeInKB / 1024)} MB` : `${sizeInKB} KB`;
          
          // Use the webUrl provided by Graph API
          const documentUrl = file.webUrl;
          
          return {
            name,
            url: documentUrl,
            type,
            size,
            modified
          };
        })
        .filter((doc: any) => doc.type !== 'unknown'); // Only include Word and PDF files
      
      console.log('Final quality policy documents to return:', documents.length);
      console.log('Document names:', documents.map((d: any) => d.name));

      res.json({
        success: true,
        documents
      });

    } catch (error) {
      console.error('Error fetching quality policy documents:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch quality policy documents'
      });
    }
  });

  // Get equipment test & tag documents from SharePoint folder
  app.get('/api/equipment-test-tag-documents', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      
      // First, test basic Graph API access
      console.log('Testing basic Graph API access for equipment test & tag...');
      const testResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });
      
      if (!testResponse.ok) {
        console.error('Basic Graph API test failed:', testResponse.status);
        throw new Error(`Graph API authentication failed: ${testResponse.status}`);
      } else {
        console.log('Basic Graph API test passed');
      }

      // Fetch files from the Equipment Maintenance site
      const siteUrl = "cranfieldglass.sharepoint.com:/sites/EquipmentMaintenance:";
      console.log('Getting site information for:', siteUrl);
      
      const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteUrl}`, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      if (!siteResponse.ok) {
        const siteErrorText = await siteResponse.text();
        console.error('Site lookup error:', siteErrorText);
        throw new Error(`Site lookup error: ${siteResponse.status} - ${siteErrorText}`);
      }

      const siteData = await siteResponse.json();
      console.log('Site data retrieved successfully, site ID:', siteData.id);
      
      // First, let's explore the drive structure to find the correct folder
      console.log('Exploring drive root to find folders...');
      const rootUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root/children?$select=name,folder,webUrl`;
      
      const rootResponse = await fetch(rootUrl, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      if (!rootResponse.ok) {
        const rootErrorText = await rootResponse.text();
        console.error('Root exploration error:', rootErrorText);
        throw new Error(`Root exploration error: ${rootResponse.status} - ${rootErrorText}`);
      }

      const rootData = await rootResponse.json();
      console.log('Root folder contents:', rootData.value.map((item: any) => ({ name: item.name, isFolder: !!item.folder })));
      
      // First check if Equipment Test - Tag folder is directly in root
      const directTestTagFolder = rootData.value.find((item: any) => 
        item.folder && (
          item.name.toLowerCase().includes('equipment test') || 
          (item.name.toLowerCase().includes('test') && item.name.toLowerCase().includes('tag'))
        )
      );
      
      let folderPath: string;
      let testTagFolderName: string;
      
      if (directTestTagFolder) {
        console.log('Found test & tag folder directly in root:', directTestTagFolder.name);
        folderPath = `/${directTestTagFolder.name}`;
        testTagFolderName = directTestTagFolder.name;
      } else {
        // Look for Shared Documents folder
        const sharedDocsFolder = rootData.value.find((item: any) => 
          item.folder && item.name.toLowerCase().includes('shared documents')
        );
        
        if (!sharedDocsFolder) {
          console.log('Available folders:', rootData.value.filter((item: any) => item.folder).map((item: any) => item.name));
          throw new Error('Could not find Equipment Test - Tag folder directly or in Shared Documents. Available folders: ' + 
            rootData.value.filter((item: any) => item.folder).map((item: any) => item.name).join(', '));
        }
        
        console.log('Found Shared Documents folder:', sharedDocsFolder.name);
        
        // Now explore Shared Documents to find the Equipment Test - Tag folder
        const sharedDocsUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root:/${encodeURIComponent(sharedDocsFolder.name)}:/children?$select=name,folder,webUrl`;
        
        const sharedDocsResponse = await fetch(sharedDocsUrl, {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Accept': 'application/json'
          }
        });

        if (!sharedDocsResponse.ok) {
          const sharedDocsErrorText = await sharedDocsResponse.text();
          console.error('Shared Documents exploration error:', sharedDocsErrorText);
          throw new Error(`Shared Documents exploration error: ${sharedDocsResponse.status} - ${sharedDocsErrorText}`);
        }

        const sharedDocsData = await sharedDocsResponse.json();
        console.log('Shared Documents contents:', sharedDocsData.value.map((item: any) => ({ name: item.name, isFolder: !!item.folder })));
        
        // Look for Equipment Test - Tag folder
        const testTagFolder = sharedDocsData.value.find((item: any) => 
          item.folder && (
            item.name.toLowerCase().includes('equipment test') || 
            (item.name.toLowerCase().includes('test') && item.name.toLowerCase().includes('tag'))
          )
        );
        
        if (!testTagFolder) {
          console.log('Available folders in Shared Documents:', sharedDocsData.value.filter((item: any) => item.folder).map((item: any) => item.name));
          throw new Error('Could not find Equipment Test - Tag folder. Available folders in Shared Documents: ' + 
            sharedDocsData.value.filter((item: any) => item.folder).map((item: any) => item.name).join(', '));
        }
        
        console.log('Found test & tag folder in Shared Documents:', testTagFolder.name);
        folderPath = `/${sharedDocsFolder.name}/${testTagFolder.name}`;
        testTagFolderName = testTagFolder.name;
      }
      
      // Now fetch files from the test & tag folder
      const graphApiUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root:${encodeURIComponent(folderPath)}:/children?$select=name,size,lastModifiedDateTime,webUrl,file&$orderby=lastModifiedDateTime desc`;
      
      console.log('Calling Graph API URL for test & tag documents:', graphApiUrl);
      
      // Retry logic for temporary SharePoint issues
      let response: Response | undefined;
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

        console.log(`Graph API response status (attempt ${attempts}):`, response.status);
        
        if (response.ok) {
          break;
        }
        
        // If it's a 503 (Service Unavailable) or 429 (Too Many Requests), wait and retry
        if ((response.status === 503 || response.status === 429) && attempts < maxAttempts) {
          console.log(`SharePoint temporarily unavailable (${response.status}), retrying in ${attempts * 2} seconds...`);
          await new Promise(resolve => setTimeout(resolve, attempts * 2000));
          continue;
        }
        
        // For other errors or final attempt, break and handle error
        break;
      }
      
      if (!response) {
        throw new Error('SharePoint did not respond after multiple attempts. Please try again shortly.');
      }

      if (!response.ok) {
        const errorText = await response.text();
        console.error('Graph API error details:', errorText);
        
        // Provide user-friendly message for temporary SharePoint issues
        if (response.status === 503) {
          throw new Error('SharePoint is temporarily unavailable. Please try again in a few minutes.');
        } else if (response.status === 429) {
          throw new Error('SharePoint request limit reached. Please try again shortly.');
        }
        
        throw new Error(`SharePoint API error: ${response.status} - ${errorText}`);
      }

      const data = await response.json();
      const files = data.value || [];
      
      console.log('Total test & tag files found:', files.length);
      console.log('Files details:', files.map((f: any) => ({ name: f.name, hasFile: !!f.file, size: f.size })));
      
      // Process files and identify types - only process actual files, not folders
      const documents = files
        .filter((file: any) => file.file) // Only include files, not folders
        .map((file: any) => {
          const name = file.name;
          const extension = name.split('.').pop()?.toLowerCase();
          const type = extension === 'docx' || extension === 'doc' ? 'word' : 
                      extension === 'pdf' ? 'pdf' : 'unknown';
          
          // Convert Graph API date format
          const modified = new Date(file.lastModifiedDateTime).toLocaleDateString();
          
          // Format file size
          const sizeInKB = Math.round(file.size / 1024);
          const size = sizeInKB > 1024 ? `${Math.round(sizeInKB / 1024)} MB` : `${sizeInKB} KB`;
          
          // Use the webUrl provided by Graph API
          const documentUrl = file.webUrl;
          
          return {
            name,
            url: documentUrl,
            type,
            size,
            modified
          };
        })
        .filter((doc: any) => doc.type !== 'unknown'); // Only include Word and PDF files
      
      console.log('Final test & tag documents to return:', documents.length);
      console.log('Document names:', documents.map((d: any) => d.name));

      res.json({
        success: true,
        documents
      });

    } catch (error) {
      console.error('Error fetching equipment test & tag documents:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch equipment test & tag documents'
      });
    }
  });

  // List Excel files in SharePoint folder
  app.get('/api/sharepoint-excel/files', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { site, folder } = req.query;

      if (!site) {
        return res.status(400).json({
          success: false,
          error: 'site parameter is required'
        });
      }

      const siteUrl = `cranfieldglass.sharepoint.com:/sites/${site}:`;
      const folderPath = folder ? String(folder) : '';

      const excelFiles = await listExcelFiles(accessToken, siteUrl, folderPath);

      res.json({
        success: true,
        files: excelFiles
      });

    } catch (error) {
      console.error('Error listing Excel files:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to list Excel files'
      });
    }
  });

  // Get Excel worksheet data
  app.get('/api/sharepoint-excel/data', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { site, folder, fileName, worksheet, refresh } = req.query;

      if (!site || !fileName) {
        return res.status(400).json({
          success: false,
          error: 'site and fileName parameters are required'
        });
      }

      const siteUrl = `cranfieldglass.sharepoint.com:/sites/${site}:`;
      const folderPath = folder ? String(folder) : '';
      const fileNameStr = String(fileName);
      const worksheetName = worksheet ? String(worksheet) : undefined;

      // If refresh parameter is provided, clear the cache first
      if (refresh) {
        const { clearExcelCache } = await import("./sharepoint-excel-service.js");
        clearExcelCache(siteUrl, folderPath, fileNameStr, worksheetName);
        console.log(`Cache cleared for: ${fileNameStr}`);
      }

      const workbookData = await getExcelData(
        accessToken,
        siteUrl,
        folderPath,
        fileNameStr,
        worksheetName
      );

      res.json({
        success: true,
        data: workbookData
      });

    } catch (error) {
      console.error('Error fetching Excel data:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch Excel data'
      });
    }
  });

  // Proxy for opening policy documents with authentication
  app.get('/api/policy-document-proxy', async (req, res) => {
    try {
      const { url, access_token, download } = req.query;
      
      if (!access_token) {
        return res.status(401).send(`
          <!DOCTYPE html>
          <html>
          <head>
            <title>Authentication Required - Cranfield Glass</title>
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <style>
              body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 0; padding: 40px; background: #f8fafc; }
              .container { max-width: 500px; margin: 0 auto; background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; }
              .icon { width: 60px; height: 60px; background: #dc2626; border-radius: 50%; margin: 0 auto 20px; display: flex; align-items: center; justify-content: center; color: white; font-size: 24px; }
              h1 { color: #1f2937; margin: 0 0 16px; font-size: 24px; }
              p { color: #6b7280; margin: 0 0 24px; line-height: 1.6; }
              .btn { background: #2563eb; color: white; padding: 12px 24px; border: none; border-radius: 8px; text-decoration: none; display: inline-block; font-weight: 500; }
            </style>
          </head>
          <body>
            <div class="container">
              <div class="icon">🛡️</div>
              <h1>Authentication Required</h1>
              <p>Please sign in with Microsoft 365 to access the policy document.</p>
              <a href="/policy/health-safety" class="btn">Return to Policy Page</a>
            </div>
          </body>
          </html>
        `);
      }

      // Validate the token
      const validationResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${access_token}`,
          'Content-Type': 'application/json'
        }
      });

      if (!validationResponse.ok) {
        return res.status(401).send(`
          <!DOCTYPE html>
          <html>
          <head><title>Invalid Authentication - Cranfield Glass</title></head>
          <body><h1>Authentication Expired</h1><p>Please sign in again.</p></body>
          </html>
        `);
      }

      // Instead of using webUrl, use Microsoft Graph API to get the actual file content
      console.log('Getting file content from Graph API for URL:', url);
      
      // Extract file path from the SharePoint URL to construct Graph API URL
      const urlObj = new URL(url as string);
      const pathParts = urlObj.pathname.split('/');
      const siteIndex = pathParts.indexOf('sites');
      const siteName = pathParts[siteIndex + 1];
      
      // Get site ID first
      const siteUrl = `cranfieldglass.sharepoint.com:/sites/${siteName}:`;
      const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteUrl}`, {
        headers: {
          'Authorization': `Bearer ${access_token}`,
          'Accept': 'application/json'
        }
      });
      
      if (!siteResponse.ok) {
        console.error('Failed to get site info:', siteResponse.status);
        return res.status(404).send('Site not found');
      }
      
      const siteData = await siteResponse.json();
      
      // Extract filename from URL
      const urlFilename = decodeURIComponent(urlObj.pathname.split('/').pop() || '');
      console.log('Looking for file:', urlFilename);
      
      // Use Graph API to search for the file by name
      const searchUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/root/search(q='${encodeURIComponent(urlFilename)}')`;
      const searchResponse = await fetch(searchUrl, {
        headers: {
          'Authorization': `Bearer ${access_token}`,
          'Accept': 'application/json'
        }
      });
      
      if (!searchResponse.ok) {
        console.error('Failed to search for file:', searchResponse.status);
        return res.status(404).send('File search failed');
      }
      
      const searchData = await searchResponse.json();
      const file = searchData.value?.[0];
      
      if (!file) {
        console.error('File not found in search results');
        return res.status(404).send('File not found');
      }
      
      // Get file content using Graph API download URL
      const downloadUrl = `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drive/items/${file.id}/content`;
      const docResponse = await fetch(downloadUrl, {
        headers: {
          'Authorization': `Bearer ${access_token}`,
          'Accept': '*/*'
        }
      });

      if (!docResponse.ok) {
        console.error('Failed to fetch document:', docResponse.status, docResponse.statusText);
        return res.status(404).send(`
          <!DOCTYPE html>
          <html>
          <head><title>Document Not Found - Cranfield Glass</title></head>
          <body><h1>Document Not Found</h1><p>The requested document could not be found.</p></body>
          </html>
        `);
      }

      // Get the filename from the file object
      const filename = file.name || urlFilename || 'document';
      
      // Set appropriate headers based on file type and download parameter
      const contentType = docResponse.headers.get('content-type') || 
                         (filename.endsWith('.pdf') ? 'application/pdf' : 
                          filename.endsWith('.docx') ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' :
                          filename.endsWith('.doc') ? 'application/msword' : 'application/octet-stream');
      
      res.set({
        'Content-Type': contentType,
        'Content-Length': docResponse.headers.get('content-length') || '',
        'Cache-Control': 'no-cache',
        'X-Frame-Options': 'SAMEORIGIN', // Allow embedding in iframes from same origin
      });

      // If download parameter is set, add download headers
      if (download === 'true') {
        res.set({
          'Content-Disposition': `attachment; filename="${filename}"`,
        });
      } else {
        // For preview, use inline disposition
        res.set({
          'Content-Disposition': `inline; filename="${filename}"`,
        });
      }

      // Stream the document content back to the client
      const buffer = await docResponse.arrayBuffer();
      res.send(Buffer.from(buffer));
      
    } catch (error) {
      console.error('Policy document proxy error:', error);
      return res.status(500).send(`
        <!DOCTYPE html>
        <html>
        <head><title>Error - Cranfield Glass</title></head>
        <body><h1>Server Error</h1><p>Please try again later.</p></body>
        </html>
      `);
    }
  });

  // Meeting History endpoint for compliance reporting (requires SPA authentication)
  app.get('/api/meeting-history', async (req, res) => {
    try {
      // Extract Bearer token from Authorization header
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          configured: false,
          authenticated: false,
          message: 'Authentication required. Please sign in with Microsoft 365 to access SharePoint data.',
          data: []
        });
      }

      const accessToken = authHeader.substring(7); // Remove 'Bearer ' prefix
      
      // Use clean SharePoint Lists service with delegated permissions
      const listsService = new SharePointListsService(accessToken);
      
      // Fetch main lists and local action data in parallel
      const [businessIdeas, safetyIdeas, nearMiss, localActionItems] = await Promise.all([
        listsService.getBusinessIdeas(),
        listsService.getSafetyIdeas(),
        listsService.getNearMiss(),
        storage.getAllActionItems()
      ]);
      
      // Create a map of local action items by sharePointItemId for quick lookup
      const actionItemsMap = new Map<string, any>();
      for (const actionItem of localActionItems) {
        actionItemsMap.set(actionItem.sharePointItemId, actionItem);
      }
      
      // Combine all items into single array and merge local action data
      // Use nullish coalescing (??) to allow intentional clears (empty strings)
      const allItems = [...businessIdeas, ...safetyIdeas, ...nearMiss].map(item => {
        const localAction = actionItemsMap.get(item.id);
        if (localAction) {
          return {
            ...item,
            actionPriority: localAction.actionPriority ?? item.actionPriority,
            actionStatus: localAction.actionStatus ?? item.actionStatus,
            actionAssignedTo: localAction.actionAssignedTo ?? item.actionAssignedTo,
            actionStartDate: localAction.actionStartDate ?? item.actionStartDate,
            actionDueDate: localAction.actionDueDate ?? item.actionDueDate,
            actionNotes: localAction.actionNotes ?? item.actionNotes,
            // Local meeting notes override SharePoint (local is always latest save)
            meetingNotes: localAction.meetingNotes !== null && localAction.meetingNotes !== undefined
              ? localAction.meetingNotes
              : item.meetingNotes
          };
        }
        return item;
      });

      res.json({
        configured: true,
        authenticated: true,
        data: allItems
      });
      
    } catch (error) {
      console.error('Error fetching meeting history:', error);
      res.status(500).json({
        configured: true,
        authenticated: false,
        error: 'Failed to fetch meeting data',
        message: error instanceof Error ? error.message : 'Unknown error',
        data: []
      });
    }
  });

  // Action Items API - Local database storage for action tracking
  // Get all action items from local database
  app.get('/api/action-items', async (req, res) => {
    try {
      const actionItems = await storage.getAllActionItems();
      res.json({ success: true, data: actionItems });
    } catch (error) {
      console.error('Error fetching action items:', error);
      res.status(500).json({ 
        success: false, 
        error: error instanceof Error ? error.message : 'Failed to fetch action items' 
      });
    }
  });

  // Upsert (create or update) action item in local database
  app.post('/api/action-items', async (req, res) => {
    try {
      const { listType, sharePointItemId, actionPriority, actionStatus, actionAssignedTo, actionStartDate, actionDueDate, actionNotes, meetingNotes } = req.body;
      
      // Validate required fields
      if (!listType || typeof listType !== 'string') {
        return res.status(400).json({ 
          success: false, 
          error: 'listType is required and must be a string' 
        });
      }
      
      if (!sharePointItemId || typeof sharePointItemId !== 'string') {
        return res.status(400).json({ 
          success: false, 
          error: 'sharePointItemId is required and must be a string' 
        });
      }
      
      // Normalize optional fields: convert empty strings to null, ensure string types
      const normalizeField = (value: any): string | null => {
        if (value === null || value === undefined) return null;
        const strValue = String(value).trim();
        return strValue === '' ? null : strValue;
      };
      
      const actionItem = await storage.upsertActionItem({
        listType: listType.trim(),
        sharePointItemId: sharePointItemId.trim(),
        actionPriority: normalizeField(actionPriority),
        actionStatus: normalizeField(actionStatus),
        actionAssignedTo: normalizeField(actionAssignedTo),
        actionStartDate: normalizeField(actionStartDate),
        actionDueDate: normalizeField(actionDueDate),
        actionNotes: normalizeField(actionNotes),
        meetingNotes: meetingNotes !== undefined ? meetingNotes : undefined
      });
      
      res.json({ success: true, data: actionItem });
    } catch (error) {
      console.error('Error saving action item:', error);
      res.status(500).json({ 
        success: false, 
        error: error instanceof Error ? error.message : 'Failed to save action item' 
      });
    }
  });

  // AI Title Status endpoint - checks both OpenAI and items needing titles
  app.get('/api/ai-title-status', async (req, res) => {
    try {
      // Check if user is authenticated
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.json({
          working: false,
          needsAI: false,
          emptyTitles: 0,
          totalItems: 0,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      
      // Check OpenAI status
      const openaiStatus = await OpenAIService.verifyAPIConnection();
      
      // Get all three list types to check for empty titles
      const listsService = new SharePointListsService(accessToken);
      const [businessIdeas, safetyIdeas, nearMiss] = await Promise.all([
        listsService.getBusinessIdeas(),
        listsService.getSafetyIdeas(),
        listsService.getNearMiss()
      ]);
      
      // Combine all lists and count items with empty titles
      const allItems = [...businessIdeas, ...safetyIdeas, ...nearMiss];
      const emptyTitleItems = allItems.filter(item => !item.title || item.title.trim() === '');
      const needsAI = openaiStatus.working && emptyTitleItems.length > 0;
      
      res.json({
        working: openaiStatus.working,
        needsAI: needsAI,
        emptyTitles: emptyTitleItems.length,
        totalItems: allItems.length,
        error: openaiStatus.error,
        testTitle: openaiStatus.testTitle
      });
      
    } catch (error) {
      console.error('Error checking AI title status:', error);
      res.status(500).json({ 
        working: false,
        needsAI: false,
        emptyTitles: 0,
        totalItems: 0,
        error: 'Failed to check AI status',
        message: error instanceof Error ? error.message : 'Unknown error'
      });
    }
  });

  // Legacy endpoint - no longer needed with clean SharePoint service
  app.get('/api/unmapped-ids', async (req, res) => {
    res.json({
      configured: true,
      unmappedIds: [],
      totalCount: 0,
      message: 'Clean SharePoint service uses expanded person objects - no ID mapping needed'
    });
  });

  // AI Title Generation endpoint - generates and updates titles using SharePoint REST API
  app.post('/api/run-ai-titles', async (req, res) => {
    try {
      // Check if user is authenticated
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      
      // Check OpenAI status
      const openaiStatus = await OpenAIService.verifyAPIConnection();
      if (!openaiStatus.working) {
        return res.status(400).json({
          success: false,
          error: `OpenAI not available: ${openaiStatus.error}`
        });
      }
      
      // Get all three list types using clean SharePoint Lists service
      const listsService = new SharePointListsService(accessToken);
      const [businessIdeas, safetyIdeas, nearMiss] = await Promise.all([
        listsService.getBusinessIdeas(),
        listsService.getSafetyIdeas(),
        listsService.getNearMiss()
      ]);
      
      // Combine all lists and find items with empty titles
      const allItems = [...businessIdeas, ...safetyIdeas, ...nearMiss];
      const emptyTitleItems = allItems.filter(item => !item.title || item.title.trim() === '');
      
      if (emptyTitleItems.length === 0) {
        return res.json({
          success: true,
          message: 'No items need AI title generation',
          processed: 0
        });
      }

      
      // Prepare items for AI processing
      const itemsForAI = emptyTitleItems.map(item => ({
        id: item.id,
        content: item.description || '',
        type: item.type || 'Business Ideas'
      }));
      
      // Generate titles using OpenAI
      const generatedTitles = await OpenAIService.bulkGenerateTitles(itemsForAI);
      
      // Update SharePoint items with new titles using REST API
      let updatedCount = 0;
      for (let i = 0; i < emptyTitleItems.length && i < generatedTitles.length; i++) {
        const item = emptyTitleItems[i];
        const title = generatedTitles[i];
        
        try {
          // Use the item's type to determine which list to update
          await listsService.updateItemTitle(item.id, title, item.type);
          updatedCount++;
        } catch (error) {
          console.error(`❌ Failed to update ${item.type} item ${item.id}:`, error);
        }
        
        // Add delay between SharePoint updates
        if (i < emptyTitleItems.length - 1) {
          await new Promise(resolve => setTimeout(resolve, 200));
        }
      }
      
      
      res.json({
        success: true,
        message: `Successfully generated and updated ${updatedCount} titles`,
        processed: updatedCount,
        total: emptyTitleItems.length
      });
      
    } catch (error) {
      console.error('Error in AI title generation:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'AI title generation failed'
      });
    }
  });

  // Shareable HTML endpoint - serves HTML files with shareable URLs
  app.get('/api/view-html/:id', (req, res) => {
    try {
      const { id } = req.params;
      const cachedItem = (global as any).htmlCache?.[id];
      
      // Check if cached item exists and hasn't expired
      if (!cachedItem || (cachedItem.expireAt && cachedItem.expireAt < Date.now())) {
        // Clean up expired entry
        if (cachedItem && cachedItem.expireAt < Date.now()) {
          delete (global as any).htmlCache[id];
        }
        
        return res.status(404).send(`
          <html>
            <head>
              <title>Document Not Found</title>
              <meta name="robots" content="noindex, nofollow, noarchive, nosnippet">
            </head>
            <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
              <h1>Document Not Found</h1>
              <p>This HTML document was not found or has expired.</p>
              <p>Documents are available for 7 days after generation.</p>
              <p>Please regenerate the export from the meeting history page.</p>
            </body>
          </html>
        `);
      }
      
      // Add SEO protection headers and serve the HTML
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate'); // Prevent search engine caching
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Expires', '0');
      res.setHeader('X-Robots-Tag', 'noindex, nofollow, noarchive, nosnippet'); // Block search engines
      
      // Extract content from cache structure
      const htmlContent = cachedItem.content || cachedItem; // Support both old and new format
      
      // Add additional SEO protection to HTML content
      const protectedHtml = htmlContent.replace(
        '<head>',
        `<head>
          <meta name="robots" content="noindex, nofollow, noarchive, nosnippet">
          <meta name="googlebot" content="noindex, nofollow, noarchive, nosnippet">`
      );
      
      res.send(protectedHtml);
    } catch (error) {
      console.error('Error serving HTML:', error);
      res.status(500).send(`
        <html>
          <head>
            <title>Error</title>
            <meta name="robots" content="noindex, nofollow, noarchive, nosnippet">
          </head>
          <body>
            <h1>Error serving document</h1>
            <p>Please try again later.</p>
          </body>
        </html>
      `);
    }
  });

  // HTML export functionality for meeting data
  app.post('/api/generate-meeting-html', async (req, res) => {
    try {
      const { meetingData, selectedMeeting, selectedType, meetingAttendance } = req.body;
      
      if (!meetingData || meetingData.length === 0) {
        return res.status(400).json({ error: 'No meeting data provided' });
      }

      // Filter data based on selection
      let filteredData = meetingData;
      if (selectedMeeting !== 'all') {
        filteredData = meetingData.filter((item: any) => {
          const itemDate = new Date(item.meetingDate).toISOString().split('T')[0];
          const selectedDate = new Date(selectedMeeting).toISOString().split('T')[0];
          return itemDate === selectedDate;
        });
      }

      if (selectedType !== 'all') {
        filteredData = filteredData.filter((item: any) => item.type === selectedType);
      }

      // Sort data according to requested priority: Near Miss → Safety Ideas → Business Ideas (prioritizing feedback-related)
      filteredData.sort((a: any, b: any) => {
        const typeOrder = { 'Near Miss': 1, 'Safety Ideas': 2, 'Business Ideas': 3, 'Actions': 4 };
        const aOrder = typeOrder[a.type as keyof typeof typeOrder] || 5;
        const bOrder = typeOrder[b.type as keyof typeof typeOrder] || 5;
        
        if (aOrder !== bOrder) {
          return aOrder - bOrder;
        }
        
        // Within Business Ideas, prioritize feedback-related items
        if (a.type === 'Business Ideas' && b.type === 'Business Ideas') {
          const aIsFeedback = (a.title || a.description || '').toLowerCase().includes('feedback') || 
                             (a.ideaType || '').toLowerCase().includes('feedback');
          const bIsFeedback = (b.title || b.description || '').toLowerCase().includes('feedback') || 
                             (b.ideaType || '').toLowerCase().includes('feedback');
          
          if (aIsFeedback && !bIsFeedback) return -1;
          if (!aIsFeedback && bIsFeedback) return 1;
        }
        
        // Otherwise sort by submission date (newest first)
        return new Date(b.submittedDate).getTime() - new Date(a.submittedDate).getTime();
      });

      // Generate HTML content
      const meetingDate = selectedMeeting === 'all' ? 'All Meetings' : new Date(selectedMeeting).toLocaleDateString('en-GB');
      const currentDate = new Date().toLocaleDateString('en-GB');

      // Fetch signatures directly from DB — more reliable than trusting client-sent data
      let meetingSignatures: Record<string, Record<string, { status: string; signatureData: string | null; signedAt: string }>> | undefined;
      try {
        if (selectedMeeting && selectedMeeting !== 'all') {
          // Try exact key first, then fallback to all-signatures with date normalization
          let sigsForDate = await storage.getMeetingSignatures(selectedMeeting);
          if (Object.keys(sigsForDate).length === 0) {
            // Try normalizing date to YYYY-MM-DD
            const normalizedSelected = new Date(selectedMeeting).toISOString().split('T')[0];
            sigsForDate = await storage.getMeetingSignatures(normalizedSelected);
          }
          if (Object.keys(sigsForDate).length === 0) {
            // Fallback: search all signatures for a date-matching key
            const allSigs = await storage.getAllMeetingSignatures();
            const normalizedSelected = new Date(selectedMeeting).toISOString().split('T')[0];
            const matchingKey = Object.keys(allSigs).find(k => {
              try { return new Date(k).toISOString().split('T')[0] === normalizedSelected; } catch { return false; }
            });
            if (matchingKey) sigsForDate = allSigs[matchingKey];
          }
          meetingSignatures = { [selectedMeeting]: sigsForDate };
        } else {
          meetingSignatures = await storage.getAllMeetingSignatures();
        }
      } catch (sigErr) {
        console.warn('Could not fetch signatures from DB for export:', sigErr);
        meetingSignatures = req.body.meetingSignatures;
      }

      const htmlContent = generateMeetingMinutesHTML(filteredData, meetingDate, currentDate, meetingAttendance, selectedMeeting, meetingSignatures);
      
      // Generate unique ID for shareable URL
      const shareId = 'meeting-' + Date.now() + '-' + Math.random().toString(36).substr(2, 9);
      
      // Store HTML in memory cache with expiration (7 days)
      if (!(global as any).htmlCache) {
        (global as any).htmlCache = {};
      }
      const expireAt = Date.now() + (7 * 24 * 60 * 60 * 1000); // 7 days from now
      (global as any).htmlCache[shareId] = {
        content: htmlContent,
        expireAt: expireAt
      };
      
      // Clean up expired cache entries
      const currentTime = Date.now();
      Object.keys((global as any).htmlCache).forEach(key => {
        if ((global as any).htmlCache[key].expireAt < currentTime) {
          delete (global as any).htmlCache[key];
        }
      });
      
      // Also clean up old cache entries if too many (keep only last 100)
      const cacheKeys = Object.keys((global as any).htmlCache);
      if (cacheKeys.length > 100) {
        const oldKeys = cacheKeys.slice(0, -100);
        oldKeys.forEach(key => delete (global as any).htmlCache[key]);
      }
      
      // Set appropriate headers for HTML download
      const filename = selectedMeeting === 'all' 
        ? `Cranfield-Glass-Meeting-Minutes-${currentDate.replace(/\//g, '-')}.html`
        : `Cranfield-Glass-Meeting-${new Date(selectedMeeting).toLocaleDateString('en-GB').replace(/\//g, '-')}.html`;
      
      // Generate shareable URL
      const shareUrl = `${req.protocol}://${req.get('host')}/api/view-html/${shareId}`;
      
      res.json({
        success: true,
        filename: filename,
        htmlContent: htmlContent,
        shareUrl: shareUrl,
        shareId: shareId,
        message: 'HTML meeting minutes generated successfully with shareable URL'
      });
      
    } catch (error) {
      console.error('Error generating HTML:', error);
      res.status(500).json({ error: 'Failed to generate HTML meeting minutes' });
    }
  });

  // Markdown export functionality for PDF generation (Typora-style navigation)
  app.post('/api/generate-meeting-markdown', async (req, res) => {
    try {
      const { meetingData, selectedMeeting, selectedType, meetingAttendance } = req.body;
      
      if (!meetingData || meetingData.length === 0) {
        return res.status(400).json({ error: 'No meeting data provided' });
      }

      // Filter data based on selection
      let filteredData = meetingData;
      if (selectedMeeting !== 'all') {
        filteredData = meetingData.filter((item: any) => {
          const itemDate = new Date(item.meetingDate).toISOString().split('T')[0];
          const selectedDate = new Date(selectedMeeting).toISOString().split('T')[0];
          return itemDate === selectedDate;
        });
      }

      if (selectedType !== 'all') {
        filteredData = filteredData.filter((item: any) => item.type === selectedType);
      }

      // Sort data according to priority: Near Miss → Safety Ideas → Business Ideas
      filteredData.sort((a: any, b: any) => {
        const typeOrder = { 'Near Miss': 1, 'Safety Ideas': 2, 'Business Ideas': 3, 'Actions': 4 };
        const aOrder = typeOrder[a.type as keyof typeof typeOrder] || 5;
        const bOrder = typeOrder[b.type as keyof typeof typeOrder] || 5;
        
        if (aOrder !== bOrder) {
          return aOrder - bOrder;
        }
        
        // Within type, sort by submission date (newest first)
        return new Date(b.submittedDate).getTime() - new Date(a.submittedDate).getTime();
      });

      // Generate Markdown content with proper heading hierarchy for PDF bookmarks
      const meetingDate = selectedMeeting === 'all' ? 'All Meetings' : new Date(selectedMeeting).toLocaleDateString('en-GB');
      const currentDate = new Date().toLocaleDateString('en-GB');
      
      const markdownContent = MarkdownMeetingGenerator.generateMeetingMarkdown(
        filteredData, 
        meetingDate, 
        currentDate, 
        meetingAttendance, 
        selectedMeeting
      );
      
      // Set appropriate headers for markdown download
      const filename = selectedMeeting === 'all' 
        ? `Cranfield-Glass-Meeting-Minutes-${currentDate.replace(/\//g, '-')}.md`
        : `Cranfield-Glass-Meeting-${new Date(selectedMeeting).toLocaleDateString('en-GB').replace(/\//g, '-')}.md`;
      
      res.setHeader('Content-Type', 'text/markdown; charset=utf-8');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.send(markdownContent);
      
    } catch (error) {
      console.error('Error generating markdown meeting minutes:', error);
      res.status(500).json({ error: 'Failed to generate markdown meeting minutes' });
    }
  });

  // CSV export functionality for meeting data
  app.post('/api/generate-meeting-csv', async (req, res) => {
    try {
      const { meetingData, selectedMeeting, selectedType } = req.body;
      
      if (!meetingData || meetingData.length === 0) {
        return res.status(400).json({ error: 'No meeting data provided' });
      }

      // Filter data based on selection
      let filteredData = meetingData;
      if (selectedMeeting !== 'all') {
        filteredData = meetingData.filter((item: any) => {
          const itemDate = new Date(item.meetingDate).toISOString().split('T')[0];
          const selectedDate = new Date(selectedMeeting).toISOString().split('T')[0];
          return itemDate === selectedDate;
        });
      }

      if (selectedType !== 'all') {
        filteredData = filteredData.filter((item: any) => item.type === selectedType);
      }

      // Sort data according to requested priority: Near Miss → Safety Ideas → Business Ideas (prioritizing feedback-related)
      filteredData.sort((a: any, b: any) => {
        const typeOrder = { 'Near Miss': 1, 'Safety Ideas': 2, 'Business Ideas': 3, 'Actions': 4 };
        const aOrder = typeOrder[a.type as keyof typeof typeOrder] || 5;
        const bOrder = typeOrder[b.type as keyof typeof typeOrder] || 5;
        
        if (aOrder !== bOrder) {
          return aOrder - bOrder;
        }
        
        // Within Business Ideas, prioritize feedback-related items
        if (a.type === 'Business Ideas' && b.type === 'Business Ideas') {
          const aIsFeedback = (a.title || a.description || '').toLowerCase().includes('feedback') || 
                             (a.ideaType || '').toLowerCase().includes('feedback');
          const bIsFeedback = (b.title || b.description || '').toLowerCase().includes('feedback') || 
                             (b.ideaType || '').toLowerCase().includes('feedback');
          
          if (aIsFeedback && !bIsFeedback) return -1;
          if (!aIsFeedback && bIsFeedback) return 1;
        }
        
        // Otherwise sort by submission date (newest first)
        return new Date(b.submittedDate).getTime() - new Date(a.submittedDate).getTime();
      });

      // Data processed successfully

      // Generate analytics for this meeting
      const analytics = generateMeetingAnalytics(filteredData);

      // Create analytics summary section for CSV
      const analyticsRows = [
        ['MEETING ANALYTICS DASHBOARD'],
        [''],
        ['📊 Summary Statistics'],
        ['Total Items', analytics.totalItems],
        ['Outstanding Actions', analytics.outstandingActions],
        ['Completion Rate', `${analytics.completionRate}%`],
        [''],
        ['📈 Status Breakdown'],
        ...Object.entries(analytics.statusBreakdown).map(([status, count]) => [status, count]),
        [''],
        ['📊 Type Distribution'],
        ...Object.entries(analytics.typeBreakdown).map(([type, count]) => [type, count]),
        [''],
        ['👥 Top Contributors'],
        ...analytics.topContributors.map((contributor: any, index: number) => [`#${index + 1} ${contributor.name}`, `${contributor.count} items`]),
        [''],
        ['📋 Assignment Overview'],
        ...Object.entries(analytics.assignmentStats).slice(0, 5).map(([person, count]) => [person, `${count} items`]),
        [''],
        ['MEETING MINUTES DETAIL'],
        ['']
      ];

      // Professional Meeting Minutes CSV format matching Word template
      const csvHeaders = [
        'Agenda Item',
        'Discussion Notes',
        'Meeting Notes',
        'HSO Action',
        'Follow-up Required',
        'Type',
        'Status',
        'Submitted By',
        'Meeting Date',
        'Priority',
        'ID Reference'
      ];

      const csvRows = filteredData.map((item: any) => {
        // Format agenda item title professionally
        const agendaItem = item.title || `${item.type} Item`;
        
        // Discussion Notes = the actual idea/description content
        let discussionNotes = item.description || '';
        if (item.secondaryDescription) {
          discussionNotes += discussionNotes ? `\n\nHow it happened: ${item.secondaryDescription}` : item.secondaryDescription;
        }
        
        // Meeting Notes = actual meeting notes from SharePoint
        const meetingNotes = item.meetingNotes || '';
        
        // Remove fake HSO action text - we don't have an HSO
        const hsoAction = '';
        
        // Follow-up requirements
        let followUp = '';
        if (item.assignedTo) {
          followUp = `Assigned to: ${item.assignedTo}`;
        }
        if (item.actionStartDate) {
          followUp += followUp ? ` | Start Date: ${item.actionStartDate}` : `Start Date: ${item.actionStartDate}`;
        }
        
        return [
          `"${agendaItem.replace(/"/g, '""')}"`,
          `"${discussionNotes.replace(/"/g, '""')}"`,
          `"${meetingNotes.replace(/"/g, '""')}"`,
          `"${hsoAction.replace(/"/g, '""')}"`,
          `"${followUp.replace(/"/g, '""')}"`,
          item.type || '',
          item.status || '',
          item.submittedBy || '',
          item.meetingDate || '',
          item.priority || 'Standard',
          item.id || ''
        ];
      });

      const csvContent = [
        // Analytics section first
        ...analyticsRows.map((row: any[]) => row.map((cell: any) => `"${String(cell).replace(/"/g, '""')}"`).join(',')),
        // Then headers and meeting data
        csvHeaders.join(','),
        ...csvRows.map((row: any[]) => row.join(','))
      ].join('\n');

      // Set appropriate headers for CSV download
      const currentDate = new Date().toLocaleDateString('en-GB').replace(/\//g, '-');
      let filename;
      
      if (selectedMeeting === 'all') {
        filename = `Cranfield-Glass-Meeting-Data-${currentDate}.csv`;
      } else {
        const meetingDateFormatted = new Date(selectedMeeting).toLocaleDateString('en-GB').replace(/\//g, '-');
        filename = `Cranfield-Glass-Meeting-${meetingDateFormatted}.csv`;
      }

      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.setHeader('Content-Length', Buffer.byteLength(csvContent));
      
      res.send(csvContent);
      
    } catch (error) {
      console.error('Error generating CSV:', error);
      res.status(500).json({ error: 'Failed to generate CSV' });
    }
  });

  // Merge meeting dates to next meeting functionality
  app.post('/api/merge-to-next-meeting', async (req, res) => {
    try {
      // Check if user is authenticated
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { currentMeetingDate, nextMeetingDate } = req.body;

      if (!currentMeetingDate || !nextMeetingDate) {
        return res.status(400).json({
          success: false,
          error: 'Both current and next meeting dates are required'
        });
      }

      // Get all items from the current meeting date
      const listsService = new SharePointListsService(accessToken);
      const [businessIdeas, safetyIdeas, nearMiss] = await Promise.all([
        listsService.getBusinessIdeas(),
        listsService.getSafetyIdeas(),
        listsService.getNearMiss()
      ]);

      // Combine all lists and filter by current meeting date
      const allItems = [...businessIdeas, ...safetyIdeas, ...nearMiss];
      const currentDateKey = new Date(currentMeetingDate).toISOString().split('T')[0];
      
      console.log(`   Target date: ${currentDateKey}`);
      console.log(`   Total items: ${allItems.length} (${businessIdeas.length} Business, ${safetyIdeas.length} Safety, ${nearMiss.length} Near Miss)`);
      
      // Debug: Show first few actual meeting dates in SharePoint
      allItems.slice(0, 5).forEach((item, idx) => {
        const itemDateKey = new Date(item.meetingDate).toISOString().split('T')[0];
        console.log(`   ${idx + 1}. ${item.type} (${item.id}): ${item.meetingDate} -> ${itemDateKey}`);
      });
      
      // Debug: Show unique meeting dates
      const uniqueMeetingDates = Array.from(new Set(allItems.map((item: any) => new Date(item.meetingDate).toISOString().split('T')[0])));
      
      const itemsToUpdate = allItems.filter(item => {
        const itemDateKey = new Date(item.meetingDate).toISOString().split('T')[0];
        const matches = itemDateKey === currentDateKey;
        if (matches) {
        }
        return matches;
      });
      

      if (itemsToUpdate.length === 0) {
        return res.json({
          success: true,
          message: 'No items found for the specified meeting date',
          updated: 0
        });
      }


      // Update meeting dates for all items with detailed logging
      let updatedCount = 0;
      const totalItems = itemsToUpdate.length;
      
      for (let i = 0; i < totalItems; i++) {
        const item = itemsToUpdate[i];
        try {
          await listsService.updateItemMeetingDate(item.id, nextMeetingDate, item.type);
          updatedCount++;
        } catch (error) {
          console.error(`❌ [${i + 1}/${totalItems}] Failed to update ${item.type} item ${item.id}:`, error);
        }
        
        // Add delay between SharePoint updates
        await new Promise(resolve => setTimeout(resolve, 200));
      }


      res.json({
        success: true,
        message: `Successfully merged ${updatedCount} items to next meeting`,
        updated: updatedCount,
        total: itemsToUpdate.length
      });

    } catch (error) {
      console.error('Error in meeting merge:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Meeting merge failed'
      });
    }
  });

  // High-quality Word document generation with professional styling (kept for compatibility)
  app.post('/api/generate-meeting-word', async (req, res) => {
    try {
      const { meetingData, selectedMeeting, selectedType, meetingAttendance } = req.body;
      
      if (!meetingData || meetingData.length === 0) {
        return res.status(400).json({ error: 'No meeting data provided' });
      }

      // Generate Word document using docx library
      const wordBuffer = await generateMeetingWordDoc({
        meetingData,
        selectedMeeting,
        selectedType,
        meetingAttendance
      });

      // Set appropriate headers for Word document download
      const currentDate = new Date().toLocaleDateString('en-GB').replace(/\//g, '-');
      let filename;
      
      if (selectedMeeting === 'all') {
        filename = `Cranfield-Glass-Meeting-Minutes-${currentDate}.docx`;
      } else {
        const meetingDateFormatted = new Date(selectedMeeting).toLocaleDateString('en-GB').replace(/\//g, '-');
        filename = `Cranfield-Glass-Meeting-${meetingDateFormatted}.docx`;
      }

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.setHeader('Content-Length', wordBuffer.length);
      
      res.send(wordBuffer);
      
    } catch (error) {
      console.error('Error generating Word document:', error);
      res.status(500).json({ error: 'Failed to generate Word document' });
    }
  });

  // Create new SharePoint list item
  app.post('/api/sharepoint/create-item', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { listType, itemData, deferTitle } = req.body;

      if (!listType || !itemData) {
        return res.status(400).json({
          success: false,
          error: 'Missing required fields: listType and itemData'
        });
      }

      const listsService = new SharePointListsService(accessToken);
      const newItemId = await listsService.createListItem(listType, itemData);

      const generateTitle = async () => {
        try {
          if (itemData.description && itemData.description.trim()) {
            const aiTitle = await OpenAIService.generateSmartTitle(itemData.description, listType);
            if (aiTitle && aiTitle !== 'New Idea') {
              await listsService.updateItemTitle(newItemId, aiTitle, listType);
            }
          }
        } catch (aiError) {
          console.error('AI title generation failed for new item:', aiError);
          // Continue anyway - the item was created successfully
        }
      };

      // When deferTitle is set (e.g. the Teams tab), respond as soon as the
      // item is persisted and generate the AI title in the background. This
      // makes submission feel instant. The default path keeps the original
      // synchronous behaviour so existing main-app callers are unaffected.
      if (deferTitle) {
        void generateTitle();
      } else {
        await generateTitle();
      }

      res.json({
        success: true,
        itemId: newItemId,
        message: `Successfully created ${listType} item`
      });

    } catch (error) {
      console.error('Error creating SharePoint item:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create item'
      });
    }
  });

  // Update SharePoint list item
  app.post('/api/sharepoint/update-item', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { itemId, listType, updates } = req.body;

      if (!itemId || !listType || !updates) {
        return res.status(400).json({
          success: false,
          error: 'Missing required fields: itemId, listType, and updates'
        });
      }

      const listsService = new SharePointListsService(accessToken);
      await listsService.updateItemFields(itemId, listType, updates);

      res.json({
        success: true,
        message: `Successfully updated ${listType} item ${itemId}`
      });

    } catch (error) {
      console.error('Error updating SharePoint item:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update item'
      });
    }
  });

  // Move a SharePoint list item to a different list
  app.post('/api/sharepoint/move-item', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { itemId, fromList, toList } = req.body;

      if (!itemId || !fromList || !toList) {
        return res.status(400).json({
          success: false,
          error: 'Missing required fields: itemId, fromList, and toList'
        });
      }

      // itemId is a formatted id like "safety-ideas-42"; ensure it ends with a numeric id
      if (typeof itemId !== 'string' || !/-\d+$/.test(itemId)) {
        return res.status(400).json({
          success: false,
          error: 'Invalid itemId format'
        });
      }

      const listsService = new SharePointListsService(accessToken);
      const newItemId = await listsService.moveItem(itemId, fromList, toList);

      res.json({
        success: true,
        newItemId,
        message: `Successfully moved item from ${fromList} to ${toList}`
      });

    } catch (error) {
      console.error('Error moving SharePoint item:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to move item'
      });
    }
  });

  // Get SharePoint Choice field options
  app.get('/api/sharepoint/choice-options/:listType/:fieldName', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { listType, fieldName } = req.params;

      const listsService = new SharePointListsService(accessToken);
      const choices = await listsService.getChoiceFieldOptions(listType, fieldName);

      res.json({
        success: true,
        choices,
        listType,
        fieldName
      });

    } catch (error) {
      console.error('Error fetching choice options:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch choice options'
      });
    }
  });

  // Get SharePoint users for person field dropdowns
  app.get('/api/sharepoint/users', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      
      // Get site users from SharePoint
      const response = await fetch('https://cranfieldglass.sharepoint.com/_api/web/siteusers?$select=Id,Title,Email,LoginName&$top=100', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (!response.ok) {
        throw new Error(`SharePoint API error: ${response.status}`);
      }
      
      const data = await response.json();
      
      // Define approved staff members only - based on actual SharePoint users
      const approvedStaff = [
        'Hoani Hunt',
        'Simon Hubbard', 
        'James Waites',
        'Emma White',
        'Kevin Young',
        'Ryan Newman',
        'Dan Conlan',
        'Struan O\'Donnell',
        'Sam Chang'
      ];
      
      // Filter and deduplicate users, preferring specific emails
      const filteredUsers = data.d.results
        .filter((user: any) => user.Title && approvedStaff.includes(user.Title))
        .map((user: any) => ({
          id: user.Id,
          title: user.Title,
          email: user.Email,
          loginName: user.LoginName
        }));

      // Debug logging to see all users
      filteredUsers.forEach((user: any) => {
        console.log(`  - ${user.title} (ID: ${user.id}, Email: ${user.email})`);
      });

      // Remove duplicates, preferring specific email addresses
      const uniqueUsers = new Map();
      filteredUsers.forEach((user: any) => {
        const existingUser = uniqueUsers.get(user.title);
        if (!existingUser) {
          uniqueUsers.set(user.title, user);
        } else {
          // Prefer users with preferred email domains or specific emails
          if (user.title === 'Simon Hubbard' && user.email === 'simon@cranfield.co.nz') {
            uniqueUsers.set(user.title, user);
          }
          // For other duplicates, keep the first one unless we have a preference
        }
      });

      const users = Array.from(uniqueUsers.values())
        .sort((a: any, b: any) => a.title.localeCompare(b.title));
      
      users.forEach(user => {
        console.log(`  - ${user.title} (${user.email})`);
      });

      res.json({ 
        success: true,
        users 
      });
    } catch (error) {
      console.error('❌ Failed to get SharePoint users:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch SharePoint users'
      });
    }
  });

  // Get current user's role information
  app.get('/api/current-user/role', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      
      // Get user info from Microsoft Graph
      const response = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });
      
      if (!response.ok) {
        throw new Error(`Microsoft Graph API error: ${response.status}`);
      }
      
      const userInfo = await response.json();
      const userEmail = userInfo.mail || userInfo.userPrincipalName;
      
      // Look up staff record by email to get role information
      const staffRecord = await storage.getStaffByEmail(userEmail);
      
      if (!staffRecord) {
        // Default to field role if not found in staff database
        return res.json({ 
          success: true,
          user: {
            email: userEmail,
            name: userInfo.displayName,
            jobTitle: userInfo.jobTitle,
            role: 'field',
            roleRank: 3
          }
        });
      }
      
      res.json({ 
        success: true,
        user: {
          email: userEmail,
          name: userInfo.displayName,
          jobTitle: userInfo.jobTitle,
          role: staffRecord.role,
          roleRank: staffRecord.roleRank
        }
      });
    } catch (error) {
      console.error('❌ Failed to get current user role:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch user role'
      });
    }
  });

  // AI Note Enhancement endpoint
  app.post('/api/ai-enhance-notes', async (req, res) => {
    try {
      // Check if user is authenticated
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const { content, itemType } = req.body;
      
      if (!content || typeof content !== 'string') {
        return res.status(400).json({
          success: false,
          error: 'Content is required'
        });
      }
      
      // Check OpenAI status
      const openaiStatus = await OpenAIService.verifyAPIConnection();
      if (!openaiStatus.working) {
        return res.status(400).json({
          success: false,
          error: `AI enhancement not available: ${openaiStatus.error}`
        });
      }
      
      
      // Enhance the content using OpenAI
      const enhancedContent = await OpenAIService.enhanceNotes(content, itemType || 'Meeting Notes');
      
      
      res.json({
        success: true,
        enhancedContent
      });
      
    } catch (error) {
      console.error('AI note enhancement error:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'AI note enhancement failed'
      });
    }
  });

  // AI Auto-Suggestions endpoint
  app.post('/api/ai-suggestions', async (req, res) => {
    try {
      // Check if user is authenticated
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const { content, type } = req.body;
      
      if (!content || typeof content !== 'string') {
        return res.status(400).json({
          success: false,
          error: 'Content is required'
        });
      }
      
      // Check OpenAI status
      const openaiStatus = await OpenAIService.verifyAPIConnection();
      if (!openaiStatus.working) {
        return res.status(400).json({
          success: false,
          error: `AI suggestions not available: ${openaiStatus.error}`
        });
      }
      
      
      // Generate suggestions using OpenAI
      const suggestions = await OpenAIService.generateSuggestions(content, type || 'Business Ideas');
      
      
      res.json({
        success: true,
        suggestions
      });
      
    } catch (error) {
      console.error('AI suggestions error:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'AI suggestions failed'
      });
    }
  });

  // Analyze job listing section using OpenAI
  app.post('/api/analyze-job-section', async (req, res) => {
    try {
      const { sectionText, sectionType } = req.body;
      
      if (!sectionText || typeof sectionText !== 'string' || sectionText.trim().length === 0) {
        return res.status(400).json({
          success: false,
          error: 'Section text is required'
        });
      }
      
      // Check OpenAI status
      const openaiStatus = await OpenAIService.verifyAPIConnection();
      if (!openaiStatus.working) {
        return res.status(400).json({
          success: false,
          error: `OpenAI not available: ${openaiStatus.error}`
        });
      }
      
      // Analyze the job listing section using OpenAI
      const analysis = await OpenAIService.analyzeJobListingSection(sectionText, sectionType || 'Job Section');
      
      res.json({
        success: true,
        analysis
      });
      
    } catch (error) {
      console.error('Job section analysis failed:', error);
      res.status(500).json({
        success: false,
        error: 'Failed to analyze job section'
      });
    }
  });

  // Normalize any date string to YYYY-MM-DD for consistent lock key storage
  const normaliseLockDate = (raw: string): string => {
    try {
      return new Date(raw).toISOString().split('T')[0];
    } catch {
      return raw;
    }
  };

  // Meeting lock endpoints
  // GET all locks as { [YYYY-MM-DD]: boolean }
  app.get('/api/meeting-locks', async (req, res) => {
    try {
      const { db } = await import('./db');
      const { meetingLocks } = await import('../shared/schema');
      const rows = await db.select().from(meetingLocks);
      const locks: Record<string, boolean> = {};
      for (const row of rows) {
        locks[row.meetingDate] = row.isLocked;
      }
      res.json({ success: true, locks });
    } catch (error) {
      console.error('Error fetching all meeting locks:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch meeting locks' });
    }
  });

  app.get('/api/meeting-locks/:meetingDate', async (req, res) => {
    try {
      const meetingDate = normaliseLockDate(decodeURIComponent(req.params.meetingDate));
      const meetingLock = await storage.getMeetingLock(meetingDate);
      
      res.json({
        success: true,
        meetingLock: meetingLock || { meetingDate, isLocked: false, isClosed: false, lockedAt: null, lockedBy: null, closedAt: null, closedBy: null }
      });
    } catch (error) {
      console.error('Error fetching meeting lock:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch meeting lock'
      });
    }
  });

  app.post('/api/meeting-locks', async (req, res) => {
    try {
      const { meetingDate, isLocked, lockedBy } = req.body;
      
      if (!meetingDate) {
        return res.status(400).json({
          success: false,
          error: 'Meeting date is required'
        });
      }

      const meetingLock = await storage.updateMeetingLock(normaliseLockDate(meetingDate), isLocked, lockedBy);
      
      res.json({
        success: true,
        meetingLock
      });
    } catch (error) {
      console.error('Error updating meeting lock:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update meeting lock'
      });
    }
  });

  // Meeting closed endpoints
  app.post('/api/meeting-closed', async (req, res) => {
    try {
      const { meetingDate, isClosed, closedBy } = req.body;
      
      if (!meetingDate) {
        return res.status(400).json({
          success: false,
          error: 'Meeting date is required'
        });
      }

      const meetingLock = await storage.updateMeetingClosed(meetingDate, isClosed, closedBy);
      
      res.json({
        success: true,
        meetingLock
      });
    } catch (error) {
      console.error('Error updating meeting closed status:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update meeting closed status'
      });
    }
  });

  // Meeting attendance endpoints
  app.get('/api/meeting-attendance', async (req, res) => {
    try {
      const attendanceData = await storage.getAllMeetingAttendance();
      
      res.json({
        success: true,
        attendance: attendanceData
      });
    } catch (error) {
      console.error('Error fetching meeting attendance:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch meeting attendance'
      });
    }
  });

  app.get('/api/meeting-attendance/:meetingDate', async (req, res) => {
    try {
      const { meetingDate } = req.params;
      const attendanceRecords = await storage.getMeetingAttendance(meetingDate);
      
      // Convert to the format expected by the frontend
      const attendees = attendanceRecords
        .filter(record => record.isPresent)
        .map(record => record.attendeeName);
      
      res.json({
        success: true,
        attendees
      });
    } catch (error) {
      console.error('Error fetching meeting attendance:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch meeting attendance'
      });
    }
  });

  app.post('/api/meeting-attendance', async (req, res) => {
    try {
      const { meetingDate, attendeeName, isPresent } = req.body;
      
      if (!meetingDate || !attendeeName) {
        return res.status(400).json({
          success: false,
          error: 'Meeting date and attendee name are required'
        });
      }

      const lock = await storage.getMeetingLock(meetingDate);
      if (lock?.isLocked) {
        return res.status(403).json({
          success: false,
          error: 'Attendance is locked — no changes permitted'
        });
      }

      const attendanceRecord = await storage.updateMeetingAttendance(
        meetingDate, 
        attendeeName, 
        isPresent !== false
      );
      
      res.json({
        success: true,
        attendanceRecord
      });
    } catch (error) {
      console.error('Error updating meeting attendance:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update meeting attendance'
      });
    }
  });

  // Meeting signature endpoints
  app.get('/api/meeting-signatures', async (req, res) => {
    try {
      const signatures = await storage.getAllMeetingSignatures();
      res.json({ success: true, signatures });
    } catch (error) {
      console.error('Error fetching meeting signatures:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch signatures' });
    }
  });

  app.get('/api/meeting-signatures/:meetingDate', async (req, res) => {
    try {
      const { meetingDate } = req.params;
      const signatures = await storage.getMeetingSignatures(decodeURIComponent(meetingDate));
      res.json({ success: true, signatures });
    } catch (error) {
      console.error('Error fetching meeting signatures:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch signatures' });
    }
  });

  app.post('/api/meeting-signatures', async (req, res) => {
    try {
      const { meetingDate, attendeeName, status, signatureData, signedAt } = req.body;
      if (!meetingDate || !attendeeName || !status) {
        return res.status(400).json({ success: false, error: 'meetingDate, attendeeName, and status are required' });
      }
      const existingLock = await storage.getMeetingLock(meetingDate);
      if (existingLock?.isLocked) {
        return res.status(403).json({ success: false, error: 'Attendance is locked — signatures cannot be modified' });
      }
      const record = await storage.updateMeetingSignature(meetingDate, attendeeName, status, signatureData ?? null, signedAt ?? new Date().toISOString());
      res.json({ success: true, record });
    } catch (error) {
      console.error('Error saving signature:', error);
      res.status(500).json({ success: false, error: 'Failed to save signature' });
    }
  });

  // Card ordering endpoints
  app.get('/api/card-ordering', async (req, res) => {
    try {
      const orderedCards = await db
        .select()
        .from(cardOrdering)
        .orderBy(cardOrdering.position);
      
      res.json({
        success: true,
        cardOrdering: orderedCards
      });
    } catch (error) {
      console.error('Error fetching card ordering:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to fetch card ordering'
      });
    }
  });

  app.post('/api/card-ordering', async (req, res) => {
    try {
      const { cardOrders } = req.body;
      
      if (!Array.isArray(cardOrders)) {
        return res.status(400).json({
          success: false,
          error: 'cardOrders must be an array'
        });
      }

      // Validate each card order entry
      for (const cardOrder of cardOrders) {
        const parseResult = insertCardOrderingSchema.safeParse(cardOrder);
        if (!parseResult.success) {
          return res.status(400).json({
            success: false,
            error: `Invalid card order data: ${parseResult.error.message}`
          });
        }
      }

      // Clear existing ordering
      await db.delete(cardOrdering);

      // Insert new ordering
      if (cardOrders.length > 0) {
        await db.insert(cardOrdering).values(cardOrders);
      }
      
      res.json({
        success: true,
        message: 'Card ordering updated successfully'
      });
    } catch (error) {
      console.error('Error updating card ordering:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update card ordering'
      });
    }
  });

  // Image proxy for SharePoint URLs requiring authentication
  app.get('/api/images/proxy', async (req, res) => {
    try {
      const { url } = req.query;
      
      if (!url || typeof url !== 'string') {
        return res.status(400).json({ error: 'URL parameter is required' });
      }

      // Validate allowed hosts for security
      const allowedHosts = ['graph.microsoft.com', 'cranfieldglass.sharepoint.com'];
      const urlObj = new URL(url);
      const isAllowedHost = allowedHosts.some(host => 
        urlObj.hostname === host || urlObj.hostname.endsWith(`.${host}`)
      );

      if (!isAllowedHost) {
        return res.status(403).json({ error: 'Host not allowed' });
      }

      // Get access token for SharePoint
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({ error: 'Authentication required' });
      }
      const accessToken = authHeader.substring(7);
      
      // Fetch the image with authentication
      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'image/*'
        }
      });

      if (!response.ok) {
        console.error(`Failed to fetch image: ${response.status} ${response.statusText}`);
        return res.status(response.status).json({ error: 'Failed to fetch image' });
      }

      // Check content type is an image
      const contentType = response.headers.get('content-type');
      if (!contentType || !contentType.startsWith('image/')) {
        return res.status(400).json({ error: 'URL does not point to an image' });
      }

      // Set appropriate headers
      res.set({
        'Content-Type': contentType,
        'Cache-Control': 'public, max-age=3600', // Cache for 1 hour
        'Access-Control-Allow-Origin': '*'
      });

      // Stream the image
      const buffer = await response.arrayBuffer();
      res.send(Buffer.from(buffer));

    } catch (error) {
      console.error('Image proxy error:', error);
      res.status(500).json({ error: 'Internal server error' });
    }
  });

  // ============================================================================
  // SKILLS MATRIX API ROUTES - Replacing Excel-based training management
  // ============================================================================

  // Import validation schemas for skills matrix
  const {
    insertStaffSchema,
    insertSkillSchema,
    insertTrainingRecordSchema,
    insertPpeRecordSchema,
    insertInductionRecordSchema,
    insertEquipmentAuthorizationSchema,
    insertPhotoAssetSchema,
    insertGloveRecordSchema
  } = await import("@shared/schema");

  // STAFF MANAGEMENT ROUTES
  // ----------------------

  // Get all staff
  app.get('/api/staff', async (req, res) => {
    try {
      const allStaff = await storage.getAllStaff();
      res.json({ success: true, data: allStaff });
    } catch (error) {
      console.error('Error fetching staff:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch staff' });
    }
  });

  // Get staff by ID
  app.get('/api/staff/:id', async (req, res) => {
    try {
      const staffId = parseInt(req.params.id);
      if (isNaN(staffId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID' });
      }
      
      const staff = await storage.getStaff(staffId);
      if (!staff) {
        return res.status(404).json({ success: false, error: 'Staff member not found' });
      }
      
      res.json({ success: true, data: staff });
    } catch (error) {
      console.error('Error fetching staff:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch staff' });
    }
  });

  // Create new staff member
  app.post('/api/staff', async (req, res) => {
    try {
      const parseResult = insertStaffSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid staff data: ${parseResult.error.message}`
        });
      }

      const newStaff = await storage.createStaff(parseResult.data);
      res.status(201).json({ success: true, data: newStaff });
    } catch (error) {
      console.error('Error creating staff:', error);
      res.status(500).json({ success: false, error: 'Failed to create staff member' });
    }
  });

  // Update staff member
  app.put('/api/staff/:id', async (req, res) => {
    try {
      const staffId = parseInt(req.params.id);
      if (isNaN(staffId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID' });
      }

      const parseResult = insertStaffSchema.partial().safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid staff data: ${parseResult.error.message}`
        });
      }

      const updatedStaff = await storage.updateStaff(staffId, parseResult.data);
      res.json({ success: true, data: updatedStaff });
    } catch (error) {
      console.error('Error updating staff:', error);
      res.status(500).json({ success: false, error: 'Failed to update staff member' });
    }
  });

  // Delete staff member
  app.delete('/api/staff/:id', async (req, res) => {
    try {
      const staffId = parseInt(req.params.id);
      if (isNaN(staffId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID' });
      }

      await storage.deleteStaff(staffId);
      res.json({ success: true, message: 'Staff member deleted successfully' });
    } catch (error) {
      console.error('Error deleting staff:', error);
      res.status(500).json({ success: false, error: 'Failed to delete staff member' });
    }
  });

  // Sync staff from Microsoft Graph API (specific group or all users)
  app.post('/api/staff/sync-sharepoint', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      
      console.log('Syncing staff from Microsoft Graph API...');
      
      let azureUsers = [];
      
      // First, try to find the "Wolfpack" group
      console.log('Searching for Wolfpack group...');
      const groupsResponse = await fetch("https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Wolfpack'&$select=id,displayName", {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        }
      });

      if (groupsResponse.ok) {
        const groupsData = await groupsResponse.json();
        const wolfpackGroup = groupsData.value && groupsData.value[0];
        
        if (wolfpackGroup) {
          console.log(`Found Wolfpack group: ${wolfpackGroup.displayName} (${wolfpackGroup.id})`);
          
          // Get members of the Wolfpack group
          const membersResponse = await fetch(`https://graph.microsoft.com/v1.0/groups/${wolfpackGroup.id}/members?$select=id,displayName,mail,jobTitle,employeeHireDate,accountEnabled`, {
            method: 'GET',
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Accept': 'application/json',
              'Content-Type': 'application/json'
            }
          });

          if (membersResponse.ok) {
            const membersData = await membersResponse.json();
            azureUsers = membersData.value || [];
            console.log(`Found ${azureUsers.length} members in Wolfpack group`);
          } else {
            console.error(`Failed to fetch Wolfpack group members: ${membersResponse.status}`);
            throw new Error(`Failed to fetch Wolfpack group members: ${membersResponse.status}`);
          }
        } else {
          console.log('Wolfpack group not found, falling back to all active users');
          
          // Fallback: Get all users from Azure AD
          const allUsersResponse = await fetch('https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,jobTitle,employeeHireDate,accountEnabled&$filter=accountEnabled eq true', {
            method: 'GET',
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Accept': 'application/json',
              'Content-Type': 'application/json'
            }
          });

          if (!allUsersResponse.ok) {
            const errorText = await allUsersResponse.text();
            console.error(`Graph API error: ${allUsersResponse.status} ${allUsersResponse.statusText}`, errorText);
            throw new Error(`Failed to fetch users from Microsoft Graph: ${allUsersResponse.status} ${allUsersResponse.statusText}`);
          }

          const allUsersData = await allUsersResponse.json();
          azureUsers = allUsersData.value || [];
          console.log(`Found ${azureUsers.length} active users in Azure AD`);
        }
      } else {
        console.error(`Failed to search for groups: ${groupsResponse.status}`);
        throw new Error(`Failed to search for Azure AD groups: ${groupsResponse.status}`);
      }
      
      let created = 0;
      let updated = 0;
      let filtered = 0;
      let deactivated = 0;
      let errors = 0;
      const syncedEmails = new Set<string>();
      
      for (const user of azureUsers) {
        try {
          // Extract user data from Azure AD
          const azureAdObjectId = user.id;
          const staffName = user.displayName || 'Unknown';
          const staffEmail = user.mail || user.userPrincipalName || '';
          const jobTitle = user.jobTitle || null;
          const startDate = user.employeeHireDate ? new Date(user.employeeHireDate) : null;
          const isActive = user.accountEnabled !== false;
          
          if (!staffEmail || !staffEmail.includes('@')) {
            console.log(`Filtering user ${staffName} - no valid email found`);
            filtered++;
            continue;
          }
          
          // Filter out non-human accounts and system entries
          // Only include users who have proper job titles and are actual staff members
          if (!jobTitle || jobTitle.trim() === '') {
            console.log(`Filtering user ${staffName} - no job title (system account)`);
            filtered++;
            continue;
          }
          
          // Filter out system accounts, distribution lists, and service accounts
          const nameToCheck = staffName.toLowerCase();
          const emailToCheck = staffEmail.toLowerCase();
          
          // Skip entries that look like system accounts or distribution lists
          const systemPatterns = [
            'interview', 'invitation', 'form', 'contact', 'website', 'orders', 'accounts', 
            'dispatch', 'afterhours', 'rolleston info', 'metro property', 'next steps',
            'permits', 'christchurchsales', 'wanaka', 'cranfield glass –'
          ];
          
          const isSystemAccount = systemPatterns.some(pattern => 
            nameToCheck.includes(pattern) || emailToCheck.includes(pattern)
          );
          
          if (isSystemAccount) {
            console.log(`Filtering system account: ${staffName} (${staffEmail})`);
            filtered++;
            continue;
          }
          
          // Skip if email looks like a distribution list or service account
          if (emailToCheck.includes('noreply') || emailToCheck.includes('donotreply') ||
              emailToCheck.includes('admin@') || emailToCheck.includes('info@') ||
              emailToCheck.includes('support@') || emailToCheck.includes('accounts@') ||
              emailToCheck.includes('orders@') || emailToCheck.includes('sales@')) {
            console.log(`Filtering service email: ${staffName} (${staffEmail})`);
            filtered++;
            continue;
          }
          
          // Check if staff member already exists by email or Azure AD Object ID
          let existingStaff = await storage.getStaffByEmail(staffEmail);
          
          if (!existingStaff && azureAdObjectId) {
            // Also check by Azure AD Object ID
            try {
              const staffByAzureId = await storage.getStaffByAzureId(azureAdObjectId);
              if (staffByAzureId) {
                existingStaff = staffByAzureId;
              }
            } catch (e) {
              // Ignore if method doesn't exist yet
            }
          }
          
          if (existingStaff) {
            // Update existing staff member
            await storage.updateStaff(existingStaff.id, {
              azureAdObjectId: azureAdObjectId,
              name: staffName,
              email: staffEmail,
              jobTitle: jobTitle,
              isActive: isActive
            });
            updated++;
          } else {
            // Create new staff member
            await storage.createStaff({
              azureAdObjectId: azureAdObjectId,
              name: staffName,
              email: staffEmail,
              jobTitle: jobTitle,
              startDate: startDate,
              isActive: isActive
            });
            created++;
          }
          // Track this email as synced (normalised to lowercase)
          if (staffEmail) syncedEmails.add(staffEmail.toLowerCase());
        } catch (error) {
          console.error(`Error processing user ${user.displayName}:`, error);
          errors++;
        }
      }

      // Deactivate any staff who are no longer in the Wolfpack group
      if (syncedEmails.size > 0) {
        const allCurrentStaff = await storage.getAllStaff();
        for (const staffMember of allCurrentStaff) {
          if (staffMember.isActive && staffMember.email && !syncedEmails.has(staffMember.email.toLowerCase())) {
            console.log(`Deactivating staff member no longer in Wolfpack: ${staffMember.name} (${staffMember.email})`);
            await storage.updateStaff(staffMember.id, { isActive: false });
            deactivated++;
          }
        }
      }
      
      const totalProcessed = created + updated + filtered + deactivated + errors;
      res.json({
        success: true,
        message: `Microsoft Graph sync completed: ${created} staff created, ${updated} updated, ${deactivated} deactivated (removed from group), ${filtered} filtered (system accounts), ${errors} errors`,
        stats: { created, updated, filtered, deactivated, errors, totalProcessed }
      });
    } catch (error) {
      console.error('Error syncing staff from Microsoft Graph:', error);
      res.status(500).json({ 
        success: false, 
        error: error instanceof Error ? error.message : 'Failed to sync staff from Microsoft Graph' 
      });
    }
  });

  // Clean up problematic staff entries - remove staff without proper job titles
  app.post('/api/staff/cleanup', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }
      
      console.log('Cleaning up problematic staff entries...');
      
      // Get all staff
      const allStaff = await storage.getAllStaff();
      
      let deletedCount = 0;
      
      for (const staff of allStaff) {
        const shouldDelete = 
          !staff.jobTitle || 
          staff.jobTitle.trim() === '' ||
          // Check if name looks like a system account
          ['interview', 'invitation', 'form', 'contact', 'website', 'orders', 'accounts', 
           'dispatch', 'afterhours', 'rolleston info', 'metro property', 'next steps',
           'permits', 'christchurchsales', 'wanaka', 'cranfield glass –'].some(pattern => 
             staff.name.toLowerCase().includes(pattern)
           ) ||
          // Check if email looks like a service account
          (staff.email && (
            staff.email.toLowerCase().includes('noreply') ||
            staff.email.toLowerCase().includes('donotreply') ||
            staff.email.toLowerCase().includes('admin@') ||
            staff.email.toLowerCase().includes('info@') ||
            staff.email.toLowerCase().includes('support@') ||
            staff.email.toLowerCase().includes('accounts@') ||
            staff.email.toLowerCase().includes('orders@') ||
            staff.email.toLowerCase().includes('sales@')
          ));
          
        if (shouldDelete) {
          console.log(`Deleting problematic staff entry: ${staff.name} (${staff.email})`);
          await storage.deleteStaff(staff.id);
          deletedCount++;
        }
      }
      
      res.json({
        success: true,
        message: `Cleanup completed: ${deletedCount} problematic staff entries removed`,
        deletedCount
      });
    } catch (error) {
      console.error('Error cleaning up staff:', error);
      res.status(500).json({ 
        success: false, 
        error: error instanceof Error ? error.message : 'Failed to cleanup staff entries' 
      });
    }
  });

  // SKILLS MANAGEMENT ROUTES
  // -----------------------

  // Get all skills
  app.get('/api/skills', async (req, res) => {
    try {
      const { category } = req.query;
      
      const skills = category && typeof category === 'string' 
        ? await storage.getSkillsByCategory(category)
        : await storage.getAllSkills();
        
      res.json({ success: true, data: skills });
    } catch (error) {
      console.error('Error fetching skills:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch skills' });
    }
  });

  // Create new skill
  app.post('/api/skills', async (req, res) => {
    try {
      const parseResult = insertSkillSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid skill data: ${parseResult.error.message}`
        });
      }

      const newSkill = await storage.createSkill(parseResult.data);
      res.status(201).json({ success: true, data: newSkill });
    } catch (error) {
      console.error('Error creating skill:', error);
      res.status(500).json({ success: false, error: 'Failed to create skill' });
    }
  });

  // Update an existing skill
  app.put('/api/skills/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid skill ID' });
      }

      const updateData = req.body;
      const updatedSkill = await storage.updateSkill(id, updateData);
      
      if (!updatedSkill) {
        return res.status(404).json({ success: false, error: 'Skill not found' });
      }

      res.json({ success: true, data: updatedSkill });
    } catch (error) {
      console.error('Error updating skill:', error);
      res.status(500).json({ success: false, error: 'Failed to update skill' });
    }
  });

  // TRAINING RECORDS ROUTES - Core Skills Matrix
  // -------------------------------------------

  // Get training matrix (pivoted view of staff vs skills)
  app.get('/api/training-matrix', async (req, res) => {
    try {
      const matrix = await storage.getTrainingMatrix();
      res.json({ success: true, data: matrix });
    } catch (error) {
      console.error('Error fetching training matrix:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch training matrix' });
    }
  });

  // Get training records by staff
  app.get('/api/training-records/staff/:staffId', async (req, res) => {
    try {
      const staffId = parseInt(req.params.staffId);
      if (isNaN(staffId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID' });
      }

      const records = await storage.getTrainingRecordsByStaff(staffId);
      res.json({ success: true, data: records });
    } catch (error) {
      console.error('Error fetching training records:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch training records' });
    }
  });

  // Create training record
  app.post('/api/training-records', async (req, res) => {
    try {
      const parseResult = insertTrainingRecordSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid training record: ${parseResult.error.message}`
        });
      }

      const newRecord = await storage.createTrainingRecord(parseResult.data);
      res.status(201).json({ success: true, data: newRecord });
    } catch (error) {
      console.error('Error creating training record:', error);
      res.status(500).json({ success: false, error: 'Failed to create training record' });
    }
  });

  // Update training record
  app.put('/api/training-records/:id', async (req, res) => {
    try {
      const recordId = parseInt(req.params.id);
      if (isNaN(recordId)) {
        return res.status(400).json({ success: false, error: 'Invalid record ID' });
      }

      const parseResult = insertTrainingRecordSchema.partial().safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid training record: ${parseResult.error.message}`
        });
      }

      const updatedRecord = await storage.updateTrainingRecord(recordId, parseResult.data);
      res.json({ success: true, data: updatedRecord });
    } catch (error) {
      console.error('Error updating training record:', error);
      res.status(500).json({ success: false, error: 'Failed to update training record' });
    }
  });

  // Bulk compliance update - set all field staff modules to passing with dynamic review date
  app.post('/api/training-records/bulk-compliance-update', async (req, res) => {
    try {
      // Get expiry date from request body, default to Feb 2026
      const requestedExpiryDate = req.body?.expiryDate || '2026-02-28';
      
      // Get all active staff members
      const allStaff = await storage.getAllStaff();
      const fieldStaff = allStaff.filter(s => s.isFieldStaff && s.isActive);
      console.log(`Found ${fieldStaff.length} field staff`);

      // Get all field training modules
      const allModules = await storage.getAllTrainingModules();
      const fieldModules = allModules.filter(m => 
        m.isActive && (m.audience === 'field' || m.audience === 'both')
      );
      console.log(`Found ${fieldModules.length} field modules`);

      // Set expiry date from request
      const expiryDate = new Date(requestedExpiryDate);
      const achievedDate = new Date(); // Today
      
      const recordsProcessed = [];
      const errors = [];

      // Create records for each staff-module combination
      // The createTrainingRecord method automatically handles versioning
      for (const staff of fieldStaff) {
        for (const module of fieldModules) {
          try {
            const recordData: any = {
              staffId: staff.id,
              skillId: module.id,
              competencyLevel: 'Competent – SOP/Module',
              achievedDate: achievedDate,
              expiryDate: expiryDate,
              ableToUse: true,
              status: 'Active',
              assessorName: 'Compliance Update',
              notes: `Temporary compliance pass - Due for review ${expiryDate.toLocaleDateString('en-NZ', { month: 'long', year: 'numeric' })}`
            };

            // Always create - the storage method handles superseding old Active records
            const newRecord = await storage.createTrainingRecord(recordData);
            recordsProcessed.push(newRecord);
          } catch (error) {
            console.error(`Error processing staff ${staff.id}, module ${module.id}:`, error);
            errors.push({
              staffId: staff.id,
              staffName: staff.name,
              moduleId: module.id,
              moduleName: module.name,
              error: error instanceof Error ? error.message : 'Unknown error'
            });
          }
        }
      }

      console.log(`Bulk update complete: ${recordsProcessed.length} records processed, ${errors.length} errors`);

      res.json({ 
        success: true, 
        data: {
          recordsProcessed: recordsProcessed.length,
          totalStaff: fieldStaff.length,
          totalModules: fieldModules.length,
          expiryDate: expiryDate.toISOString(),
          errors: errors.length > 0 ? errors : undefined
        }
      });
    } catch (error) {
      console.error('Error in bulk compliance update:', error);
      res.status(500).json({ 
        success: false, 
        error: 'Failed to perform bulk compliance update',
        details: error instanceof Error ? error.message : 'Unknown error'
      });
    }
  });

  // "ABLE TO USE" TRAINING SYSTEM API ROUTES
  // ========================================

  // Training Classification routes
  app.get('/api/training-classifications', async (req, res) => {
    try {
      const audience = req.query.audience as string;
      const classifications = await storage.getAllTrainingClassifications(audience);
      res.json({ success: true, data: classifications });
    } catch (error) {
      console.error('Error fetching training classifications:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch training classifications' });
    }
  });

  app.get('/api/training-classifications/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid classification ID' });
      }
      
      const classification = await storage.getTrainingClassification(id);
      if (!classification) {
        return res.status(404).json({ success: false, error: 'Training classification not found' });
      }
      
      res.json({ success: true, data: classification });
    } catch (error) {
      console.error('Error fetching training classification:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch training classification' });
    }
  });

  app.post('/api/training-classifications', async (req, res) => {
    try {
      const parseResult = insertTrainingClassificationSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid training classification: ${parseResult.error.message}`
        });
      }

      const newClassification = await storage.createTrainingClassification(parseResult.data);
      res.status(201).json({ success: true, data: newClassification });
    } catch (error) {
      console.error('Error creating training classification:', error);
      res.status(500).json({ success: false, error: 'Failed to create training classification' });
    }
  });

  app.put('/api/training-classifications/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid classification ID' });
      }

      const parseResult = insertTrainingClassificationSchema.partial().safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid training classification: ${parseResult.error.message}`
        });
      }

      const updatedClassification = await storage.updateTrainingClassification(id, parseResult.data);
      res.json({ success: true, data: updatedClassification });
    } catch (error) {
      console.error('Error updating training classification:', error);
      res.status(500).json({ success: false, error: 'Failed to update training classification' });
    }
  });

  // Training Module routes
  app.get('/api/training-modules', async (req, res) => {
    try {
      const classificationId = req.query.classificationId ? parseInt(req.query.classificationId as string) : undefined;
      
      const modules = classificationId
        ? await storage.getTrainingModulesByClassification(classificationId)
        : await storage.getAllTrainingModules();
        
      res.json({ success: true, data: modules });
    } catch (error) {
      console.error('Error fetching training modules:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch training modules' });
    }
  });

  app.get('/api/training-modules/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid module ID' });
      }
      
      const module = await storage.getTrainingModule(id);
      if (!module) {
        return res.status(404).json({ success: false, error: 'Training module not found' });
      }
      
      res.json({ success: true, data: module });
    } catch (error) {
      console.error('Error fetching training module:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch training module' });
    }
  });

  app.post('/api/training-modules', async (req, res) => {
    try {
      const parseResult = insertTrainingModuleSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid training module: ${parseResult.error.message}`
        });
      }

      const newModule = await storage.createTrainingModule(parseResult.data);
      res.status(201).json({ success: true, data: newModule });
    } catch (error) {
      console.error('Error creating training module:', error);
      res.status(500).json({ success: false, error: 'Failed to create training module' });
    }
  });

  app.put('/api/training-modules/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid module ID' });
      }

      const parseResult = insertTrainingModuleSchema.partial().safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid training module: ${parseResult.error.message}`
        });
      }

      const updatedModule = await storage.updateTrainingModule(id, parseResult.data);
      res.json({ success: true, data: updatedModule });
    } catch (error) {
      console.error('Error updating training module:', error);
      res.status(500).json({ success: false, error: 'Failed to update training module' });
    }
  });

  // Module Prerequisites routes
  app.get('/api/training-modules/:moduleId/prerequisites', async (req, res) => {
    try {
      const moduleId = parseInt(req.params.moduleId);
      if (isNaN(moduleId)) {
        return res.status(400).json({ success: false, error: 'Invalid module ID' });
      }
      
      const prerequisites = await storage.getModulePrerequisites(moduleId);
      res.json({ success: true, data: prerequisites });
    } catch (error) {
      console.error('Error fetching module prerequisites:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch module prerequisites' });
    }
  });

  app.post('/api/module-prerequisites', async (req, res) => {
    try {
      const parseResult = insertModulePrerequisiteSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid module prerequisite: ${parseResult.error.message}`
        });
      }

      const newPrerequisite = await storage.createModulePrerequisite(parseResult.data);
      res.status(201).json({ success: true, data: newPrerequisite });
    } catch (error) {
      console.error('Error creating module prerequisite:', error);
      res.status(500).json({ success: false, error: 'Failed to create module prerequisite' });
    }
  });

  app.delete('/api/module-prerequisites/:moduleId/:prerequisiteModuleId', async (req, res) => {
    try {
      const moduleId = parseInt(req.params.moduleId);
      const prerequisiteModuleId = parseInt(req.params.prerequisiteModuleId);
      
      if (isNaN(moduleId) || isNaN(prerequisiteModuleId)) {
        return res.status(400).json({ success: false, error: 'Invalid module IDs' });
      }
      
      await storage.deleteModulePrerequisite(moduleId, prerequisiteModuleId);
      res.json({ success: true, message: 'Module prerequisite deleted successfully' });
    } catch (error) {
      console.error('Error deleting module prerequisite:', error);
      res.status(500).json({ success: false, error: 'Failed to delete module prerequisite' });
    }
  });

  // Module Tools routes
  app.get('/api/training-modules/:moduleId/tools', async (req, res) => {
    try {
      const moduleId = parseInt(req.params.moduleId);
      if (isNaN(moduleId)) {
        return res.status(400).json({ success: false, error: 'Invalid module ID' });
      }
      
      const tools = await storage.getModuleTools(moduleId);
      res.json({ success: true, data: tools });
    } catch (error) {
      console.error('Error fetching module tools:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch module tools' });
    }
  });

  app.post('/api/module-tools', async (req, res) => {
    try {
      const parseResult = insertModuleToolSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid module tool: ${parseResult.error.message}`
        });
      }

      const newTool = await storage.createModuleTool(parseResult.data);
      res.status(201).json({ success: true, data: newTool });
    } catch (error) {
      console.error('Error creating module tool:', error);
      res.status(500).json({ success: false, error: 'Failed to create module tool' });
    }
  });

  app.put('/api/module-tools/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid tool ID' });
      }

      const parseResult = insertModuleToolSchema.partial().safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid module tool: ${parseResult.error.message}`
        });
      }

      const updatedTool = await storage.updateModuleTool(id, parseResult.data);
      res.json({ success: true, data: updatedTool });
    } catch (error) {
      console.error('Error updating module tool:', error);
      res.status(500).json({ success: false, error: 'Failed to update module tool' });
    }
  });

  app.delete('/api/module-tools/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid tool ID' });
      }
      
      await storage.deleteModuleTool(id);
      res.json({ success: true, message: 'Module tool deleted successfully' });
    } catch (error) {
      console.error('Error deleting module tool:', error);
      res.status(500).json({ success: false, error: 'Failed to delete module tool' });
    }
  });

  // Module Materials routes
  app.get('/api/training-modules/:moduleId/materials', async (req, res) => {
    try {
      const moduleId = parseInt(req.params.moduleId);
      if (isNaN(moduleId)) {
        return res.status(400).json({ success: false, error: 'Invalid module ID' });
      }
      
      const materials = await storage.getModuleMaterials(moduleId);
      res.json({ success: true, data: materials });
    } catch (error) {
      console.error('Error fetching module materials:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch module materials' });
    }
  });

  app.post('/api/module-materials', async (req, res) => {
    try {
      const parseResult = insertModuleMaterialSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid module material: ${parseResult.error.message}`
        });
      }

      const newMaterial = await storage.createModuleMaterial(parseResult.data);
      res.status(201).json({ success: true, data: newMaterial });
    } catch (error) {
      console.error('Error creating module material:', error);
      res.status(500).json({ success: false, error: 'Failed to create module material' });
    }
  });

  app.put('/api/module-materials/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid material ID' });
      }

      const parseResult = insertModuleMaterialSchema.partial().safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid module material: ${parseResult.error.message}`
        });
      }

      const updatedMaterial = await storage.updateModuleMaterial(id, parseResult.data);
      res.json({ success: true, data: updatedMaterial });
    } catch (error) {
      console.error('Error updating module material:', error);
      res.status(500).json({ success: false, error: 'Failed to update module material' });
    }
  });

  app.delete('/api/module-materials/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid material ID' });
      }
      
      await storage.deleteModuleMaterial(id);
      res.json({ success: true, message: 'Module material deleted successfully' });
    } catch (error) {
      console.error('Error deleting module material:', error);
      res.status(500).json({ success: false, error: 'Failed to delete module material' });
    }
  });

  // Staff Module Progress routes
  app.get('/api/staff-module-progress', async (req, res) => {
    try {
      const staffId = req.query.staffId ? parseInt(req.query.staffId as string) : undefined;
      const moduleId = req.query.moduleId ? parseInt(req.query.moduleId as string) : undefined;
      
      let progress;
      if (staffId && moduleId) {
        progress = await storage.getModuleProgressByStaffAndModule(staffId, moduleId);
        progress = progress ? [progress] : [];
      } else if (staffId) {
        progress = await storage.getStaffModuleProgress(staffId);
      } else if (moduleId) {
        progress = await storage.getModuleProgress(moduleId);
      } else {
        progress = await storage.getAllStaffModuleProgress();
      }
      
      res.json({ success: true, data: progress });
    } catch (error) {
      console.error('Error fetching staff module progress:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch staff module progress' });
    }
  });

  app.post('/api/staff-module-progress', async (req, res) => {
    try {
      const parseResult = insertStaffModuleProgressSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid staff module progress: ${parseResult.error.message}`
        });
      }

      const newProgress = await storage.createStaffModuleProgress(parseResult.data);
      res.status(201).json({ success: true, data: newProgress });
    } catch (error) {
      console.error('Error creating staff module progress:', error);
      res.status(500).json({ success: false, error: 'Failed to create staff module progress' });
    }
  });

  app.put('/api/staff-module-progress/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ success: false, error: 'Invalid progress ID' });
      }

      const parseResult = insertStaffModuleProgressSchema.partial().safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid staff module progress: ${parseResult.error.message}`
        });
      }

      const updatedProgress = await storage.updateStaffModuleProgress(id, parseResult.data);
      res.json({ success: true, data: updatedProgress });
    } catch (error) {
      console.error('Error updating staff module progress:', error);
      res.status(500).json({ success: false, error: 'Failed to update staff module progress' });
    }
  });

  // Training Module Matrix (replacement for the training matrix)
  app.get('/api/training-module-matrix', async (req, res) => {
    try {
      const matrix = await storage.getTrainingModuleMatrix();
      res.json({ success: true, data: matrix });
    } catch (error) {
      console.error('Error fetching training module matrix:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch training module matrix' });
    }
  });

  // Check "Able to Use" status for staff member and module
  app.get('/api/staff/:staffId/able-to-use/:moduleId', async (req, res) => {
    try {
      const staffId = parseInt(req.params.staffId);
      const moduleId = parseInt(req.params.moduleId);
      
      if (isNaN(staffId) || isNaN(moduleId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID or module ID' });
      }

      // Get staff module progress
      const progress = await storage.getModuleProgressByStaffAndModule(staffId, moduleId);
      if (!progress) {
        return res.json({ 
          success: true, 
          data: { 
            ableToUse: false, 
            reason: 'No training record found',
            competencyLevel: 'Not Trained'
          } 
        });
      }

      // Get module details to check safety-critical status
      const module = await storage.getTrainingModule(moduleId);
      if (!module) {
        return res.status(404).json({ success: false, error: 'Module not found' });
      }

      // Import competency validation function
      const { isAbleToUse } = await import("@shared/schema");
      
      // Check if competency level allows "Able to Use"
      const canUse = isAbleToUse(progress.competencyLevel);
      let reason = '';
      
      if (!canUse) {
        if (progress.competencyLevel === 'Not Trained') {
          reason = 'Staff member has not received training for this module';
        } else if (progress.competencyLevel === 'In Training (Supervised)') {
          reason = 'Staff member is still in training and requires supervision';
        } else if (progress.competencyLevel === 'Competent – Supervised') {
          reason = 'Staff member requires supervision - no SOP/module training completed';
        } else {
          reason = `Competency level "${progress.competencyLevel}" does not allow independent use`;
        }
      }

      // For safety-critical modules, check additional requirements
      if (canUse && module.isSafetyCritical) {
        if (!progress.trainedAgainstSop) {
          return res.json({
            success: true,
            data: {
              ableToUse: false,
              reason: 'Safety-critical module requires SOP/module training',
              competencyLevel: progress.competencyLevel,
              isSafetyCritical: true
            }
          });
        }
        
        // Check expiry for safety-critical modules
        if (progress.expiryDate && new Date(progress.expiryDate) < new Date()) {
          return res.json({
            success: true,
            data: {
              ableToUse: false,
              reason: 'Training has expired and requires renewal',
              competencyLevel: progress.competencyLevel,
              isSafetyCritical: true,
              expired: true
            }
          });
        }
      }

      res.json({
        success: true,
        data: {
          ableToUse: canUse,
          reason: canUse ? 'Staff member is authorized for independent use' : reason,
          competencyLevel: progress.competencyLevel,
          isSafetyCritical: module.isSafetyCritical,
          trainedAgainstSop: progress.trainedAgainstSop,
          authorizedDate: progress.authorizedDate,
          expiryDate: progress.expiryDate
        }
      });
    } catch (error) {
      console.error('Error checking able to use status:', error);
      res.status(500).json({ success: false, error: 'Failed to check able to use status' });
    }
  });

  // PPE REGISTER ROUTES
  // ------------------

  // Get PPE records by staff
  app.get('/api/ppe-records/staff/:staffId', async (req, res) => {
    try {
      const staffId = parseInt(req.params.staffId);
      if (isNaN(staffId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID' });
      }

      const records = await storage.getPpeRecordsByStaff(staffId);
      res.json({ success: true, data: records });
    } catch (error) {
      console.error('Error fetching PPE records:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch PPE records' });
    }
  });

  // Get expiring PPE
  app.get('/api/ppe-records/expiring', async (req, res) => {
    try {
      const withinDays = parseInt(req.query.days as string) || 30;
      const expiringPpe = await storage.getExpiringPpe(withinDays);
      res.json({ success: true, data: expiringPpe });
    } catch (error) {
      console.error('Error fetching expiring PPE:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch expiring PPE' });
    }
  });

  // Create PPE record
  app.post('/api/ppe-records', async (req, res) => {
    try {
      const parseResult = insertPpeRecordSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid PPE record: ${parseResult.error.message}`
        });
      }

      const newRecord = await storage.createPpeRecord(parseResult.data);
      res.status(201).json({ success: true, data: newRecord });
    } catch (error) {
      console.error('Error creating PPE record:', error);
      res.status(500).json({ success: false, error: 'Failed to create PPE record' });
    }
  });

  // INDUCTION REGISTER ROUTES
  // -------------------------

  // Get induction records by staff
  app.get('/api/inductions/staff/:staffId', async (req, res) => {
    try {
      const staffId = parseInt(req.params.staffId);
      if (isNaN(staffId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID' });
      }

      const records = await storage.getInductionRecordsByStaff(staffId);
      res.json({ success: true, data: records });
    } catch (error) {
      console.error('Error fetching induction records:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch induction records' });
    }
  });

  // Create induction record
  app.post('/api/inductions', async (req, res) => {
    try {
      const parseResult = insertInductionRecordSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid induction record: ${parseResult.error.message}`
        });
      }

      const newRecord = await storage.createInductionRecord(parseResult.data);
      res.status(201).json({ success: true, data: newRecord });
    } catch (error) {
      console.error('Error creating induction record:', error);
      res.status(500).json({ success: false, error: 'Failed to create induction record' });
    }
  });

  // EQUIPMENT AUTHORIZATION ROUTES ("Able to Use" Matrix)
  // ----------------------------------------------------

  // Get equipment authorizations by staff
  app.get('/api/equipment-auth/staff/:staffId', async (req, res) => {
    try {
      const staffId = parseInt(req.params.staffId);
      if (isNaN(staffId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID' });
      }

      const auths = await storage.getEquipmentAuthorizationsByStaff(staffId);
      res.json({ success: true, data: auths });
    } catch (error) {
      console.error('Error fetching equipment authorizations:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch equipment authorizations' });
    }
  });

  // Create equipment authorization
  app.post('/api/equipment-auth', async (req, res) => {
    try {
      const parseResult = insertEquipmentAuthorizationSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid equipment authorization: ${parseResult.error.message}`
        });
      }

      const newAuth = await storage.createEquipmentAuthorization(parseResult.data);
      res.status(201).json({ success: true, data: newAuth });
    } catch (error) {
      console.error('Error creating equipment authorization:', error);
      res.status(500).json({ success: false, error: 'Failed to create equipment authorization' });
    }
  });

  // PHOTO MANAGEMENT ROUTES
  // ----------------------

  // ============================================================================
  // STAFF PHOTO MANAGEMENT - SharePoint Integration for ID Photos
  // ============================================================================

  // Helper function to resolve a SharePoint site by its URL/path
  async function getSiteInfo(accessToken: string, siteUrl: string) {
    const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteUrl}`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json'
      }
    });

    if (!siteResponse.ok) {
      const errorText = await siteResponse.text();
      throw new Error(`Failed to resolve SharePoint site: ${siteResponse.status} - ${errorText}`);
    }

    return await siteResponse.json();
  }

  // Helper function to upload photo to SharePoint document library
  async function uploadPhotoToSharePoint(
    accessToken: string,
    siteUrl: string,
    folderPath: string,
    fileName: string,
    fileBuffer: Buffer,
    mimeType: string
  ) {
    // Get site info first
    const siteInfo = await getSiteInfo(accessToken, siteUrl);
    
    // Create folder if it doesn't exist
    await ensureSharePointFolder(accessToken, siteInfo.id, folderPath);
    
    // Upload file to SharePoint
    const uploadUrl = folderPath 
      ? `https://graph.microsoft.com/v1.0/sites/${siteInfo.id}/drive/root:/${encodeURIComponent(folderPath)}/${encodeURIComponent(fileName)}:/content`
      : `https://graph.microsoft.com/v1.0/sites/${siteInfo.id}/drive/root:/${encodeURIComponent(fileName)}:/content`;

    const uploadResponse = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': mimeType,
        'Content-Length': fileBuffer.length.toString()
      },
      body: fileBuffer
    });

    if (!uploadResponse.ok) {
      const errorText = await uploadResponse.text();
      throw new Error(`Failed to upload photo to SharePoint: ${uploadResponse.status} - ${errorText}`);
    }

    const uploadResult = await uploadResponse.json();
    
    return {
      siteId: siteInfo.id,
      fileId: uploadResult.id,
      webUrl: uploadResult.webUrl,
      downloadUrl: uploadResult['@microsoft.graph.downloadUrl'],
      fileName: uploadResult.name,
      size: uploadResult.size
    };
  }

  // Helper function to ensure SharePoint folder exists
  async function ensureSharePointFolder(accessToken: string, siteId: string, folderPath: string) {
    if (!folderPath) return;

    const checkUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeURIComponent(folderPath)}`;
    
    const checkResponse = await fetch(checkUrl, {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json'
      }
    });

    if (checkResponse.status === 404) {
      // Create folder
      const createUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children`;
      
      const createResponse = await fetch(createUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          name: folderPath,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'replace'
        })
      });

      if (!createResponse.ok) {
        const errorText = await createResponse.text();
        console.warn(`Failed to create folder ${folderPath}: ${errorText}`);
      }
    }
  }

  // Helper function to generate photo thumbnail
  async function generatePhotoThumbnail(accessToken: string, siteId: string, fileId: string): Promise<string | null> {
    try {
      // Get thumbnail from SharePoint
      const thumbnailUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/thumbnails/0/medium`;
      
      const thumbnailResponse = await fetch(thumbnailUrl, {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Accept': 'application/json'
        }
      });

      if (thumbnailResponse.ok) {
        const thumbnailData = await thumbnailResponse.json();
        return thumbnailData.url || null;
      }
      
      console.warn('Failed to generate thumbnail for photo:', thumbnailResponse.status);
      return null;
    } catch (error) {
      console.warn('Error generating thumbnail:', error);
      return null;
    }
  }

  // Helper function to delete photo from SharePoint
  async function deletePhotoFromSharePoint(accessToken: string, fileId: string) {
    const deleteUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`;
    
    const deleteResponse = await fetch(deleteUrl, {
      method: 'DELETE',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    if (!deleteResponse.ok && deleteResponse.status !== 404) {
      const errorText = await deleteResponse.text();
      throw new Error(`Failed to delete photo from SharePoint: ${deleteResponse.status} - ${errorText}`);
    }
  }

  // Get staff photo
  app.get('/api/photos/staff/:staffId', async (req, res) => {
    try {
      const staffId = parseInt(req.params.staffId);
      if (isNaN(staffId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID' });
      }

      const photo = await storage.getPhotoAssetByStaff(staffId);
      res.json({ success: true, data: photo });
    } catch (error) {
      console.error('Error fetching staff photo:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch staff photo' });
    }
  });

  // Upload staff photo to SharePoint and save metadata
  app.post('/api/photos/upload', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { staffId, fileName, fileData, mimeType, site, folder } = req.body;

      if (!staffId || !fileName || !fileData || !mimeType) {
        return res.status(400).json({
          success: false,
          error: 'staffId, fileName, fileData, and mimeType are required'
        });
      }

      // Convert base64 to buffer
      const buffer = Buffer.from(fileData, 'base64');
      
      // Upload to SharePoint using existing infrastructure
      const siteUrl = `cranfieldglass.sharepoint.com:/sites/${site || 'CranfieldGlass'}:`;
      const folderPath = folder || 'Staff Photos';
      
      const uploadResult = await uploadPhotoToSharePoint(
        accessToken,
        siteUrl,
        folderPath,
        fileName,
        buffer,
        mimeType
      );

      // Generate thumbnail
      const thumbnailUrl = await generatePhotoThumbnail(
        accessToken,
        uploadResult.siteId,
        uploadResult.fileId
      );

      // Save metadata to database
      const photoAssetData = {
        staffId: parseInt(staffId),
        filename: fileName,
        sharePointDriveItemId: uploadResult.fileId,
        sharePointWebUrl: uploadResult.webUrl,
        thumbnailUrl: thumbnailUrl,
        fileSize: buffer.length,
        mimeType: mimeType,
        uploadedBy: 'current-user', // TODO: Get from auth context
        isActive: true
      };

      const newPhoto = await storage.createPhotoAsset(photoAssetData);
      res.status(201).json({ success: true, data: newPhoto });
    } catch (error) {
      console.error('Error uploading staff photo:', error);
      res.status(500).json({ 
        success: false, 
        error: error instanceof Error ? error.message : 'Failed to upload photo' 
      });
    }
  });

  // Update staff photo metadata
  app.put('/api/photos/:photoId', async (req, res) => {
    try {
      const photoId = parseInt(req.params.photoId);
      if (isNaN(photoId)) {
        return res.status(400).json({ success: false, error: 'Invalid photo ID' });
      }

      const parseResult = insertPhotoAssetSchema.partial().safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid photo data: ${parseResult.error.message}`
        });
      }

      const updatedPhoto = await storage.updatePhotoAsset(photoId, parseResult.data);
      res.json({ success: true, data: updatedPhoto });
    } catch (error) {
      console.error('Error updating photo:', error);
      res.status(500).json({ success: false, error: 'Failed to update photo' });
    }
  });

  // Delete staff photo
  app.delete('/api/photos/:photoId', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const photoId = parseInt(req.params.photoId);
      
      if (isNaN(photoId)) {
        return res.status(400).json({ success: false, error: 'Invalid photo ID' });
      }

      // Get photo metadata
      const photo = await storage.getPhotoAsset(photoId);
      if (!photo) {
        return res.status(404).json({ success: false, error: 'Photo not found' });
      }

      // Delete from SharePoint if fileId exists
      if (photo.sharePointDriveItemId) {
        try {
          await deletePhotoFromSharePoint(accessToken, photo.sharePointDriveItemId);
        } catch (error) {
          console.warn('Failed to delete photo from SharePoint:', error);
          // Continue with database deletion even if SharePoint deletion fails
        }
      }

      // Delete metadata from database
      await storage.deletePhotoAsset(photoId);
      res.json({ success: true, message: 'Photo deleted successfully' });
    } catch (error) {
      console.error('Error deleting photo:', error);
      res.status(500).json({ success: false, error: 'Failed to delete photo' });
    }
  });

  // Upload staff photo metadata only (for externally uploaded files)
  app.post('/api/photos', async (req, res) => {
    try {
      const parseResult = insertPhotoAssetSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid photo data: ${parseResult.error.message}`
        });
      }

      const newPhoto = await storage.createPhotoAsset(parseResult.data);
      res.status(201).json({ success: true, data: newPhoto });
    } catch (error) {
      console.error('Error saving photo metadata:', error);
      res.status(500).json({ success: false, error: 'Failed to save photo metadata' });
    }
  });

  // GLOVE REGISTER ROUTES
  // --------------------

  // Get glove records by staff
  app.get('/api/gloves/staff/:staffId', async (req, res) => {
    try {
      const staffId = parseInt(req.params.staffId);
      if (isNaN(staffId)) {
        return res.status(400).json({ success: false, error: 'Invalid staff ID' });
      }

      const records = await storage.getGloveRecordsByStaff(staffId);
      res.json({ success: true, data: records });
    } catch (error) {
      console.error('Error fetching glove records:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch glove records' });
    }
  });

  // Create glove record
  app.post('/api/gloves', async (req, res) => {
    try {
      const parseResult = insertGloveRecordSchema.safeParse(req.body);
      if (!parseResult.success) {
        return res.status(400).json({
          success: false,
          error: `Invalid glove record: ${parseResult.error.message}`
        });
      }

      const newRecord = await storage.createGloveRecord(parseResult.data);
      res.status(201).json({ success: true, data: newRecord });
    } catch (error) {
      console.error('Error creating glove record:', error);
      res.status(500).json({ success: false, error: 'Failed to create glove record' });
    }
  });

  // ============================================================================
  // EXCEL IMPORT WIZARD - Migrate from Excel-based Skills Matrix to Database
  // ============================================================================

  // Helper function to analyze column headers and suggest mappings
  function analyzeColumnMappings(headers: string[]) {
    const mappings: Array<{ column: string; suggestedField: string; confidence: number }> = [];
    
    for (const header of headers) {
      const lowerHeader = header.toLowerCase().trim();
      let suggestedField = '';
      let confidence = 0;

      // Staff/Name mappings
      if (lowerHeader.includes('name') || lowerHeader.includes('staff') || lowerHeader.includes('employee')) {
        suggestedField = 'staff.name';
        confidence = 0.9;
      }
      // Email mappings
      else if (lowerHeader.includes('email') || lowerHeader.includes('mail')) {
        suggestedField = 'staff.email';
        confidence = 0.9;
      }
      // Job title mappings
      else if (lowerHeader.includes('title') || lowerHeader.includes('position') || lowerHeader.includes('role')) {
        suggestedField = 'staff.jobTitle';
        confidence = 0.8;
      }
      // Start date mappings
      else if (lowerHeader.includes('start') && lowerHeader.includes('date')) {
        suggestedField = 'staff.startDate';
        confidence = 0.8;
      }
      // Competency level mappings
      else if (lowerHeader.includes('level') || lowerHeader.includes('competency') || lowerHeader.includes('grade')) {
        suggestedField = 'trainingRecord.competencyLevel';
        confidence = 0.7;
      }
      // Date-related mappings
      else if (lowerHeader.includes('achieved') || lowerHeader.includes('completion') || lowerHeader.includes('passed')) {
        suggestedField = 'trainingRecord.achievedDate';
        confidence = 0.7;
      }
      else if (lowerHeader.includes('expiry') || lowerHeader.includes('expires') || lowerHeader.includes('renewal')) {
        suggestedField = 'trainingRecord.expiryDate';
        confidence = 0.8;
      }
      // Assessor mappings
      else if (lowerHeader.includes('assessor') || lowerHeader.includes('trainer') || lowerHeader.includes('instructor')) {
        suggestedField = 'trainingRecord.assessorName';
        confidence = 0.7;
      }
      // Training provider mappings
      else if (lowerHeader.includes('provider') || lowerHeader.includes('training') && lowerHeader.includes('org')) {
        suggestedField = 'trainingRecord.trainingProvider';
        confidence = 0.7;
      }
      // Certificate mappings
      else if (lowerHeader.includes('certificate') || lowerHeader.includes('cert') || lowerHeader.includes('number')) {
        suggestedField = 'trainingRecord.certificateNumber';
        confidence = 0.6;
      }
      // PPE mappings
      else if (lowerHeader.includes('ppe') || lowerHeader.includes('equipment') || lowerHeader.includes('safety')) {
        suggestedField = 'ppeRecord.ppeType';
        confidence = 0.6;
      }
      // Status mappings
      else if (lowerHeader.includes('status') || lowerHeader.includes('active') || lowerHeader.includes('current')) {
        suggestedField = 'trainingRecord.status';
        confidence = 0.6;
      }
      // Skills - anything else might be a skill name
      else if (lowerHeader.length > 2 && !lowerHeader.includes('id')) {
        suggestedField = 'skill.name';
        confidence = 0.3;
      }

      if (suggestedField) {
        mappings.push({
          column: header,
          suggestedField,
          confidence
        });
      }
    }

    return mappings;
  }

  // Helper function to execute the skills matrix import
  async function executeSkillsMatrixImport(workbookData: any, mappings: any, options: any = {}) {
    const importResult = {
      success: true,
      summary: {
        staffCreated: 0,
        skillsCreated: 0,
        trainingRecordsCreated: 0,
        ppeRecordsCreated: 0,
        inductionsCreated: 0,
        equipmentAuthsCreated: 0,
        errors: [] as string[]
      },
      details: [] as any[]
    };

    try {
      // Process each worksheet
      for (const sheet of workbookData.sheets) {
        if (!mappings[sheet.name]) {
          console.log(`No mappings defined for worksheet: ${sheet.name}, skipping`);
          continue;
        }

        const sheetMappings = mappings[sheet.name];
        console.log(`Processing worksheet: ${sheet.name} with ${sheet.rows.length} rows`);

        // Create a map of column indices to field mappings
        const columnMap: Record<number, string> = {};
        for (const mapping of sheetMappings) {
          const columnIndex = sheet.headers.indexOf(mapping.column);
          if (columnIndex !== -1) {
            columnMap[columnIndex] = mapping.field;
          }
        }

        // Process each row
        for (let rowIndex = 0; rowIndex < sheet.rows.length; rowIndex++) {
          const row = sheet.rows[rowIndex];
          const rowData: Record<string, any> = {};

          // Extract data based on column mappings
          for (const [colIndex, fieldPath] of Object.entries(columnMap)) {
            const cellValue = row[parseInt(colIndex)];
            if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
              rowData[fieldPath] = cellValue;
            }
          }

          try {
            // Process staff data
            if (rowData['staff.name']) {
              const staffData = {
                name: String(rowData['staff.name']),
                email: rowData['staff.email'] ? String(rowData['staff.email']) : `${String(rowData['staff.name']).toLowerCase().replace(/\s+/g, '.')}@cranfieldglass.com`,
                jobTitle: rowData['staff.jobTitle'] ? String(rowData['staff.jobTitle']) : null,
                startDate: rowData['staff.startDate'] ? new Date(rowData['staff.startDate']) : null,
                isActive: true
              };

              // Check if staff already exists by name
              const existingStaff = await storage.getAllStaff();
              let staff = existingStaff.find(s => s.name.toLowerCase() === staffData.name.toLowerCase());
              
              if (!staff) {
                try {
                  staff = await storage.createStaff(staffData);
                  importResult.summary.staffCreated++;
                  console.log(`Created staff: ${staff.name}`);
                } catch (error) {
                  console.error(`Error creating staff ${staffData.name}:`, error);
                  importResult.summary.errors.push(`Failed to create staff ${staffData.name}: ${error instanceof Error ? error.message : 'Unknown error'}`);
                  continue;
                }
              }

              // Process skill data if this row represents a training record
              if (rowData['skill.name'] || Object.keys(rowData).some(key => key.startsWith('trainingRecord.'))) {
                let skill = null;
                
                if (rowData['skill.name']) {
                  const skillName = String(rowData['skill.name']);
                  const existingSkills = await storage.getAllSkills();
                  skill = existingSkills.find(s => s.name.toLowerCase() === skillName.toLowerCase());
                  
                  if (!skill) {
                    try {
                      skill = await storage.createSkill({
                        name: skillName,
                        category: options.defaultSkillCategory || 'General',
                        description: null,
                        requiresCertification: false,
                        validityPeriod: null,
                        isActive: true
                      });
                      importResult.summary.skillsCreated++;
                      console.log(`Created skill: ${skill.name}`);
                    } catch (error) {
                      console.error(`Error creating skill ${skillName}:`, error);
                      importResult.summary.errors.push(`Failed to create skill ${skillName}: ${error instanceof Error ? error.message : 'Unknown error'}`);
                      continue;
                    }
                  }
                }

                // Create training record if we have both staff and skill
                if (staff && skill) {
                  try {
                    const trainingRecordData = {
                      staffId: staff.id,
                      skillId: skill.id,
                      competencyLevel: rowData['trainingRecord.competencyLevel'] ? String(rowData['trainingRecord.competencyLevel']) : 'Competent',
                      achievedDate: rowData['trainingRecord.achievedDate'] ? new Date(rowData['trainingRecord.achievedDate']) : new Date(),
                      expiryDate: rowData['trainingRecord.expiryDate'] ? new Date(rowData['trainingRecord.expiryDate']) : null,
                      assessorName: rowData['trainingRecord.assessorName'] ? String(rowData['trainingRecord.assessorName']) : null,
                      trainingProvider: rowData['trainingRecord.trainingProvider'] ? String(rowData['trainingRecord.trainingProvider']) : null,
                      certificateNumber: rowData['trainingRecord.certificateNumber'] ? String(rowData['trainingRecord.certificateNumber']) : null,
                      notes: null,
                      status: rowData['trainingRecord.status'] ? String(rowData['trainingRecord.status']) : 'Active'
                    };

                    await storage.createTrainingRecord(trainingRecordData as any);
                    importResult.summary.trainingRecordsCreated++;
                    console.log(`Created training record for ${staff.name} - ${skill.name}`);
                  } catch (error) {
                    console.error(`Error creating training record for ${staff.name} - ${skill?.name}:`, error);
                    importResult.summary.errors.push(`Failed to create training record for ${staff.name} - ${skill?.name}: ${error instanceof Error ? error.message : 'Unknown error'}`);
                  }
                }
              }

              // Process PPE data if present
              if (rowData['ppeRecord.ppeType'] && staff) {
                try {
                  const ppeData = {
                    staffId: staff.id,
                    ppeType: String(rowData['ppeRecord.ppeType']),
                    brand: rowData['ppeRecord.brand'] ? String(rowData['ppeRecord.brand']) : null,
                    size: rowData['ppeRecord.size'] ? String(rowData['ppeRecord.size']) : null,
                    issueDate: rowData['ppeRecord.issueDate'] ? new Date(rowData['ppeRecord.issueDate']) : new Date(),
                    expiryDate: rowData['ppeRecord.expiryDate'] ? new Date(rowData['ppeRecord.expiryDate']) : null,
                    condition: rowData['ppeRecord.condition'] ? String(rowData['ppeRecord.condition']) : 'Good',
                    location: rowData['ppeRecord.location'] ? String(rowData['ppeRecord.location']) : null,
                    status: 'Issued',
                    notes: null
                  };

                  await storage.createPpeRecord(ppeData);
                  importResult.summary.ppeRecordsCreated++;
                  console.log(`Created PPE record for ${staff.name} - ${ppeData.ppeType}`);
                } catch (error) {
                  console.error(`Error creating PPE record for ${staff.name}:`, error);
                  importResult.summary.errors.push(`Failed to create PPE record for ${staff.name}: ${error instanceof Error ? error.message : 'Unknown error'}`);
                }
              }
            }
          } catch (error) {
            console.error(`Error processing row ${rowIndex + 1} in ${sheet.name}:`, error);
            importResult.summary.errors.push(`Error processing row ${rowIndex + 1} in ${sheet.name}: ${error instanceof Error ? error.message : 'Unknown error'}`);
          }
        }
      }

      console.log('Import completed:', importResult.summary);
      return importResult;
    } catch (error) {
      console.error('Import failed:', error);
      importResult.success = false;
      importResult.summary.errors.push(`Import failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      return importResult;
    }
  }

  // Preview Excel data for import mapping
  app.post('/api/import/preview', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { site, folder, fileName, worksheet } = req.body;

      if (!site || !fileName) {
        return res.status(400).json({
          success: false,
          error: 'site and fileName are required'
        });
      }

      // Read Excel data using existing infrastructure
      const siteUrl = `cranfieldglass.sharepoint.com:/sites/${site}:`;
      const folderPath = folder || '';
      
      const workbookData = await getExcelData(
        accessToken,
        siteUrl,
        folderPath,
        fileName,
        worksheet
      );

      // Analyze data structure for import mapping
      const analysis = {
        workbook: workbookData,
        totalWorksheets: workbookData.sheets.length,
        worksheetAnalysis: workbookData.sheets.map(sheet => ({
          name: sheet.name,
          totalRows: sheet.rows.length,
          totalColumns: sheet.headers.length,
          headers: sheet.headers,
          sampleData: sheet.rows.slice(0, 5), // First 5 rows for preview
          potentialMappings: analyzeColumnMappings(sheet.headers)
        }))
      };

      res.json({ success: true, data: analysis });
    } catch (error) {
      console.error('Error previewing Excel data:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to preview Excel data'
      });
    }
  });

  // Execute import with user-defined mappings
  app.post('/api/import/execute', async (req, res) => {
    try {
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({
          success: false,
          error: 'Authentication required'
        });
      }

      const accessToken = authHeader.substring(7);
      const { site, folder, fileName, worksheet, mappings, options } = req.body;

      if (!site || !fileName || !mappings) {
        return res.status(400).json({
          success: false,
          error: 'site, fileName, and mappings are required'
        });
      }

      // Read Excel data
      const siteUrl = `cranfieldglass.sharepoint.com:/sites/${site}:`;
      const folderPath = folder || '';
      
      const workbookData = await getExcelData(
        accessToken,
        siteUrl,
        folderPath,
        fileName,
        worksheet
      );

      // Execute the import with the provided mappings
      const importResult = await executeSkillsMatrixImport(workbookData, mappings, options);

      res.json({ success: true, data: importResult });
    } catch (error) {
      console.error('Error executing import:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Failed to execute import'
      });
    }
  });

// Generate Meeting Analytics Data
function generateMeetingAnalytics(filteredData: any[]): any {
  const analytics = {
    totalItems: filteredData.length,
    statusBreakdown: {} as Record<string, number>,
    typeBreakdown: {} as Record<string, number>,
    submissionTrends: {} as Record<string, number>,
    assignmentStats: {} as Record<string, number>,
    priorityBreakdown: {} as Record<string, number>,
    outstandingActions: 0,
    completionRate: 0,
    averageResponseTime: 0,
    topContributors: [] as Array<{name: string, count: number}>
  };

  // Process each item for analytics
  filteredData.forEach(item => {
    // Status breakdown
    const status = item.status || 'Unknown';
    analytics.statusBreakdown[status] = (analytics.statusBreakdown[status] || 0) + 1;
    
    // Type breakdown
    const type = item.type || 'Unknown';
    analytics.typeBreakdown[type] = (analytics.typeBreakdown[type] || 0) + 1;
    
    // Submission trends (by month)
    const submissionDate = new Date(item.submittedDate);
    const monthKey = submissionDate.toLocaleDateString('en-GB', { month: 'short', year: 'numeric' });
    analytics.submissionTrends[monthKey] = (analytics.submissionTrends[monthKey] || 0) + 1;
    
    // Assignment stats
    const assignedTo = item.assignedTo || 'Unassigned';
    analytics.assignmentStats[assignedTo] = (analytics.assignmentStats[assignedTo] || 0) + 1;
    
    // Priority breakdown (for Actions)
    if (item.priority) {
      analytics.priorityBreakdown[item.priority] = (analytics.priorityBreakdown[item.priority] || 0) + 1;
    }
    
    // Outstanding actions (items not closed/completed)
    if (status !== 'Closed' && status !== 'Completed') {
      analytics.outstandingActions++;
    }
  });

  // Calculate completion rate
  const closedItems = analytics.statusBreakdown['Closed'] || 0;
  const completedItems = analytics.statusBreakdown['Completed'] || 0;
  analytics.completionRate = filteredData.length > 0 ? 
    Math.round(((closedItems + completedItems) / filteredData.length) * 100) : 0;

  // Top contributors
  const contributorCounts = {} as Record<string, number>;
  filteredData.forEach(item => {
    const contributor = item.submittedBy || 'Unknown';
    contributorCounts[contributor] = (contributorCounts[contributor] || 0) + 1;
  });
  
  analytics.topContributors = Object.entries(contributorCounts)
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 5);

  return analytics;
}

// Generate Simplified Analytics Dashboard HTML
function generateAnalyticsChartsHTML(analytics: any): string {
  const currentDate = new Date().toLocaleDateString('en-NZ', { 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });
  
  const version = "v1.2.0";

  return `
    <div class="analytics-dashboard">
      <div class="analytics-bar">
        <div class="analytics-heading">
          <span class="analytics-h1">Meeting Analytics</span>
          <span class="analytics-sub">${version} &middot; Generated ${currentDate}</span>
        </div>
        <div class="analytics-metrics">
          <div class="metric-chip">
            <span class="metric-chip-num">${analytics.totalItems}</span>
            <span class="metric-chip-label">Total Items</span>
          </div>
          <div class="metric-chip">
            <span class="metric-chip-num">${analytics.outstandingActions}</span>
            <span class="metric-chip-label">Outstanding</span>
          </div>
          <div class="metric-chip">
            <span class="metric-chip-num metric-chip-accent">${analytics.completionRate}%</span>
            <span class="metric-chip-label">Complete</span>
          </div>
        </div>
      </div>
      <div class="analytics-breakdowns">
        <div class="breakdown-col">
          <div class="breakdown-title">Status</div>
          ${Object.entries(analytics.statusBreakdown).map(([status, count]) =>
            `<div class="bd-row"><span class="bd-label">${status}</span><span class="bd-count">${count}</span></div>`
          ).join('')}
        </div>
        <div class="breakdown-col">
          <div class="breakdown-title">Type</div>
          ${Object.entries(analytics.typeBreakdown).map(([type, count]) =>
            `<div class="bd-row"><span class="bd-label">${type}</span><span class="bd-count">${count}</span></div>`
          ).join('')}
        </div>
        <div class="breakdown-col">
          <div class="breakdown-title">Top Contributors</div>
          ${analytics.topContributors.slice(0, 3).map((contributor: any) =>
            `<div class="bd-row"><span class="bd-label">${contributor.name}</span><span class="bd-count">${contributor.count}</span></div>`
          ).join('')}
        </div>
      </div>
    </div>`;
}

// HTML Meeting Minutes Generator Function
function generateMeetingMinutesHTML(filteredData: any[], meetingDate: string, currentDate: string, meetingAttendance?: Record<string, string[]>, selectedMeeting?: string, meetingSignatures?: Record<string, Record<string, { status: string; signatureData: string | null; signedAt: string }>>): string {
  const typeColors = {
    'Business Ideas': '#2563eb',
    'Safety Ideas': '#dc2626', 
    'Near Miss': '#ea580c',
    'Actions': '#7c3aed'
  };

  // Generate analytics for this meeting
  const analytics = generateMeetingAnalytics(filteredData);
  const analyticsChartsHTML = generateAnalyticsChartsHTML(analytics);

  const htmlContent = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cranfield Glass - Meeting Minutes</title>
    <style>
        @page {
            size: A4;
            margin: 1.5cm 1.2cm 1.6cm 1.2cm;
            @bottom-center {
                content: "Cranfield Glass Christchurch  |  Health & Safety Meeting Minutes";
                font-family: Arial, sans-serif;
                font-size: 8pt;
                color: #9ca3af;
            }
            @bottom-right {
                content: "Page " counter(page) " of " counter(pages);
                font-family: Arial, sans-serif;
                font-size: 8pt;
                color: #9ca3af;
            }
        }
        
        body {
            font-family: Arial, sans-serif;
            font-size: 11pt;
            line-height: 1.4;
            color: #333;
            background-color: #ffffff;
            margin: 0;
            /* No body padding: the @page rule already supplies the A4 page
               margins. Extra padding here would stack on top and shrink the
               usable content width. */
            padding: 0;
        }
        
        .container {
            max-width: none;
            width: 100%;
            margin: 0;
            background: #ffffff;
            border-radius: 0;
            box-shadow: none;
            /* Let content fill the full A4 content area defined by @page. */
            padding: 0;
            border: none;
        }
        
        @media print {
            body {
                margin: 0 !important;
                padding: 0 !important;
            }
            
            .container {
                padding: 0 !important;
                margin: 0 !important;
                width: 100% !important;
                max-width: none !important;
            }
            
            .print-button {
                display: none !important;
            }
        }
        
        .print-button {
            position: fixed;
            top: 20px;
            right: 20px;
            background: #3b82f6;
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            z-index: 1000;
        }
        
        .print-button:hover {
            background: #2563eb;
        }
        
        .header {
            text-align: center;
            margin-bottom: 30px;
            border-bottom: 2px solid #3b82f6;
            padding-bottom: 20px;
        }
        
        .company-name {
            font-size: 24pt;
            font-weight: bold;
            background: linear-gradient(135deg, #14b8a6, #06b6d4, #3b82f6);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            margin-bottom: 5px;
        }
        
        .document-title {
            font-size: 18pt;
            font-weight: bold;
            color: #333;
            margin-bottom: 10px;
        }
        
        .meeting-info {
            display: flex;
            justify-content: space-between;
            margin-bottom: 30px;
            padding: 15px;
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
        }
        
        .info-block {
            flex: 1;
        }
        
        .info-label {
            font-weight: bold;
            margin-bottom: 5px;
        }
        
        .agenda-section {
            margin-bottom: 30px;
        }
        
        .section-header {
            font-size: 14pt;
            font-weight: bold;
            color: #3b82f6;
            margin-bottom: 15px;
            border-bottom: 1px solid #3b82f6;
            padding-bottom: 5px;
        }
        
        .items-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
            border: 1px solid #dee2e6;
        }
        
        .items-table th {
            background-color: #f8f9fa;
            padding: 12px;
            text-align: left;
            font-weight: bold;
            border: 1px solid #dee2e6;
            font-size: 10pt;
        }
        
        .items-table td {
            padding: 12px;
            border: 1px solid #dee2e6;
            vertical-align: top;
            font-size: 10pt;
        }
        
        .type-badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 4px;
            color: white;
            font-size: 9pt;
            font-weight: bold;
            margin-bottom: 5px;
        }
        
        .agenda-item {
            font-weight: bold;
            margin-bottom: 8px;
        }
        
        .discussion-notes {
            margin-bottom: 8px;
            white-space: pre-wrap;
        }
        
        .meeting-notes {
            background-color: #fff3cd;
            padding: 8px;
            border-left: 4px solid #ffc107;
            margin-bottom: 8px;
            font-style: italic;
        }
        
        .hso-action {
            background-color: #d1ecf1;
            padding: 8px;
            border-left: 4px solid #0dcaf0;
            margin-bottom: 8px;
        }
        
        .follow-up {
            font-size: 9pt;
            color: #666;
        }
        
        .signatures {
            margin-top: 40px;
            page-break-inside: avoid;
        }
        
        .signature-row {
            display: flex;
            margin-bottom: 30px;
        }
        
        .signature-block {
            flex: 1;
            margin-right: 20px;
        }
        
        .signature-line {
            border-bottom: 1px solid #333;
            height: 40px;
            margin-bottom: 5px;
        }
        
        .signature-label {
            font-size: 10pt;
            color: #666;
        }
        
        /* Analytics Dashboard Styles */
        .analytics-dashboard {
            margin-bottom: 22px;
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            overflow: hidden;
            page-break-inside: avoid;
            break-inside: avoid;
        }

        .analytics-bar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            padding: 10px 16px;
            background: #f8fafc;
            border-bottom: 1px solid #e5e7eb;
        }

        .analytics-heading {
            display: flex;
            flex-direction: column;
        }

        .analytics-h1 {
            font-size: 13pt;
            font-weight: bold;
            color: #1f2937;
            line-height: 1.2;
        }

        .analytics-sub {
            font-size: 7.5pt;
            color: #9ca3af;
            margin-top: 2px;
        }

        .analytics-metrics {
            display: flex;
            gap: 10px;
        }

        .metric-chip {
            display: flex;
            flex-direction: column;
            align-items: center;
            min-width: 64px;
            padding: 4px 12px;
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 6px;
        }

        .metric-chip-num {
            font-size: 16pt;
            font-weight: bold;
            color: #1f2937;
            line-height: 1.1;
        }

        .metric-chip-num.metric-chip-accent {
            color: #2563eb;
        }

        .metric-chip-label {
            font-size: 7pt;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            margin-top: 1px;
        }

        .analytics-breakdowns {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 0;
        }

        .breakdown-col {
            padding: 10px 16px;
            border-right: 1px solid #f1f5f9;
        }

        .breakdown-col:last-child {
            border-right: none;
        }

        .breakdown-title {
            font-size: 8pt;
            font-weight: bold;
            color: #374151;
            text-transform: uppercase;
            letter-spacing: 0.4px;
            margin-bottom: 6px;
            padding-bottom: 4px;
            border-bottom: 1px solid #e5e7eb;
        }

        .bd-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 8px;
            padding: 2px 0;
            font-size: 9pt;
        }

        .bd-label {
            color: #4b5563;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .bd-count {
            font-weight: bold;
            color: #1f2937;
            font-variant-numeric: tabular-nums;
        }

        @media print {
            body { font-size: 10pt; }
            .items-table th, .items-table td { font-size: 9pt; }
            .analytics-dashboard { break-inside: avoid; }
            .analytics-breakdowns { grid-template-columns: repeat(3, 1fr); }
        }
    </style>
</head>
<body>
    <div class="container">
    <div class="header">
        <div class="company-name">Cranfield Glass Christchurch</div>
        <div class="document-title">Health & Safety Meeting Minutes</div>
    </div>

    <div class="meeting-info">
        <div class="info-block">
            <div class="info-label">Meeting Topic:</div>
            <div>Tuesday Staff Meeting</div>
        </div>
        <div class="info-block">
            <div class="info-label">Date:</div>
            <div>${meetingDate}</div>
        </div>
        <div class="info-block">
            <div class="info-label">Generated:</div>
            <div>${currentDate}</div>
        </div>
    </div>

    ${analyticsChartsHTML}

    <div class="agenda-section">
        <div class="section-header">I. Agenda</div>
        <ul>
            <li>Matters arising from previous meeting minutes</li>
            <li>Health and Safety – Refer below</li>
            <li>BAU – Business as Usual</li>
            <li>Other items as required</li>
        </ul>
    </div>

    <div class="section-header">II. Meeting Minutes</div>
    
    <table class="items-table">
        <thead>
            <tr>
                <th style="width: 25%;">Agenda Item</th>
                <th style="width: 35%;">Discussion Notes</th>
                <th style="width: 40%;">Action Required</th>
            </tr>
        </thead>
        <tbody>
            ${filteredData.map(item => {
              const typeColor = typeColors[item.type as keyof typeof typeColors] || '#6b7280';
              
              let discussionNotes = item.description || '';
              if (item.secondaryDescription) {
                discussionNotes += discussionNotes ? `\n\nHow it happened: ${item.secondaryDescription}` : item.secondaryDescription;
              }
              
              return `
                <tr>
                    <td>
                        <div class="type-badge" style="background-color: ${typeColor};">${item.type}</div>
                        <div class="agenda-item">${item.title || `${item.type} Item`}</div>
                        <div class="follow-up">
                            <strong>Submitted by:</strong> ${item.submittedBy || 'Unknown'}<br>
                            <strong>Status:</strong> ${item.status || 'Submitted'}
                        </div>
                    </td>
                    <td>
                        <div class="discussion-notes">${discussionNotes}</div>
                        ${item.meetingNotes ? `<div class="meeting-notes"><strong>Meeting Discussion:</strong> ${item.meetingNotes}</div>` : ''}
                    </td>
                    <td>
                        ${generateActionRequiredSection(item)}
                    </td>
                </tr>
              `;
            }).join('')}
        </tbody>
    </table>

    ${(() => {
      const readyToCloseItems = filteredData.filter((item: any) => item.actionStatus === 'Ready to Close');
      if (readyToCloseItems.length === 0) return '';
      return `
    <div class="section-header" style="margin-top:24px;">III. Completed Actions – Group Review</div>
    <div style="background:#f0fdf4;border:1px solid #86efac;border-radius:6px;padding:12px;margin-bottom:16px;">
      <div style="font-size:10pt;color:#166534;margin-bottom:10px;">
        <strong>&#10003; ${readyToCloseItems.length} completed action${readyToCloseItems.length !== 1 ? 's' : ''} require group discussion and sign-off to formally close.</strong>
        This section demonstrates our circular compliance system — actions are completed, reviewed by the group, and officially closed at the meeting.
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:10pt;">
        <thead>
          <tr style="background:#dcfce7;">
            <th style="padding:6px 8px;border:1px solid #86efac;text-align:left;width:28%;">Item</th>
            <th style="padding:6px 8px;border:1px solid #86efac;text-align:left;width:20%;">Actioned By</th>
            <th style="padding:6px 8px;border:1px solid #86efac;text-align:left;width:32%;">What Was Done</th>
            <th style="padding:6px 8px;border:1px solid #86efac;text-align:left;width:20%;">Group Outcome</th>
          </tr>
        </thead>
        <tbody>
          ${readyToCloseItems.map((item: any) => `
          <tr>
            <td style="padding:6px 8px;border:1px solid #bbf7d0;vertical-align:top;">
              <div style="background:${item.type === 'Safety Ideas' ? '#dc2626' : item.type === 'Near Miss' ? '#ea580c' : '#2563eb'};color:#fff;font-size:8pt;padding:2px 6px;border-radius:3px;display:inline-block;margin-bottom:4px;">${item.type}</div>
              <div style="font-weight:600;font-size:10pt;">${item.title || `${item.type} Item`}</div>
              <div style="font-size:9pt;color:#555;margin-top:2px;">Submitted by: ${item.submittedBy || 'Unknown'}</div>
            </td>
            <td style="padding:6px 8px;border:1px solid #bbf7d0;vertical-align:top;">
              ${item.actionAssignedTo ? `<div style="font-weight:600;">${item.actionAssignedTo}</div>` : '<div style="color:#888;">—</div>'}
              ${item.actionDueDate ? `<div style="font-size:9pt;color:#555;">Due: ${new Date(item.actionDueDate).toLocaleDateString('en-NZ')}</div>` : ''}
            </td>
            <td style="padding:6px 8px;border:1px solid #bbf7d0;vertical-align:top;">
              ${item.actionNotes ? `<div>${item.actionNotes}</div>` : ''}
              ${item.meetingNotes ? `<div style="margin-top:4px;color:#555;font-size:9pt;"><em>Discussion: ${item.meetingNotes}</em></div>` : ''}
              ${!item.actionNotes && !item.meetingNotes ? '<div style="color:#888;">No notes recorded</div>' : ''}
            </td>
            <td style="padding:6px 8px;border:1px solid #bbf7d0;vertical-align:top;">
              <div style="background:#dcfce7;border:1px dashed #86efac;border-radius:4px;padding:6px;font-size:9pt;color:#166534;min-height:36px;">
                Group decision:<br><br>____________________
              </div>
            </td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
    })()}

    <div class="signatures">
        <div class="section-header">IV. Meeting Attendance &amp; Sign-Off</div>
        
        ${generateAttendanceSection(meetingAttendance, selectedMeeting, meetingDate, meetingSignatures)}
    </div>
    </div>

    <script>
      // Add the floating Print button once (used by Paged.js and the fallback).
      function addPrintButton() {
        if (document.querySelector('.print-button')) return;
        var btn = document.createElement('button');
        btn.className = 'print-button';
        btn.textContent = '\u{1F5A8}\uFE0F Print Meeting Minutes';
        btn.addEventListener('click', function () { window.print(); });
        document.body.appendChild(btn);
      }
      // Configure Paged.js: re-add the Print button once pagination completes.
      window.PagedConfig = { auto: true, after: addPrintButton };
      // Resilience fallback: if Paged.js never finishes (rare), still give the
      // user a Print button so the document remains usable.
      setTimeout(addPrintButton, 8000);
    </script>
    <script>${PAGEDJS_SCRIPT}</script>
</body>
</html>
  `;

  return htmlContent;
}

// Generate action required section with improved logic instead of generic text
function generateActionRequiredSection(item: any): string {
  let actionHtml = '';
  
  // Check for embedded action data first
  if (item.actionAssignedTo || item.actionStatus || item.actionNotes) {
    if (item.actionAssignedTo) {
      actionHtml += `<div class="follow-up"><strong>Assigned to:</strong> ${item.actionAssignedTo}</div>`;
    }
    
    if (item.actionStatus) {
      actionHtml += `<div class="follow-up"><strong>Status:</strong> ${item.actionStatus}</div>`;
    }
    
    if (item.actionDueDate) {
      actionHtml += `<div class="follow-up"><strong>Due:</strong> ${new Date(item.actionDueDate).toLocaleDateString('en-NZ')}</div>`;
    }
    
    if (item.actionNotes) {
      actionHtml += `<div class="follow-up"><strong>Notes:</strong> ${item.actionNotes}</div>`;
    }
    
    return actionHtml || '<div class="follow-up">Action details to be updated.</div>';
  }
  
  // If there are actual meeting notes, use them as the action/decision
  if (item.meetingNotes && item.meetingNotes.trim() && item.meetingNotes.trim() !== '') {
    actionHtml += `<div class="follow-up"><strong>Decision:</strong> ${item.meetingNotes}</div>`;
    
    // Add assignment if available
    if (item.assignedTo) {
      actionHtml += `<div class="follow-up"><strong>Assigned to:</strong> ${item.assignedTo}</div>`;
    }
    
    return actionHtml;
  }

  // If assigned to someone specific, provide personalized action
  if (item.assignedTo) {
    actionHtml += `<div class="follow-up"><strong>Assigned to:</strong> ${item.assignedTo}</div>`;
    const actionText = item.status === 'Actioned' ? 'implement and report back' : 'review and assess feasibility';
    actionHtml += `<div class="follow-up"><strong>Action:</strong> ${actionText} by next meeting</div>`;
    return actionHtml;
  }
  
  // Generate appropriate action based on type and status (no generic "Monitor progress")
  switch (item.type) {
    case 'Near Miss':
      if (item.status === 'Closed') {
        actionHtml = '<div class="follow-up"><strong>Status:</strong> Incident reviewed, preventive measures implemented, risk register updated.</div>';
      } else if (item.status === 'Actioned') {
        actionHtml = '<div class="follow-up"><strong>Action:</strong> HSC implementing preventive measures and updating documentation.</div>';
      } else {
        actionHtml = '<div class="follow-up"><strong>Required:</strong> HSC to review incident details and implement preventive measures.</div>';
      }
      break;
    case 'Safety Ideas':
      if (item.status === 'Closed') {
        actionHtml = '<div class="follow-up"><strong>Status:</strong> Safety improvement evaluated and implementation decision made.</div>';
      } else if (item.status === 'Actioned') {
        actionHtml = '<div class="follow-up"><strong>Action:</strong> HSC assessing implementation feasibility and resource requirements.</div>';
      } else {
        actionHtml = '<div class="follow-up"><strong>Required:</strong> Pending review by Health & Safety Coordinator for feasibility assessment.</div>';
      }
      break;
    case 'Business Ideas':
      if (item.status === 'Closed') {
        actionHtml = '<div class="follow-up"><strong>Status:</strong> Business case evaluated and implementation decision finalized.</div>';
      } else if (item.status === 'Actioned') {
        actionHtml = '<div class="follow-up"><strong>Action:</strong> Management evaluating business case and implementation timeline.</div>';
      } else {
        actionHtml = '<div class="follow-up"><strong>Required:</strong> Pending management review for business case evaluation.</div>';
      }
      break;
    case 'Actions':
      if (item.status === 'Closed') {
        actionHtml = '<div class="follow-up"><strong>Status:</strong> Action completed and verified by responsible person.</div>';
      } else if (item.status === 'In Progress') {
        actionHtml = '<div class="follow-up"><strong>Action:</strong> Action in progress, monitoring completion status.</div>';
      } else {
        actionHtml = '<div class="follow-up"><strong>Required:</strong> Action item pending assignment and commencement.</div>';
      }
      break;
    default:
      actionHtml = '<div class="follow-up"><strong>Required:</strong> Item requires review to determine appropriate action plan.</div>';
  }
  
  if (item.actionStartDate) {
    actionHtml += `<div class="follow-up"><strong>Start Date:</strong> ${item.actionStartDate}</div>`;
  }
  
  return actionHtml;
}

// Generate attendance section with UI synchronization
function generateAttendanceSection(meetingAttendance?: Record<string, string[]>, selectedMeeting?: string, meetingDate?: string, meetingSignatures?: Record<string, Record<string, { status: string; signatureData: string | null; signedAt: string }>>): string {
  const allAttendees = [
    { name: 'Hoani Hunt', role: 'Company Director' },
    { name: 'Simon Hubbard', role: 'Health & Safety Coordinator' },
    { name: 'James Waites', role: 'Glazing Supervisor' },
    { name: 'Emma White', role: 'Administrator' },
    { name: 'Kevin Young', role: 'Glazier' },
    { name: 'Ryan Newman', role: 'Glazier' },
    { name: 'Dan Conlan', role: 'Glazier' },
    { name: "Struan O'Donnell", role: 'Glazier' },
    { name: 'Sam Chang', role: 'Glazier' }
  ];

  const isPresent = (name: string): boolean => {
    if (!meetingAttendance || !selectedMeeting || selectedMeeting === 'all') return true;
    const list = meetingAttendance[selectedMeeting];
    if (!list) return true;
    return list.includes(name);
  };

  const signaturesForMeeting = meetingSignatures && selectedMeeting && selectedMeeting !== 'all'
    ? (meetingSignatures[selectedMeeting] ?? {})
    : {};

  const hasSigs = Object.keys(signaturesForMeeting).length > 0;

  const formatDate = (iso: string) => {
    try { return new Date(iso).toLocaleDateString('en-NZ', { day: '2-digit', month: 'short', year: 'numeric' }); } catch { return iso; }
  };

  const signedAttendees = allAttendees.filter(a => {
    const sig = signaturesForMeeting[a.name];
    return sig && (sig.status === 'signed' || sig.status === 'remote');
  });

  const absentAttendees = allAttendees.filter(a => {
    const sig = signaturesForMeeting[a.name];
    return !isPresent(a.name) || (sig && sig.status === 'absent');
  });

  const pendingAttendees = allAttendees.filter(a => {
    const sig = signaturesForMeeting[a.name];
    return isPresent(a.name) && !sig;
  });

  if (hasSigs) {
    const sigBlocks = signedAttendees.map(a => {
      const sig = signaturesForMeeting[a.name];
      const isRemote = sig.status === 'remote';
      return `
        <div style="break-inside: avoid; border: 1px solid #e5e7eb; border-radius: 8px; padding: 16px; background: #fafafa;">
          <div style="font-weight: bold; font-size: 11pt; color: #1f2937;">${a.name}</div>
          <div style="font-size: 9pt; color: #6b7280; margin-bottom: 10px;">${a.role}${isRemote ? ' &nbsp;·&nbsp; <span style="color:#7c3aed;">Attended Remotely</span>' : ''}</div>
          ${isRemote
            ? `<div style="font-family: 'Georgia', serif; font-style: italic; font-size: 22pt; color: #1f2937; padding: 8px 0; border-bottom: 2px solid #374151; min-height: 48px;">${a.name}</div>`
            : sig.signatureData
              ? `<img src="${sig.signatureData}" style="max-width: 200px; max-height: 60px; display: block; border-bottom: 2px solid #374151;" alt="Signature">`
              : `<div style="border-bottom: 2px solid #374151; height: 48px;"></div>`
          }
          <div style="font-size: 8pt; color: #9ca3af; margin-top: 6px;">Signed: ${formatDate(sig.signedAt)}</div>
        </div>`;
    }).join('');

    const pendingList = pendingAttendees.map(a =>
      `<span style="display: inline-block; padding: 4px 10px; background: #fef3c7; border: 1px solid #fcd34d; border-radius: 999px; font-size: 9pt; color: #92400e; margin: 3px;">${a.name}</span>`
    ).join('');

    const absentList = absentAttendees.map(a =>
      `<span style="display: inline-block; padding: 4px 10px; background: #f3f4f6; border: 1px solid #d1d5db; border-radius: 999px; font-size: 9pt; color: #6b7280; margin: 3px;">${a.name}</span>`
    ).join('');

    return `
      <div style="margin-bottom: 24px;">
        <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 16px; margin-bottom: 20px;">
          ${sigBlocks || '<p style="color:#9ca3af; font-size:10pt;">No signatures collected yet.</p>'}
        </div>
        ${pendingList ? `<div style="margin-bottom: 12px;"><span style="font-size:9pt; font-weight:600; color:#92400e;">Pending signature: </span>${pendingList}</div>` : ''}
        ${absentList ? `<div><span style="font-size:9pt; font-weight:600; color:#6b7280;">Not present: </span>${absentList}</div>` : ''}
      </div>`;
  }

  // Fallback: no signatures yet — show simple checkbox attendance table
  const rows = allAttendees.map(a => `
    <div style="display: flex; align-items: center; gap: 8px; padding: 8px; border: 1px solid #e5e7eb; border-radius: 6px;">
        <span style="font-size: 16px;">${isPresent(a.name) ? '✅' : '⬜'}</span>
        <span style="font-size: 10pt; color: #374151;">${a.name} – ${a.role}</span>
    </div>`).join('');

  return `
    <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; margin-bottom: 20px;">
      ${rows}
    </div>
    <div style="text-align: center; margin-top: 20px; padding: 16px; background: #f8fafc; border-radius: 8px; border: 1px solid #e5e7eb;">
      <p style="margin: 0; font-size: 9pt; color: #9ca3af;">Signatures not yet collected — use the Collect Signatures feature in the app to capture sign-off.</p>
    </div>`;
}

  // Skills Matrix CSV export - Updated for 5-level competency system
  app.post('/api/export-skills-matrix-csv', async (req, res) => {
    try {
      const { staffFilter, skillCategory } = req.body;
      
      // Get real data that matches the UI
      const allStaff = await storage.getAllStaff();
      const allClassifications = await storage.getAllTrainingClassifications();
      const allModules = await storage.getAllTrainingModules();
      const allMatrixData = await storage.getTrainingModuleMatrix();
      
      // Filter staff
      let filteredStaff = allStaff;
      if (staffFilter === 'field') {
        filteredStaff = allStaff.filter(staff => staff.isFieldStaff);
      } else if (staffFilter === 'administration') {
        filteredStaff = allStaff.filter(staff => staff.isAdministrationStaff);
      } else if (staffFilter === 'active') {
        filteredStaff = allStaff.filter(staff => staff.isActive);
      } else if (staffFilter === 'inactive') {
        filteredStaff = allStaff.filter(staff => !staff.isActive);
      }
      
      // Filter classifications by audience if needed
      let filteredClassifications = allClassifications.filter(c => c.isActive);
      if (skillCategory === 'field') {
        filteredClassifications = filteredClassifications.filter(c => c.audience === 'field');
      } else if (skillCategory === 'administration') {
        filteredClassifications = filteredClassifications.filter(c => c.audience === 'administration');
      }
      
      // Build modules list from classifications for CSV headers - include safety-critical info
      const allModulesForExport: any[] = [];
      filteredClassifications.forEach(classification => {
        const classificationModules = allModules.filter(m => m.classificationId === classification.id);
        classificationModules.forEach(module => {
          allModulesForExport.push({
            id: module.id,
            name: module.name,
            code: module.code,
            classificationName: classification.name,
            isSafetyCritical: module.isSafetyCritical || false,
            requiresCertification: module.requiresCertification || false
          });
        });
      });
      
      // Sort modules: Site Safety Induction first, then safety-critical, then cert-required, then by classification and name
      allModulesForExport.sort((a, b) => {
        // Site Safety Induction always first
        if (a.name === 'Site Safety Induction' && b.name !== 'Site Safety Induction') return -1;
        if (a.name !== 'Site Safety Induction' && b.name === 'Site Safety Induction') return 1;
        
        // Safety-critical first
        if (a.isSafetyCritical && !b.isSafetyCritical) return -1;
        if (!a.isSafetyCritical && b.isSafetyCritical) return 1;
        
        // Then cert-required
        if (a.requiresCertification && !b.requiresCertification) return -1;
        if (!a.requiresCertification && b.requiresCertification) return 1;
        
        // Then by classification
        if (a.classificationName !== b.classificationName) {
          return a.classificationName.localeCompare(b.classificationName);
        }
        
        // Finally by name
        return a.name.localeCompare(b.name);
      });
      
      // Generate CSV content with clean, readable headers
      const headers = ['Staff Name', 'Email', 'Job Title', 'Status', ...allModulesForExport.map(mod => {
        const safetyCriticalIndicator = mod.isSafetyCritical ? '⚠ ' : '';
        const certificationIndicator = mod.requiresCertification ? '📋 ' : '';
        return `${safetyCriticalIndicator}${certificationIndicator}${mod.code} - ${mod.name}`;
      })];
      
      const csvRows = filteredStaff.map(staff => {
        const row = [
          staff.name,
          staff.email || '',
          staff.jobTitle || '',
          staff.isActive ? 'Active' : 'Inactive'
        ];
        
        // Add module competencies with enhanced information
        allModulesForExport.forEach(module => {
          const matrixRecord = allMatrixData.find(record => 
            record.staffId === staff.id && record.moduleId === module.id
          );
          
          if (matrixRecord) {
            const competencyLevel = matrixRecord.competencyLevel || COMPETENCY_LEVELS.NOT_TRAINED;
            const isAbleToUseStatus = isAbleToUse(competencyLevel) ? ' - ABLE TO USE' : '';
            const expiryInfo = matrixRecord.expiryDate ? 
              ` (Expires: ${new Date(matrixRecord.expiryDate).toLocaleDateString('en-GB')})` : '';
            const isExpired = matrixRecord.expiryDate && new Date(matrixRecord.expiryDate) < new Date() ? ' [EXPIRED]' : '';
            row.push(`${competencyLevel}${isAbleToUseStatus}${expiryInfo}${isExpired}`);
          } else {
            row.push(COMPETENCY_LEVELS.NOT_TRAINED);
          }
        });
        
        return row;
      });
      
      // Create CSV content
      const csvContent = [
        headers.map(header => `"${header.replace(/"/g, '""')}"`).join(','),
        ...csvRows.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(','))
      ].join('\n');
      
      // Set headers
      const filename = `Cranfield-Glass-Skills-Matrix-${new Date().toLocaleDateString('en-GB').replace(/\//g, '-')}.csv`;
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
      res.send(csvContent);
      
    } catch (error) {
      console.error('Error generating Skills Matrix CSV:', error);
      res.status(500).json({ error: 'Failed to generate Skills Matrix CSV' });
    }
  });

  // Skills Matrix HTML export
  app.post('/api/export-skills-matrix-html', async (req, res) => {
    try {
      const { staffFilter, skillCategory } = req.body;
      
      // Get real data that matches the UI
      const allStaff = await storage.getAllStaff();
      const allClassifications = await storage.getAllTrainingClassifications();
      const allModules = await storage.getAllTrainingModules();
      const allMatrixData = await storage.getTrainingModuleMatrix();
      
      // Filter staff
      let filteredStaff = allStaff;
      if (staffFilter === 'field') {
        filteredStaff = allStaff.filter(staff => staff.isFieldStaff);
      } else if (staffFilter === 'administration') {
        filteredStaff = allStaff.filter(staff => staff.isAdministrationStaff);
      } else if (staffFilter === 'active') {
        filteredStaff = allStaff.filter(staff => staff.isActive);
      } else if (staffFilter === 'inactive') {
        filteredStaff = allStaff.filter(staff => !staff.isActive);
      }
      
      // Filter classifications by audience if needed
      let filteredClassifications = allClassifications.filter(c => c.isActive);
      if (skillCategory === 'field') {
        filteredClassifications = filteredClassifications.filter(c => c.audience === 'field');
      } else if (skillCategory === 'administration') {
        filteredClassifications = filteredClassifications.filter(c => c.audience === 'administration');
      }
      
      // Build modules list from classifications - include safety-critical info
      const allModulesForExport: any[] = [];
      filteredClassifications.forEach(classification => {
        const classificationModules = allModules.filter(m => m.classificationId === classification.id);
        classificationModules.forEach(module => {
          allModulesForExport.push({
            id: module.id,
            name: module.name,
            code: module.code,
            classificationName: classification.name,
            isSafetyCritical: module.isSafetyCritical || false,
            requiresCertification: module.requiresCertification || false
          });
        });
      });
      
      // Sort modules: Site Safety Induction first, then safety-critical, then cert-required, then by classification and name
      allModulesForExport.sort((a, b) => {
        // Site Safety Induction always first
        if (a.name === 'Site Safety Induction' && b.name !== 'Site Safety Induction') return -1;
        if (a.name !== 'Site Safety Induction' && b.name === 'Site Safety Induction') return 1;
        
        // Safety-critical first
        if (a.isSafetyCritical && !b.isSafetyCritical) return -1;
        if (!a.isSafetyCritical && b.isSafetyCritical) return 1;
        
        // Then cert-required
        if (a.requiresCertification && !b.requiresCertification) return -1;
        if (!a.requiresCertification && b.requiresCertification) return 1;
        
        // Then by classification
        if (a.classificationName !== b.classificationName) {
          return a.classificationName.localeCompare(b.classificationName);
        }
        
        // Finally by name
        return a.name.localeCompare(b.name);
      });
      
      // Generate HTML content
      const currentDate = new Date().toLocaleDateString('en-GB');
      const htmlContent = `<!DOCTYPE html>
<html>
<head>
    <title>Cranfield Glass - Skills Matrix</title>
    <meta charset="utf-8">
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif; margin: 20px; font-size: 13px; }
        .header { text-align: center; margin-bottom: 30px; }
        .company-logo { color: #1e40af; font-size: 24px; font-weight: bold; }
        .document-title { color: #374151; font-size: 20px; margin: 10px 0; }
        .date-info { color: #6b7280; font-size: 13px; }
        .skills-matrix { border-collapse: collapse; width: 100%; margin: 20px 0; font-size: 12px; }
        .skills-matrix th, .skills-matrix td { border: 1px solid #d1d5db; padding: 6px; text-align: center; vertical-align: middle; }
        .skills-matrix th { background-color: #f3f4f6; font-weight: 600; font-size: 10px; line-height: 1.3; white-space: normal; max-width: 120px; }
        .skills-matrix td:first-child, .skills-matrix th:first-child { text-align: left; white-space: normal; }
        
        .status-badge { display: inline-block; padding: 3px 8px; border-radius: 12px; font-size: 10px; font-weight: 600; white-space: nowrap; }
        .status-able { background-color: #059669; color: white; }
        .status-training { background-color: #d97706; color: white; }
        .status-not-trained { background-color: #6b7280; color: white; }
        .status-expired { background-color: #dc2626; color: white; }
        
        .expiry-text { display: block; font-size: 10px; color: #6b7280; margin-top: 3px; }
        .expiry-warning { color: #dc2626; font-weight: 600; }
        
        .module-badge { display: inline-block; padding: 2px 6px; border-radius: 3px; font-size: 9px; font-weight: 600; margin: 2px 0; }
        .badge-critical { background-color: #dc2626; color: white; }
        .badge-cert { background-color: #ea580c; color: white; }
        
        .footer { margin-top: 40px; font-size: 11px; color: #6b7280; text-align: center; }
    </style>
</head>
<body>
    <div class="header">
        <div class="company-logo">CRANFIELD GLASS CHRISTCHURCH</div>
        <div class="document-title">Staff Skills Matrix</div>
        <div class="date-info">Generated: ${currentDate} | Staff Filter: ${staffFilter} | Skills: ${skillCategory === 'all' ? 'All Categories' : skillCategory}</div>
    </div>
    
    <table class="skills-matrix">
        <thead>
            <tr>
                <th style="min-width: 150px;">Staff</th>
                <th>Status</th>
                ${allModulesForExport.map(module => {
                  const safetyCriticalBadge = module.isSafetyCritical ? '<br><span class="module-badge badge-critical">⚠ SAFETY CRITICAL</span>' : '';
                  const certificationBadge = module.requiresCertification ? '<br><span class="module-badge badge-cert">📋 CERT REQUIRED</span>' : '';
                  return `<th><div style="line-height: 1.3;">${module.name}<br><small style="color: #6b7280;">${module.code}</small>${safetyCriticalBadge}${certificationBadge}</div></th>`;
                }).join('')}
            </tr>
        </thead>
        <tbody>
            ${filteredStaff.map(staff => {
              return `<tr>
                <td><strong>${staff.name}</strong><br><small style="color: #6b7280;">${staff.jobTitle || ''}</small></td>
                <td>${staff.isActive ? 'Active' : 'Inactive'}</td>
                ${allModulesForExport.map(module => {
                  const matrixRecord = allMatrixData.find(record => 
                    record.staffId === staff.id && record.moduleId === module.id
                  );
                  
                  if (matrixRecord) {
                    const competencyLevel = matrixRecord.competencyLevel || COMPETENCY_LEVELS.NOT_TRAINED;
                    const isExpired = matrixRecord.expiryDate && new Date(matrixRecord.expiryDate) < new Date();
                    const canUse = isAbleToUse(competencyLevel) || competencyLevel === 'Competent';
                    
                    // Determine badge style
                    let badgeClass = 'status-not-trained';
                    let badgeText = '✗';
                    
                    if (isExpired) {
                      badgeClass = 'status-expired';
                      badgeText = '⚠ EXPIRED';
                    } else if (canUse) {
                      badgeClass = 'status-able';
                      badgeText = '✓ Able to Use';
                    } else if (competencyLevel === 'In Training' || competencyLevel === COMPETENCY_LEVELS.IN_TRAINING_SUPERVISED) {
                      badgeClass = 'status-training';
                      badgeText = '⟳ Training';
                    }
                    
                    const expiryInfo = matrixRecord.expiryDate ? 
                      `<span class="expiry-text ${isExpired ? 'expiry-warning' : ''}">Exp: ${new Date(matrixRecord.expiryDate).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit' })}</span>` : '';
                    
                    return `<td><span class="status-badge ${badgeClass}">${badgeText}</span>${expiryInfo}</td>`;
                  } else {
                    return `<td><span class="status-badge status-not-trained">✗</span></td>`;
                  }
                }).join('')}
              </tr>`;
            }).join('')}
        </tbody>
    </table>
    
    <div class="footer">
        <p>Cranfield Glass Christchurch - Staff Training & Competency Management System</p>
        <p>This document contains confidential training records and competency assessments.</p>
    </div>
</body>
</html>`;
      
      // Set headers
      const filename = `Cranfield-Glass-Skills-Matrix-${currentDate.replace(/\//g, '-')}.html`;
      res.json({
        success: true,
        filename: filename,
        htmlContent: htmlContent,
        message: 'Skills Matrix HTML generated successfully'
      });
      
    } catch (error) {
      console.error('Error generating Skills Matrix HTML:', error);
      res.status(500).json({ error: 'Failed to generate Skills Matrix HTML' });
    }
  });

  // AI Classification endpoint for Teams Personal Tab
  app.post('/api/ai-classify', async (req, res) => {
    try {
      // Validate bearer token by calling Graph /me — prevents cost-abuse from fake tokens
      const authHeader = req.headers.authorization;
      if (!authHeader || !authHeader.startsWith('Bearer ')) {
        return res.status(401).json({ success: false, error: 'Authentication required' });
      }
      const accessToken = authHeader.substring(7);
      const meResp = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      if (!meResp.ok) {
        return res.status(401).json({ success: false, error: 'Invalid or expired token' });
      }

      const { text } = req.body;

      if (!text || typeof text !== 'string' || text.trim().length === 0) {
        return res.status(400).json({
          success: false,
          error: 'Missing required field: text'
        });
      }

      const result = await OpenAIService.classifySubmission(text);

      res.json({ success: true, ...result });

    } catch (error) {
      console.error('Error in AI classify endpoint:', error);
      res.status(500).json({
        success: false,
        error: error instanceof Error ? error.message : 'Classification failed'
      });
    }
  });

  // ─── Order Items endpoints (Teams Whiteboard ordering pad) ──────────────────

  // Helper: resolve caller's display name from Bearer token via Graph /me.
  // Returns displayName on success, null if token is missing or invalid.
  async function resolveCallerFromToken(authHeader: string | undefined): Promise<{ displayName: string; email: string } | null> {
    if (!authHeader || !authHeader.startsWith('Bearer ')) return null;
    const token = authHeader.substring(7);
    try {
      const meRes = await fetch('https://graph.microsoft.com/v1.0/me?$select=userPrincipalName,displayName', {
        headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
      });
      if (!meRes.ok) return null;
      const me = await meRes.json();
      const email = (me.userPrincipalName || '').toLowerCase().trim();
      const displayName = me.displayName || email;
      if (!email) return null;
      return { displayName, email };
    } catch {
      return null;
    }
  }

  // Helper: verify caller is admin/supervisor via staff table (email from Graph token).
  async function resolveOrderAdminFromToken(authHeader: string | undefined): Promise<string | null> {
    const caller = await resolveCallerFromToken(authHeader);
    if (!caller) return null;
    const staffMember = await storage.getStaffByEmail(caller.email);
    if (!staffMember) return null;
    if (staffMember.role === 'admin' || staffMember.role === 'supervisor') {
      return caller.displayName;
    }
    return null;
  }

  // GET /api/orders/is-admin — returns { isAdmin, displayName } for the Bearer token holder
  app.get('/api/orders/is-admin', async (req, res) => {
    try {
      const displayName = await resolveOrderAdminFromToken(req.headers.authorization);
      res.json({ success: true, isAdmin: displayName !== null, displayName });
    } catch (error) {
      console.error('Error checking admin status:', error);
      res.status(500).json({ success: false, error: 'Failed to check admin status' });
    }
  });

  // GET /api/orders — list active items (no auth required, fast)
  app.get('/api/orders', async (req, res) => {
    try {
      const items = await storage.getActiveOrderItems();
      res.json({ success: true, items });
    } catch (error) {
      console.error('Error fetching order items:', error);
      res.status(500).json({ success: false, error: 'Failed to fetch order items' });
    }
  });

  // POST /api/orders — add a new item (requires valid staff Bearer token)
  // Identity (addedBy) is derived server-side from Graph /me — not trusted from body.
  app.post('/api/orders', async (req, res) => {
    try {
      const caller = await resolveCallerFromToken(req.headers.authorization);
      if (!caller) {
        return res.status(401).json({ success: false, error: 'Authentication required. Please sign in with your Cranfield Glass account.' });
      }
      const { itemName } = req.body;
      if (!itemName || typeof itemName !== 'string' || itemName.trim().length === 0) {
        return res.status(400).json({ success: false, error: 'itemName is required' });
      }
      const item = await storage.createOrderItem({
        itemName: itemName.trim(),
        addedBy: caller.displayName,
        status: 'active',
      });
      res.json({ success: true, item });
    } catch (error) {
      console.error('Error creating order item:', error);
      res.status(500).json({ success: false, error: 'Failed to create order item' });
    }
  });

  // PATCH /api/orders/:id — mark ordered or archived (admin/supervisor only)
  // Identity & role resolved server-side via Graph /me + staff table — never trusted from body.
  app.patch('/api/orders/:id', async (req, res) => {
    try {
      const id = parseInt(req.params.id);
      if (isNaN(id)) return res.status(400).json({ success: false, error: 'Invalid id' });
      const { status } = req.body;
      if (!status || !['ordered', 'archived'].includes(status)) {
        return res.status(400).json({ success: false, error: 'status must be "ordered" or "archived"' });
      }
      // Resolve caller identity & role from Bearer token — not from request body
      const callerName = await resolveOrderAdminFromToken(req.headers.authorization);
      if (!callerName) {
        return res.status(403).json({ success: false, error: 'Only admin or supervisor staff can mark items as ordered' });
      }
      const item = await storage.updateOrderItemStatus(id, status, callerName);
      res.json({ success: true, item });
    } catch (error) {
      console.error('Error updating order item:', error);
      res.status(500).json({ success: false, error: 'Failed to update order item' });
    }
  });

  const server = createServer(app);
  return server;
}