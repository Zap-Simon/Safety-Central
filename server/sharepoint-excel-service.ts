import { LRUCache } from 'lru-cache';
import * as XLSX from 'xlsx';
import pino from 'pino';
import type { WorkbookData, WorksheetData, SharePointExcelFile } from '../shared/schema.js';

// Use same logger configuration as server
const logger = pino({
  level: process.env.LOG_LEVEL || (process.env.NODE_ENV === 'production' ? 'info' : 'debug'),
  transport: process.env.NODE_ENV === 'development' ? {
    target: 'pino-pretty'
  } : undefined
});

// Cache configuration
const cache = new LRUCache<string, any>({
  max: 100,
  ttl: 15 * 60 * 1000, // 15 minutes
});

interface CachedData {
  data: any;
  etag?: string;
  lastModified?: string;
}

// Function to clear Excel cache for a specific file
export function clearExcelCache(
  siteUrl: string, 
  folderPath: string, 
  fileName: string, 
  worksheetName?: string
) {
  const cacheKey = `excel:${siteUrl}:${folderPath}:${fileName}:${worksheetName || 'all'}`;
  cache.delete(cacheKey);
  logger.debug({ cacheKey }, 'Cleared Excel cache');
}

// Retry configuration
const MAX_RETRIES = 3;
const RETRY_DELAYS = [2000, 4000, 6000]; // 2, 4, 6 seconds

/**
 * Makes a request to Microsoft Graph API with retry logic
 */
async function makeGraphRequest(url: string, accessToken: string, options: RequestInit = {}): Promise<Response> {
  let response: Response | undefined;
  let attempts = 0;

  while (attempts < MAX_RETRIES) {
    attempts++;
    
    response = await fetch(url, {
      ...options,
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json',
        ...options.headers,
      }
    });

    if (response.ok) {
      return response;
    }
    
    // If it's a 503 (Service Unavailable) or 429 (Too Many Requests), wait and retry
    if ((response.status === 503 || response.status === 429) && attempts < MAX_RETRIES) {
      logger.warn({ status: response.status, retryDelay: RETRY_DELAYS[attempts - 1] / 1000, attempt: attempts }, 'SharePoint temporarily unavailable, retrying');
      await new Promise(resolve => setTimeout(resolve, RETRY_DELAYS[attempts - 1]));
      continue;
    }
    
    break;
  }

  // This should never happen, but TypeScript needs to know response is defined
  if (!response) {
    throw new Error('Failed to make request after retries');
  }

  return response;
}

/**
 * Gets SharePoint site information
 */
async function getSiteInfo(accessToken: string, siteUrl: string) {
  if (process.env.NODE_ENV === 'development') {
    console.log(`[DEBUG] Getting site information for: ${siteUrl}`);
  }
  
  const response = await makeGraphRequest(
    `https://graph.microsoft.com/v1.0/sites/${siteUrl}`,
    accessToken
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Site lookup error: ${response.status} - ${errorText}`);
  }

  const siteData = await response.json();
  if (process.env.NODE_ENV === 'development') {
    console.log(`[DEBUG] Site data retrieved successfully, site ID: ${siteData.id}`);
  }
  
  return siteData;
}

/**
 * Discovers Excel files in a SharePoint folder
 */
async function discoverExcelFiles(
  accessToken: string, 
  siteId: string, 
  folderPath: string = ''
): Promise<SharePointExcelFile[]> {
  if (process.env.NODE_ENV === 'development') {
    console.log(`[DEBUG] Discovering Excel files in folder: ${folderPath || 'root'}`);
  }
  
  // Use the same proven URL structure as the working test & tag documents query
  let url: string;
  if (folderPath) {
    // Remove leading slash if present and encode properly
    const cleanPath = folderPath.startsWith('/') ? folderPath.substring(1) : folderPath;
    const encodedPath = encodeURIComponent(cleanPath);
    url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/children?$select=id,name,size,lastModifiedDateTime,webUrl,file&$orderby=lastModifiedDateTime desc`;
  } else {
    url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children?$select=id,name,size,lastModifiedDateTime,webUrl,file&$orderby=lastModifiedDateTime desc`;
  }
  
  if (process.env.NODE_ENV === 'development') {
    console.log(`[DEBUG] Graph API URL for Excel files: ${url}`);
  }
  
  const response = await makeGraphRequest(url, accessToken);
  
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to list files: ${response.status} - ${errorText}`);
  }

  const data = await response.json();
  const files = data.value || [];
  const excelFiles: SharePointExcelFile[] = [];

  if (process.env.NODE_ENV === 'development') {
    console.log(`[DEBUG] Total files found in folder: ${files.length}`);
    console.log('[DEBUG] File details:', files.map((f: any) => ({ name: f.name, hasFile: !!f.file, size: f.size })));
  }

  // Process files and identify Excel types - only process actual files, not folders
  for (const file of files) {
    if (file.file) { // Only include files, not folders
      const name = file.name.toLowerCase();
      let type: 'xlsx' | 'xls' | 'csv' | null = null;
      
      if (name.endsWith('.xlsx')) {
        type = 'xlsx';
      } else if (name.endsWith('.xls')) {
        type = 'xls';
      } else if (name.endsWith('.csv')) {
        type = 'csv';
      }
      
      if (type) {
        if (process.env.NODE_ENV === 'development') {
          console.log(`[DEBUG] Processing Excel file: ${file.name}, ID: ${file.id}, Type: ${type}`);
        }
        excelFiles.push({
          id: file.id,
          name: file.name,
          size: file.size,
          lastModified: file.lastModifiedDateTime,
          webUrl: file.webUrl,
          type
        });
      }
    }
  }

  if (process.env.NODE_ENV === 'development') {
    console.log(`[DEBUG] Found ${excelFiles.length} Excel files:`, excelFiles.map(f => f.name));
  }
  return excelFiles;
}

/**
 * Gets worksheet data using Microsoft Graph Excel API
 */
async function getWorksheetDataViaGraph(
  accessToken: string,
  siteId: string,
  fileId: string,
  worksheetName?: string
): Promise<WorksheetData[]> {
  if (process.env.NODE_ENV === 'development') {
    console.log(`[DEBUG] Fetching worksheet data via Graph API for file: ${fileId}, worksheet: ${worksheetName || 'all'}`);
  }
  
  // Get all worksheets
  const worksheetsUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/workbook/worksheets`;
  const worksheetsResponse = await makeGraphRequest(worksheetsUrl, accessToken);
  
  if (!worksheetsResponse.ok) {
    const errorText = await worksheetsResponse.text();
    throw new Error(`Failed to get worksheets: ${worksheetsResponse.status} - ${errorText}`);
  }

  const worksheetsData = await worksheetsResponse.json();
  const worksheets = worksheetsData.value;
  
  if (worksheets.length === 0) {
    return [];
  }

  const worksheetDataList: WorksheetData[] = [];
  
  // Filter to specific worksheet if provided
  const targetWorksheets = worksheetName 
    ? worksheets.filter((ws: any) => ws.name === worksheetName)
    : worksheets;

  // Process worksheets in parallel for better performance
  const worksheetPromises = targetWorksheets.map(async (worksheet: any) => {
    try {
      // Get used range for this worksheet
      const rangeUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/workbook/worksheets('${worksheet.name}')/usedRange(valuesOnly=true)`;
      const rangeResponse = await makeGraphRequest(rangeUrl, accessToken);
      
      if (!rangeResponse.ok) {
        if (process.env.NODE_ENV === 'development') {
          console.warn(`[DEBUG] Failed to get range for worksheet ${worksheet.name}, skipping`);
        }
        return null;
      }

      const rangeData = await rangeResponse.json();
      const values = rangeData.values;
      
      if (!values || values.length === 0) {
        if (process.env.NODE_ENV === 'development') {
          console.warn(`[DEBUG] No data in worksheet ${worksheet.name}, skipping`);
        }
        return null;
      }

      // First row is headers
      const headers = values[0].map((header: any) => String(header || ''));
      const rows = values.slice(1).map((row: any[]) => 
        row.map((cell: any) => {
          if (cell === null || cell === undefined || cell === '') {
            return null;
          }
          // Handle Excel formula errors
          if (typeof cell === 'string' && (cell.startsWith('#') || cell === '#VALUE!' || cell === '#REF!' || cell === '#N/A' || cell === '#DIV/0!' || cell === '#NULL!' || cell === '#NUM!' || cell === '#NAME?')) {
            return ''; // Return empty string instead of error
          }
          return cell;
        })
      );

      if (process.env.NODE_ENV === 'development') {
        console.log(`[DEBUG] Successfully processed worksheet: ${worksheet.name} (${headers.length} columns, ${rows.length} rows)`);
      }
      
      return {
        name: worksheet.name,
        headers,
        rows
      };
      
    } catch (error) {
      logger.error({ err: error, worksheetName: worksheet.name }, 'Error processing worksheet');
      return null;
    }
  });

  // Wait for all worksheets to be processed in parallel
  const worksheetResults = await Promise.all(worksheetPromises);
  
  // Filter out null results (failed worksheets)
  for (const result of worksheetResults) {
    if (result) {
      worksheetDataList.push(result);
    }
  }

  return worksheetDataList;
}

/**
 * Gets Excel data by downloading and parsing with xlsx library (fallback)
 */
async function getExcelDataViaDownload(
  accessToken: string,
  siteId: string,
  fileId: string,
  fileName: string
): Promise<WorksheetData[]> {
  if (process.env.NODE_ENV === 'development') {
    console.log(`[DEBUG] Downloading and parsing Excel file: ${fileName}`);
  }
  
  // Download the file
  const downloadUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${fileId}/content`;
  const response = await makeGraphRequest(downloadUrl, accessToken);
  
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to download file: ${response.status} - ${errorText}`);
  }

  const buffer = await response.arrayBuffer();
  
  try {
    // Parse with xlsx library
    const workbook = XLSX.read(buffer, { type: 'array' });
    const worksheetDataList: WorksheetData[] = [];
    
    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1, 
        defval: null,
        raw: false // Convert everything to strings to avoid date issues
      });
      
      if (jsonData.length === 0) {
        continue;
      }

      const headers = (jsonData[0] as any[]).map((header: any) => String(header || ''));
      const rows = jsonData.slice(1).map((row: any) => {
        if (Array.isArray(row)) {
          return row.map((cell: any) => {
            if (cell === null || cell === undefined || cell === '') {
              return null;
            }
            // Handle Excel formula errors
            if (typeof cell === 'string' && (cell.startsWith('#') || cell === '#VALUE!' || cell === '#REF!' || cell === '#N/A' || cell === '#DIV/0!' || cell === '#NULL!' || cell === '#NUM!' || cell === '#NAME?')) {
              return ''; // Return empty string instead of error
            }
            // Try to convert to number if it looks like one
            if (typeof cell === 'string' && !isNaN(Number(cell)) && cell.trim() !== '') {
              return Number(cell);
            }
            return cell;
          });
        }
        return [];
      });

      worksheetDataList.push({
        name: sheetName,
        headers,
        rows
      });
      
      if (process.env.NODE_ENV === 'development') {
        console.log(`[DEBUG] Successfully parsed worksheet: ${sheetName} (${headers.length} columns, ${rows.length} rows)`);
      }
    }

    return worksheetDataList;
    
  } catch (error) {
    logger.error({ err: error }, 'Error parsing Excel file');
    throw new Error(`Failed to parse Excel file: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * Main function to get Excel data with caching
 */
export async function getExcelData(
  accessToken: string,
  siteUrl: string,
  folderPath: string,
  fileName: string,
  worksheetName?: string
): Promise<WorkbookData> {
  const cacheKey = `excel:${siteUrl}:${folderPath}:${fileName}:${worksheetName || 'all'}`;
  const fileDiscoveryCacheKey = `files:${siteUrl}:${folderPath}`;
  
  // Check cache first (but allow cache bypass)
  const cached = cache.get(cacheKey) as CachedData;
  if (cached) {
    if (process.env.NODE_ENV === 'development') {
      console.log(`[DEBUG] Returning cached Excel data for: ${fileName}`);
    }
    return cached.data;
  }

  try {
    // Try to get cached site info and file discovery first
    let siteData;
    let excelFiles;
    
    const cachedFiles = cache.get(fileDiscoveryCacheKey) as CachedData;
    if (cachedFiles && cachedFiles.data) {
      if (process.env.NODE_ENV === 'development') {
        console.log(`[DEBUG] Using cached file discovery for folder: ${folderPath}`);
      }
      siteData = cachedFiles.data.siteData;
      excelFiles = cachedFiles.data.excelFiles;
    } else {
      // Get site info and discover files
      siteData = await getSiteInfo(accessToken, siteUrl);
      excelFiles = await discoverExcelFiles(accessToken, siteData.id, folderPath);
      
      // Cache file discovery separately (longer TTL since files change less frequently)
      cache.set(fileDiscoveryCacheKey, { 
        data: { siteData, excelFiles } 
      });
    }
    
    if (process.env.NODE_ENV === 'development') {
      console.log(`[DEBUG] Looking for file: ${fileName} among ${excelFiles.length} Excel files`);
      console.log('[DEBUG] Available Excel files:', excelFiles.map((f: SharePointExcelFile) => ({ name: f.name, id: f.id, type: f.type })));
    }
    
    const targetFile = excelFiles.find((file: SharePointExcelFile) => file.name === fileName);
    
    if (!targetFile) {
      throw new Error(`Excel file not found: ${fileName}. Available files: ${excelFiles.map((f: SharePointExcelFile) => f.name).join(', ')}`);
    }
    
    if (process.env.NODE_ENV === 'development') {
      console.log(`[DEBUG] Target file found:`, { name: targetFile.name, id: targetFile.id, type: targetFile.type });
    }

    let worksheetDataList: WorksheetData[];
    
    // Try Graph Excel API first for .xlsx files
    if (targetFile.type === 'xlsx') {
      try {
        worksheetDataList = await getWorksheetDataViaGraph(
          accessToken,
          siteData.id,
          targetFile.id,
          worksheetName
        );
      } catch (error) {
        if (process.env.NODE_ENV === 'development') {
          console.warn('[DEBUG] Graph Excel API failed, falling back to download method:', error);
        }
        worksheetDataList = await getExcelDataViaDownload(
          accessToken,
          siteData.id,
          targetFile.id,
          fileName
        );
      }
    } else {
      // Use download method for .xls and .csv files
      worksheetDataList = await getExcelDataViaDownload(
        accessToken,
        siteData.id,
        targetFile.id,
        fileName
      );
    }

    const workbookData: WorkbookData = {
      fileId: targetFile.id,
      name: targetFile.name,
      lastModified: targetFile.lastModified,
      sheets: worksheetDataList
    };

    // Cache the result
    cache.set(cacheKey, { data: workbookData });
    
    if (process.env.NODE_ENV === 'development') {
      console.log(`[DEBUG] Successfully processed Excel file: ${fileName} with ${worksheetDataList.length} worksheets`);
    }
    return workbookData;
    
  } catch (error) {
    logger.error({ err: error }, 'Error getting Excel data');
    throw error;
  }
}

/**
 * Lists Excel files in a SharePoint folder
 */
export async function listExcelFiles(
  accessToken: string,
  siteUrl: string,
  folderPath: string = ''
): Promise<SharePointExcelFile[]> {
  const cacheKey = `excel-files:${siteUrl}:${folderPath}`;
  
  // Check cache first
  const cached = cache.get(cacheKey) as CachedData;
  if (cached) {
    if (process.env.NODE_ENV === 'development') {
      console.log(`[DEBUG] Returning cached Excel files list for: ${folderPath || 'root'}`);
    }
    return cached.data;
  }

  try {
    // Get site info
    const siteData = await getSiteInfo(accessToken, siteUrl);
    
    // Get Excel files
    const excelFiles = await discoverExcelFiles(accessToken, siteData.id, folderPath);
    
    // Cache the result for a shorter time since file lists change more frequently
    cache.set(cacheKey, { data: excelFiles }, { ttl: 5 * 60 * 1000 }); // 5 minutes
    
    return excelFiles;
    
  } catch (error) {
    logger.error({ err: error }, 'Error listing Excel files');
    throw error;
  }
}