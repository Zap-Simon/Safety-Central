/**
 * SharePoint Secrets Configuration Helper
 * 
 * This module provides utilities to validate and manage Microsoft Graph API
 * credentials for connecting to SharePoint lists across multiple sites.
 * 
 * Required Azure App Registration Permissions:
 * - Sites.ReadWrite.All (for SharePoint list access)
 * - User.Read.All (for resolving SharePoint person fields like NameLookupId)
 * - Sites.Selected (alternative to Sites.ReadWrite.All if using specific site permissions)
 * 
 * Without User.Read.All permission, NameLookupId values will show as numbers
 * instead of actual staff names in the "Submitted by" and "Assigned to" fields.
 */

export interface AzureCredentials {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  redirectUri: string;
}

export interface SharePointSites {
  safetyIdeasSiteId: string;
  safetyIdeasListId: string;
  businessIdeasListId: string;
  accidentReportsSiteId: string;
  accidentReportsListId: string;
  actionsSiteId?: string;
  actionsListId?: string;
}

export interface SharePointConfig extends AzureCredentials, SharePointSites {}

/**
 * Validates that required Azure credentials are present
 */
export function validateAzureCredentials(): { valid: boolean; missing: string[] } {
  const required = ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'AZURE_REDIRECT_URI'];
  const missing = required.filter(key => !process.env[key]);
  
  return {
    valid: missing.length === 0,
    missing
  };
}

/**
 * Validates that required SharePoint site and list IDs are present
 */
export function validateSharePointConfiguration(): { valid: boolean; missing: string[] } {
  const required = [
    'SAFETY_IDEAS_SITE_ID',
    'SAFETY_IDEAS_LIST_ID', 
    'BUSINESS_IDEAS_LIST_ID',
    'ACCIDENT_REPORTS_SITE_ID',
    'ACCIDENT_REPORTS_LIST_ID'
  ];
  
  const missing = required.filter(key => !process.env[key]);
  
  return {
    valid: missing.length === 0,
    missing
  };
}

/**
 * Gets the complete SharePoint configuration from environment variables
 */
export function getSharePointConfig(): SharePointConfig | null {
  const azureValidation = validateAzureCredentials();
  const sharePointValidation = validateSharePointConfiguration();
  
  if (!azureValidation.valid || !sharePointValidation.valid) {
    return null;
  }
  
  return {
    tenantId: process.env.AZURE_TENANT_ID!,
    clientId: process.env.AZURE_CLIENT_ID!,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
    redirectUri: process.env.AZURE_REDIRECT_URI!,
    safetyIdeasSiteId: process.env.SAFETY_IDEAS_SITE_ID!,
    safetyIdeasListId: process.env.SAFETY_IDEAS_LIST_ID!,
    businessIdeasListId: process.env.BUSINESS_IDEAS_LIST_ID!,
    accidentReportsSiteId: process.env.ACCIDENT_REPORTS_SITE_ID!,
    accidentReportsListId: process.env.ACCIDENT_REPORTS_LIST_ID!,
    actionsSiteId: process.env.ACTIONS_SITE_ID,
    actionsListId: process.env.ACTIONS_LIST_ID,
  };
}

/**
 * Returns configuration status for debugging and setup
 */
export function getConfigurationStatus() {
  const azure = validateAzureCredentials();
  const sharePoint = validateSharePointConfiguration();
  
  return {
    azure: {
      configured: azure.valid,
      missing: azure.missing
    },
    sharePoint: {
      configured: sharePoint.valid,
      missing: sharePoint.missing
    },
    ready: azure.valid && sharePoint.valid
  };
}

/**
 * Pretty prints the setup instructions for missing secrets
 */
export function printSetupInstructions() {
  const status = getConfigurationStatus();
  
  if (status.ready) {
    console.log('✅ SharePoint configuration is complete and ready!');
    return;
  }
  
  console.log('🔐 SharePoint Setup Required\n');
  
  if (!status.azure.configured) {
    console.log('Missing Azure credentials:');
    status.azure.missing.forEach(key => console.log(`  - ${key}`));
    console.log('');
  }
  
  if (!status.sharePoint.configured) {
    console.log('Missing SharePoint configuration:');
    status.sharePoint.missing.forEach(key => console.log(`  - ${key}`));
    console.log('');
  }
  
  console.log('Run: node setup-secrets.js for detailed setup instructions');
}