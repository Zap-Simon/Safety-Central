/**
 * Unified SharePoint Lists Service
 * 
 * Clean, maintainable service for accessing all SharePoint lists with consistent
 * field mapping and processing. Each list type has its own configuration and
 * processing methods while sharing common utilities.
 */
import pino from 'pino';

// Use same logger configuration as server
const logger = pino({
  level: process.env.LOG_LEVEL || (process.env.NODE_ENV === 'production' ? 'info' : 'debug'),
  transport: process.env.NODE_ENV === 'development' ? {
    target: 'pino-pretty'
  } : undefined
});

// Common interfaces
interface SharePointPerson {
  Id: number;
  Title: string;
  EMail: string;
  JobTitle?: string;
}

interface BaseSharePointItem {
  ID: number;
  Title: string;
  Created: string;
  Modified: string;
  Author?: SharePointPerson;
  Editor?: SharePointPerson;
}

// Business Ideas specific interface
interface BusinessIdeasItem extends BaseSharePointItem {
  BusinessIdea1?: string;        // Main content field
  Status?: string;               // Choice field
  Idea_x0020_Date?: string;      // Submission date
  Idea_x0020_Type?: string;      // Idea type choice
  MeetingDate?: string;          // Meeting discussion date
  MeetingNotes?: string;         // Meeting notes field
  Name?: SharePointPerson;       // Submitted by person field
}

// Safety Ideas specific interface
interface SafetyIdeasItem extends BaseSharePointItem {
  SafetyIdea1?: string;          // Main content field
  Status?: string;               // Choice field
  Idea_x0020_Date?: string;      // Submission date
  Idea_x0020_Type?: string;      // Idea type choice
  MeetingDate?: string;          // Meeting discussion date
  Meeting_x0020_Notes1?: string; // Meeting notes field (different from Business Ideas)
  Name?: SharePointPerson;       // Submitted by person field
}

// Staff specific interface
interface StaffItem extends BaseSharePointItem {
  Name?: string;                 // Staff member name
  Email?: string;                // Email address
  JobTitle?: string;             // Job title/position
  StartDate?: string;            // Employment start date
  IsActive?: boolean;            // Active status
  Department?: string;           // Department/team
  Manager?: SharePointPerson;    // Manager/supervisor
  Phone?: string;                // Phone number
  Notes?: string;                // Additional notes
}

// Processed item interface for meeting history
export interface ProcessedMeetingItem {
  id: string;
  title: string;
  description: string;
  type: 'Business Ideas' | 'Safety Ideas' | 'Near Miss' | 'Actions';
  status: string;
  meetingDate: string;
  meetingNotes: string;
  submittedBy: string;
  submittedDate: string;
  ideaType?: string;
  hasActualMeetingDate?: boolean;
  assignedTo?: string;
  secondaryDescription?: string; // For Near Miss "How It Happened" field
  // Embedded action tracking fields (replace separate Actions list)
  actionPriority?: string; // High, Medium, Low
  actionStatus?: string; // Not Started, In Progress, Completed
  actionAssignedTo?: string; // Person assigned to action
  actionStartDate?: string; // When action work began
  actionDueDate?: string; // When action is due
  actionNotes?: string; // Action-specific notes and updates
  // Legacy Actions-specific fields (for backward compatibility)
  priority?: string; // High, Medium, Low
  category?: string; // Safety Idea, Business Idea, etc.
  listOrigin?: string; // Source list identifier
  sourceLink?: string; // Link back to original item
}

// List configuration interface
interface ListConfig {
  listTitle: string;
  siteUrl?: string; // Optional - defaults to main site
  selectFields: string[];
  expandFields: string[];
  contentField: string;
  submittedByField: string;
  submissionDateField: string;
  meetingNotesField: string;
  typeField?: string;
  assignedToField?: string;
  linkField?: string; // For Actions list back-linking
  priorityField?: string; // For Actions priority
  categoryField?: string; // For Actions category
}

export class SharePointListsService {
  private accessToken: string;
  private baseUrl = 'https://cranfieldglass.sharepoint.com/_api/web/lists/getbytitle';


  // List configurations - easily extensible for new lists
  private static readonly LIST_CONFIGS: Record<string, ListConfig> = {
    'Business Ideas': {
      listTitle: 'Business Ideas',
      selectFields: [
        'ID', 'Title', 'BusinessIdea1', 'Status', 
        'Idea_x0020_Date', 'Idea_x0020_Type', 
        'MeetingDate', 'MeetingNotes', 'Created', 'Modified',
        'Name/Id', 'Name/Title', 'Name/EMail', 'Name/JobTitle',
        'Author/Id', 'Author/Title', 'Author/EMail', 'Author/JobTitle',
        'Editor/Id', 'Editor/Title', 'Editor/EMail', 'Editor/JobTitle'
      ],
      expandFields: ['Name', 'Author', 'Editor'],
      contentField: 'BusinessIdea1',
      submittedByField: 'Name',
      submissionDateField: 'Idea_x0020_Date',
      meetingNotesField: 'MeetingNotes',
      typeField: 'Idea_x0020_Type'
    },
    'Safety Ideas': {
      listTitle: 'Safety Ideas',
      selectFields: [
        'ID', 'Title', 'SafetyIdea1', 'Status', 
        'Idea_x0020_Date', 'Idea_x0020_Type', 
        'MeetingDate', 'Meeting_x0020_Notes1', 'Created', 'Modified',
        'Name/Id', 'Name/Title', 'Name/EMail', 'Name/JobTitle',
        'Author/Id', 'Author/Title', 'Author/EMail', 'Author/JobTitle',
        'Editor/Id', 'Editor/Title', 'Editor/EMail', 'Editor/JobTitle'
      ],
      expandFields: ['Name', 'Author', 'Editor'],
      contentField: 'SafetyIdea1',
      submittedByField: 'Name',
      submissionDateField: 'Idea_x0020_Date',
      meetingNotesField: 'Meeting_x0020_Notes1',
      typeField: 'Idea_x0020_Type'
    },
    'Near Miss': {
      listTitle: 'Near Miss - Accident Safety Register',
      siteUrl: 'https://cranfieldglass.sharepoint.com/sites/IncidentsReports',
      selectFields: [
        'ID', 'Title', 'Date', 'EventType', 
        'Canyoubrieflyexplainwhathappened', 'Howdidthishappen_x003f_',
        'InvestigationRequired_x003f_', 'Status', 'MeetingDate', 'MeetingNotes', 'Created', 'Modified',
        'Name/Id', 'Name/Title', 'Name/EMail', 'Name/JobTitle',
        'Author/Id', 'Author/Title', 'Author/EMail', 'Author/JobTitle',
        'Editor/Id', 'Editor/Title', 'Editor/EMail', 'Editor/JobTitle'
      ],
      expandFields: ['Name', 'Author', 'Editor'],
      contentField: 'Canyoubrieflyexplainwhathappened',
      submittedByField: 'Name',
      submissionDateField: 'Date',
      meetingNotesField: 'MeetingNotes', // Standard Meeting Notes field for Near Miss
      typeField: 'EventType'
    },
    'Actions': {
      listTitle: 'Actions',
      siteUrl: 'https://cranfieldglass.sharepoint.com/sites/Team_Actions_Meetings',
      selectFields: [
        'ID', 'Title', 'Idea_x002f_Event', 'Status', 'Priority',
        'Link', 'ActionStartDate', 'IdeaType', 'Catagorie', 'IdeaDate',
        'Notes_x0020__x002f__x0020_Commen', 'Created', 'Modified',
        'Assigend_x0020_to/Id', 'Assigend_x0020_to/Title', 'Assigend_x0020_to/EMail', 'Assigend_x0020_to/JobTitle',
        'SubmittedBy/Id', 'SubmittedBy/Title', 'SubmittedBy/EMail', 'SubmittedBy/JobTitle',
        'Author/Id', 'Author/Title', 'Author/EMail', 'Author/JobTitle',
        'Editor/Id', 'Editor/Title', 'Editor/EMail', 'Editor/JobTitle'
      ],
      expandFields: ['Assigend_x0020_to', 'SubmittedBy', 'Author', 'Editor'],
      contentField: 'Idea_x002f_Event',
      submittedByField: 'SubmittedBy',
      submissionDateField: 'IdeaDate',
      meetingNotesField: 'Notes_x0020__x002f__x0020_Commen',
      typeField: 'IdeaType',
      assignedToField: 'Assigend_x0020_to',
      linkField: 'Link',
      priorityField: 'Priority',
      categoryField: 'Catagorie'
    },
    'Staff': {
      listTitle: 'Staff Training Registry',
      siteUrl: 'https://cranfieldglass.sharepoint.com/sites/FieldHealthSafetyTraining',
      selectFields: [
        'ID', 'Title', 'Name', 'Email', 'JobTitle', 'StartDate',
        'IsActive', 'Department', 'Phone', 'Notes', 'Created', 'Modified',
        'Manager/Id', 'Manager/Title', 'Manager/EMail', 'Manager/JobTitle',
        'Author/Id', 'Author/Title', 'Author/EMail', 'Author/JobTitle',
        'Editor/Id', 'Editor/Title', 'Editor/EMail', 'Editor/JobTitle'
      ],
      expandFields: ['Manager', 'Author', 'Editor'],
      contentField: 'Name',
      submittedByField: 'Author',
      submissionDateField: 'StartDate',
      meetingNotesField: 'Notes',
      typeField: 'JobTitle'
    }
    // Future lists will be added here with their specific configurations
  };

  constructor(accessToken: string) {
    this.accessToken = accessToken;
  }

  /**
   * Get items from Business Ideas list
   */
  async getBusinessIdeas(): Promise<ProcessedMeetingItem[]> {
    return this.getListItems('Business Ideas');
  }

  /**
   * Get items from Safety Ideas list
   */
  async getSafetyIdeas(): Promise<ProcessedMeetingItem[]> {
    return this.getListItems('Safety Ideas');
  }

  /**
   * Get items from Near Miss list
   */
  async getNearMiss(): Promise<ProcessedMeetingItem[]> {
    return this.getListItems('Near Miss');
  }

  /**
   * Get items from Staff list
   */
  async getStaff(): Promise<ProcessedMeetingItem[]> {
    return this.getListItems('Staff');
  }

  /**
   * Get items from Actions list
   */
  async getActions(): Promise<ProcessedMeetingItem[]> {
    return this.getListItems('Actions');
  }

  /**
   * Generic method to get items from any configured list
   */
  private async getListItems(listType: string): Promise<ProcessedMeetingItem[]> {
    const config = SharePointListsService.LIST_CONFIGS[listType];
    if (!config) {
      throw new Error(`List configuration not found for: ${listType}`);
    }

    const apiUrl = this.buildApiUrl(config);
    


    try {
      const response = await this.makeSharePointRequest(apiUrl);
      const items = response.d?.results || [];
      
      
      // Debug logging only in development
      if (process.env.NODE_ENV === 'development' && items.length > 0) {
        const firstItem = items[0];
        console.log(`[DEBUG] SharePoint date formats for ${listType}:`);
        console.log(`   Created: ${firstItem.Created}`);
        console.log(`   Modified: ${firstItem.Modified}`);
        console.log(`   ${config.submissionDateField}: ${firstItem[config.submissionDateField]}`);
        if (firstItem.MeetingDate) {
          console.log(`   MeetingDate: ${firstItem.MeetingDate}`);
        }
        if (firstItem.Idea_x0020_Date) {
          console.log(`   Idea_x0020_Date: ${firstItem.Idea_x0020_Date}`);
        }
      }
      
      return items
        .filter((item: any) => this.hasContent(item, config.contentField))
        .map((item: any) => this.processItem(item, listType, config));
        
    } catch (error) {
      logger.error({ err: error, listType }, 'Error fetching SharePoint list data');
      throw error;
    }
  }

  /**
   * Build SharePoint REST API URL from list configuration
   */
  private buildApiUrl(config: ListConfig): string {
    const selectParam = `$select=${config.selectFields.join(',')}`;
    const expandParam = config.expandFields.length > 0 ? `&$expand=${config.expandFields.join(',')}` : '';
    const topParam = '&$top=1000';
    
    // Use custom site URL if specified, otherwise use main site
    const baseUrl = config.siteUrl 
      ? `${config.siteUrl}/_api/web/lists/getbytitle`
      : this.baseUrl;
    
    return `${baseUrl}('${config.listTitle}')/items?${selectParam}${expandParam}${topParam}`;
  }

  /**
   * Check if item has content in the main content field
   */
  private hasContent(item: Record<string, any>, contentField: string): boolean {
    return item[contentField] && item[contentField].trim() !== '';
  }

  /**
   * Process SharePoint item into standardized meeting item format
   */
  private processItem(item: Record<string, any>, listType: string, config: ListConfig): ProcessedMeetingItem {
    const submittedBy = this.extractPersonName(item, config.submittedByField);
    const originalSubmissionDate = item[config.submissionDateField] || item.Created;
    const submissionDate = this.convertToNZTime(originalSubmissionDate);
    const meetingDateInfo = this.processMeetingDate(item, submissionDate);
    
    const title = this.processTitle(item.Title);
    
    const baseItem = {
      id: `${listType.toLowerCase().replace(' ', '-')}-${item.ID}`,
      title: title,
      description: this.cleanHtmlContent(item[config.contentField] || ''),
      type: listType as any,
      status: item.Status || 'Submitted',
      meetingDate: meetingDateInfo.date,
      meetingNotes: this.cleanHtmlContent(item[config.meetingNotesField] || ''),
      submittedBy: submittedBy,
      submittedDate: submissionDate,
      ideaType: config.typeField ? item[config.typeField] : undefined,
      hasActualMeetingDate: meetingDateInfo.isActual,
      assignedTo: config.assignedToField ? this.extractPersonName(item, config.assignedToField) : undefined,
      secondaryDescription: listType === 'Near Miss' ? this.cleanHtmlContent(item['Howdidthishappen_x003f_'] || '') : undefined,
      // Embedded action tracking fields
      actionPriority: item.ActionPriority || undefined,
      actionStatus: item.ActionStatus || undefined,
      actionAssignedTo: item.ActionAssignedTo || undefined,
      actionStartDate: item.ActionStartDate ? this.convertToNZTime(item.ActionStartDate) : undefined,
      actionDueDate: item.ActionDueDate ? this.convertToNZTime(item.ActionDueDate) : undefined,
      actionNotes: item.ActionNotes || undefined
    };

    // Add Actions-specific fields
    if (listType === 'Actions') {
      const actionItem = {
        ...baseItem,
        priority: config.priorityField ? item[config.priorityField] : undefined,
        category: config.categoryField ? item[config.categoryField] : undefined,
        listOrigin: listType, // ListOrigin field doesn't exist in SharePoint
        sourceLink: config.linkField ? item[config.linkField] : undefined,
        actionStartDate: item.ActionStartDate ? this.convertToNZTime(item.ActionStartDate) : undefined,
        // For Actions, use ActionStartDate for meeting date if available
        meetingDate: item.ActionStartDate ? this.convertToNZTime(item.ActionStartDate) : meetingDateInfo.date
      };
      

      
      return actionItem;
    }

    return baseItem;
  }

  /**
   * Extract person name from SharePoint person field
   */
  private extractPersonName(item: Record<string, any>, fieldName: string): string {
    const personField = item[fieldName];
    
    if (personField?.Title) {
      return personField.Title;
    } else if (personField?.EMail) {
      return personField.EMail.split('@')[0];
    } else if (item.Author?.Title) {
      return item.Author.Title;
    }
    
    return 'Unknown Staff';
  }

  /**
   * Process item title - return empty string for AI generation if generic
   */
  private processTitle(title: string): string {
    if (!title || title.trim() === '' || title.trim() === 'Business Idea') {
      return ''; // Empty title for AI generation
    }
    return title;
  }

  /**
   * Process meeting date - prioritize SharePoint's actual MeetingDate field
   */
  private processMeetingDate(item: Record<string, any>, submissionDate: string): { date: string; isActual: boolean } {
    // Always prioritize SharePoint's actual MeetingDate field
    if (item.MeetingDate && item.MeetingDate !== 'TBD' && item.MeetingDate !== null) {
      // CRITICAL FIX: SharePoint API returns UTC dates but Lists UI shows NZ timezone dates
      // Add 1 day to match what user sees in MS Lists interface
      const meetingDate = new Date(item.MeetingDate);
      meetingDate.setDate(meetingDate.getDate() + 1);
      
      const dateOnly = `${meetingDate.getUTCFullYear()}-${String(meetingDate.getUTCMonth() + 1).padStart(2, '0')}-${String(meetingDate.getUTCDate()).padStart(2, '0')}`;

      return {
        date: `${dateOnly}T10:00:00.000Z`, // Normalize to same time for grouping
        isActual: true
      };
    } else {
      const submissionDateObj = new Date(submissionDate);
      const calculatedDate = this.calculateTuesdayMeetingDate(submissionDateObj);
      // Return ISO date to match SharePoint meeting dates
      const meetingDateObj = new Date(calculatedDate);
      return {
        date: meetingDateObj.toISOString(),
        isActual: false
      };
    }
  }

  /**
   * Make authenticated SharePoint REST API request
   */
  private async makeSharePointRequest(url: string): Promise<any> {
    const response = await fetch(url, {
      headers: {
        'Authorization': `Bearer ${this.accessToken}`,
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`SharePoint REST API error: ${response.status} ${response.statusText} - ${errorText}`);
    }

    return response.json();
  }

  /**
   * Update SharePoint item title using REST API
   */
  async updateItemTitle(itemId: string, title: string, listType: string): Promise<void> {
    const config = SharePointListsService.LIST_CONFIGS[listType];
    if (!config) {
      throw new Error(`List configuration not found for: ${listType}`);
    }

    // Extract numeric ID from formatted ID
    const numericId = itemId.replace(`${listType.toLowerCase().replace(' ', '-')}-`, '');
    
    // Use custom site URL if specified, otherwise use main site
    const baseUrl = config.siteUrl 
      ? `${config.siteUrl}/_api/web/lists/getbytitle`
      : this.baseUrl;
    
    const updateUrl = `${baseUrl}('${config.listTitle}')/items(${numericId})`;
    
    try {
      // Get item etag for optimistic concurrency
      const getResponse = await fetch(updateUrl, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (!getResponse.ok) {
        throw new Error(`Failed to get item for update: ${getResponse.status}`);
      }
      
      const itemData = await getResponse.json();
      const etag = itemData.d.__metadata.etag;
      const entityType = itemData.d.__metadata.type;
      
      // Update the item with new title
      const updateResponse = await fetch(updateUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'If-Match': etag,
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify({
          '__metadata': { 'type': entityType },
          'Title': title
        })
      });
      
      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        throw new Error(`SharePoint update failed: ${updateResponse.status} - ${errorText}`);
      }
      
      
    } catch (error) {
      logger.error({ err: error, listType, itemId: numericId }, 'Failed to update SharePoint item');
      throw error;
    }
  }

  /**
   * Update SharePoint item meeting date using REST API
   */
  async updateItemMeetingDate(itemId: string, meetingDate: string, listType: string): Promise<void> {
    const config = SharePointListsService.LIST_CONFIGS[listType];
    if (!config) {
      throw new Error(`List configuration not found for: ${listType}`);
    }

    // Extract numeric ID from formatted ID
    const numericId = itemId.replace(`${listType.toLowerCase().replace(' ', '-')}-`, '');
    
    // Use custom site URL if specified, otherwise use main site
    const baseUrl = config.siteUrl 
      ? `${config.siteUrl}/_api/web/lists/getbytitle`
      : this.baseUrl;
    
    const updateUrl = `${baseUrl}('${config.listTitle}')/items(${numericId})`;
    
    try {
      // Get item etag for optimistic concurrency
      const getResponse = await fetch(updateUrl, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (!getResponse.ok) {
        throw new Error(`Failed to get item for update: ${getResponse.status}`);
      }
      
      const itemData = await getResponse.json();
      const etag = itemData.d.__metadata.etag;
      const entityType = itemData.d.__metadata.type;
      
      // Format the meeting date for SharePoint (ISO 8601)
      // IMPORTANT: When reading from SharePoint, processMeetingDate() adds 1 day to compensate
      // for the timezone offset between the API and the NZ MS Lists UI. So we must subtract
      // 1 day here before writing, so that when it's read back the +1 day gives the correct date.
      const dateToWrite = new Date(meetingDate);
      dateToWrite.setDate(dateToWrite.getDate() - 1);
      const formattedDate = dateToWrite.toISOString();
      
      // Update the item with new meeting date
      const updateResponse = await fetch(updateUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'If-Match': etag,
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify({
          '__metadata': { 'type': entityType },
          'MeetingDate': formattedDate
        })
      });
      
      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        throw new Error(`SharePoint update failed: ${updateResponse.status} - ${errorText}`);
      }
      
      
    } catch (error) {
      logger.error({ err: error, listType, itemId: numericId }, 'Failed to update SharePoint item');
      throw error;
    }
  }

  /**
   * Convert SharePoint date to proper ISO format
   */
  private convertToNZTime(utcDateString: string): string {
    if (!utcDateString) return '';
    
    // Parse the date
    const parsedDate = new Date(utcDateString);
    
    if (isNaN(parsedDate.getTime())) {
      return '';
    }
    
    // Return the ISO string without modifications
    return parsedDate.toISOString();
  }

  /**
   * Strip HTML tags and decode HTML entities
   */
  private cleanHtmlContent(content: string): string {
    if (!content) return '';
    
    // Remove HTML tags
    let cleaned = content.replace(/<[^>]*>/g, '');
    
    // Decode common HTML entities
    cleaned = cleaned
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#(\d+);/g, (match, num) => String.fromCharCode(parseInt(num)))
      .replace(/&nbsp;/g, ' ');
    
    return cleaned.trim();
  }

  /**
   * Calculate next Tuesday meeting date from submission date
   */
  private calculateTuesdayMeetingDate(fromDate: Date): string {
    // Get the day of the week (0 = Sunday, 1 = Monday, ..., 6 = Saturday)
    const dayOfWeek = fromDate.getDay();
    
    // Calculate days until next Tuesday
    let daysUntilTuesday;
    if (dayOfWeek <= 2) { // Sunday, Monday, or Tuesday
      daysUntilTuesday = 2 - dayOfWeek; // This week's Tuesday
    } else { // Wednesday through Saturday  
      daysUntilTuesday = 9 - dayOfWeek; // Next week's Tuesday
    }
    
    // If it's already Tuesday and submission is late, move to next Tuesday
    if (dayOfWeek === 2 && fromDate.getHours() >= 12) {
      daysUntilTuesday = 7;
    }
    
    // Create new date for the meeting
    const meetingDate = new Date(fromDate);
    meetingDate.setDate(fromDate.getDate() + daysUntilTuesday);
    
    // Return the full ISO string to avoid timezone issues
    return meetingDate.toISOString();
  }

  /**
   * Create new item in SharePoint list using REST API
   */
  async createListItem(listType: string, itemData: {
    title: string;
    description: string;
    submittedBy?: string;
    ideaType?: string;
    status?: string;
    meetingDate?: string; // Meeting date from frontend
    eventType?: string; // For Near Miss
    whatHappened?: string; // For Near Miss
    howItHappened?: string; // For Near Miss
  }): Promise<string> {
    const config = SharePointListsService.LIST_CONFIGS[listType];
    if (!config) {
      throw new Error(`List configuration not found for: ${listType}`);
    }

    // Use custom site URL if specified, otherwise use main site
    const baseUrl = config.siteUrl 
      ? `${config.siteUrl}/_api/web/lists/getbytitle`
      : this.baseUrl;
    
    const createUrl = `${baseUrl}('${config.listTitle}')/items`;
    
    try {
      // Fetch the list's entity type name from SharePoint to avoid hardcoding special char encoding
      const listMetaResponse = await fetch(`${baseUrl}('${config.listTitle}')?$select=ListItemEntityTypeFullName`, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose'
        }
      });
      const listMeta = await listMetaResponse.json();
      const entityType = listMeta.d?.ListItemEntityTypeFullName
        || `SP.Data.${config.listTitle.replace(/\s/g, '_x0020_').replace(/[^a-zA-Z0-9_]/g, '_x005f_')}ListItem`;

      // Build the item payload based on list type
      let payload: any = {
        '__metadata': { 'type': entityType },
        'Title': itemData.title || itemData.ideaType || 'New Idea'
      };

      // Add common fields for all list types
      if (itemData.meetingDate) {
        // If the date is already in ISO format, use it directly
        // Otherwise, ensure the date string is interpreted as UTC to avoid timezone shifts
        let meetingDate;
        if (itemData.meetingDate.includes('T')) {
          // Already has time component
          meetingDate = new Date(itemData.meetingDate);
        } else {
          // Just a date string like "2025-07-15", add time to ensure correct date
          meetingDate = new Date(itemData.meetingDate + 'T10:00:00.000Z');
        }
        payload['MeetingDate'] = meetingDate.toISOString();

      }
      
      // Add fields specific to each list type
      if (listType === 'Business Ideas') {
        payload['BusinessIdea1'] = itemData.description;
        payload['Status'] = itemData.status || 'Submitted';
        payload['Idea_x0020_Type'] = itemData.ideaType || 'Process Improvement';
        payload['Idea_x0020_Date'] = new Date().toISOString();
      } else if (listType === 'Safety Ideas') {
        payload['SafetyIdea1'] = itemData.description;
        payload['Status'] = itemData.status || 'Submitted';
        payload['Idea_x0020_Type'] = itemData.ideaType || 'Safety Improvement';
        payload['Idea_x0020_Date'] = new Date().toISOString();
      } else if (listType === 'Near Miss') {
        payload['Date'] = new Date().toISOString();
        payload['EventType'] = itemData.eventType || 'Near Miss';
        payload['Canyoubrieflyexplainwhathappened'] = itemData.whatHappened || itemData.description;
        payload['Howdidthishappen_x003f_'] = itemData.howItHappened || '';
        payload['Status'] = itemData.status || 'Submitted';
      }



      const createResponse = await fetch(createUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        body: JSON.stringify(payload)
      });
      
      if (!createResponse.ok) {
        const errorText = await createResponse.text();
        throw new Error(`SharePoint create failed: ${createResponse.status} - ${errorText}`);
      }
      
      const responseData = await createResponse.json();
      const newItemId = responseData.d.ID;
      
      return `${listType.toLowerCase().replace(' ', '-')}-${newItemId}`;
      
    } catch (error) {
      console.error(` Failed to create ${listType} item:`, error);
      throw error;
    }
  }

  /**
   * Update item status and other fields using REST API
   */
  async updateItemFields(itemId: string, listType: string, updates: {
    status?: string;
    meetingDate?: string;
    meetingNotes?: string;
    title?: string;
    description?: string;
    ideaType?: string;
    eventType?: string;
    // Action tracking fields
    actionPriority?: string;
    actionStatus?: string;
    actionAssignedTo?: string;
    actionStartDate?: string;
    actionDueDate?: string;
    actionNotes?: string;
  }): Promise<void> {
    const config = SharePointListsService.LIST_CONFIGS[listType];
    if (!config) {
      throw new Error(`List configuration not found for: ${listType}`);
    }

    // Extract numeric ID from formatted ID
    const numericId = itemId.replace(`${listType.toLowerCase().replace(' ', '-')}-`, '');
    
    // Use custom site URL if specified, otherwise use main site
    const baseUrl = config.siteUrl 
      ? `${config.siteUrl}/_api/web/lists/getbytitle`
      : this.baseUrl;
    
    const updateUrl = `${baseUrl}('${config.listTitle}')/items(${numericId})`;
    
    try {
      // Get item etag for optimistic concurrency
      const getResponse = await fetch(updateUrl, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (!getResponse.ok) {
        throw new Error(`Failed to get item for update: ${getResponse.status}`);
      }
      
      const itemData = await getResponse.json();
      const etag = itemData.d.__metadata.etag;
      const entityType = itemData.d.__metadata.type;
      
      // Build update payload based on list type and updates
      let payload: any = {
        '__metadata': { 'type': entityType }
      };

      // Add common fields
      if (updates.title) payload['Title'] = updates.title;
      if (updates.status) payload['Status'] = updates.status;
      if (updates.meetingDate) payload['MeetingDate'] = new Date(updates.meetingDate).toISOString();

      // Add list-specific fields
      if (listType === 'Business Ideas') {
        if (updates.description) payload['BusinessIdea1'] = updates.description;
        if (updates.meetingNotes !== undefined) payload['MeetingNotes'] = updates.meetingNotes;
        if (updates.ideaType) payload['Idea_x0020_Type'] = updates.ideaType;
      } else if (listType === 'Safety Ideas') {
        if (updates.description) payload['SafetyIdea1'] = updates.description;
        if (updates.meetingNotes !== undefined) payload['Meeting_x0020_Notes1'] = updates.meetingNotes;
        if (updates.ideaType) payload['Idea_x0020_Type'] = updates.ideaType;
      } else if (listType === 'Near Miss') {
        if (updates.description) payload['Canyoubrieflyexplainwhathappened'] = updates.description;
        if (updates.meetingNotes !== undefined) payload['MeetingNotes'] = updates.meetingNotes;
        if (updates.eventType) payload['EventType'] = updates.eventType;
      }

      // Add action tracking fields (common across all list types)
      // Note: These field names may need to be adjusted based on actual SharePoint column names
      if (updates.actionPriority !== undefined) payload['ActionPriority'] = updates.actionPriority;
      if (updates.actionStatus !== undefined) payload['ActionStatus'] = updates.actionStatus;
      if (updates.actionAssignedTo !== undefined) payload['ActionAssignedTo'] = updates.actionAssignedTo;
      if (updates.actionStartDate !== undefined) payload['ActionStartDate'] = updates.actionStartDate ? new Date(updates.actionStartDate).toISOString() : null;
      if (updates.actionDueDate !== undefined) payload['ActionDueDate'] = updates.actionDueDate ? new Date(updates.actionDueDate).toISOString() : null;
      if (updates.actionNotes !== undefined) payload['ActionNotes'] = updates.actionNotes;

      // Update the item
      const updateResponse = await fetch(updateUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'If-Match': etag,
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(payload)
      });
      
      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        console.error(` SharePoint update failed: ${updateResponse.status} - ${errorText}`);
        throw new Error(`SharePoint update failed: ${updateResponse.status} - ${errorText}`);
      }
      
      
    } catch (error) {
      logger.error({ err: error, listType, itemId: numericId }, 'Failed to update SharePoint item');
      throw error;
    }
  }

  /**
   * Change who an existing item was submitted by. The "submitted by" column is a
   * SharePoint Person field, so the chosen person has to be resolved to a
   * site-specific user id (via ensureuser) before the field can be set. This is
   * done on the item's own site, which matters for Near Miss because it lives on
   * a separate site collection with its own user ids.
   */
  async updateItemSubmitter(itemId: string, listType: string, userLoginName: string): Promise<void> {
    const allowed = ['Business Ideas', 'Safety Ideas', 'Near Miss'];
    if (!allowed.includes(listType)) {
      throw new Error('Submitter can only be changed on Business Ideas, Safety Ideas and Near Miss items');
    }

    const config = SharePointListsService.LIST_CONFIGS[listType];
    if (!config) {
      throw new Error(`List configuration not found for: ${listType}`);
    }
    if (!userLoginName) {
      throw new Error('A person must be selected to change the submitter');
    }

    const personFieldId = `${config.submittedByField}Id`; // e.g. "NameId"
    const webBase = config.siteUrl
      ? `${config.siteUrl}/_api/web`
      : 'https://cranfieldglass.sharepoint.com/_api/web';
    const baseUrl = config.siteUrl
      ? `${config.siteUrl}/_api/web/lists/getbytitle`
      : this.baseUrl;

    // Resolve the person to a user id on the item's own site collection.
    const ensureResponse = await fetch(`${webBase}/ensureuser`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${this.accessToken}`,
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      },
      body: JSON.stringify({ logonName: userLoginName })
    });

    if (!ensureResponse.ok) {
      const errorText = await ensureResponse.text();
      throw new Error(`Failed to resolve the selected person on SharePoint: ${ensureResponse.status} - ${errorText}`);
    }

    const ensureData = await ensureResponse.json();
    const userId = ensureData.d?.Id;
    if (!userId) {
      throw new Error('Could not resolve the selected person to a SharePoint user');
    }

    // Extract numeric ID from formatted ID (e.g. "near-miss-42" -> "42")
    const numericId = itemId.replace(`${listType.toLowerCase().replace(' ', '-')}-`, '');
    const updateUrl = `${baseUrl}('${config.listTitle}')/items(${numericId})`;

    try {
      // Get item etag + entity type for the MERGE update
      const getResponse = await fetch(updateUrl, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose'
        }
      });

      if (!getResponse.ok) {
        throw new Error(`Failed to get item for update: ${getResponse.status}`);
      }

      const itemData = await getResponse.json();
      const etag = itemData.d.__metadata.etag;
      const entityType = itemData.d.__metadata.type;

      const payload: any = {
        '__metadata': { 'type': entityType },
        [personFieldId]: userId
      };

      const updateResponse = await fetch(updateUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'If-Match': etag,
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(payload)
      });

      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        throw new Error(`SharePoint submitter update failed: ${updateResponse.status} - ${errorText}`);
      }
    } catch (error) {
      logger.error({ err: error, listType, itemId: numericId }, 'Failed to update SharePoint item submitter');
      throw error;
    }
  }

  /**
   * Move an item from one list to another (e.g. a Safety idea that should be
   * a Business idea). This copies the item's content into the destination list
   * preserving the title, description, status, dates and meeting notes, then
   * deletes the original. The original is only deleted after the copy succeeds,
   * so a failure never loses data.
   */
  async moveItem(itemId: string, fromList: string, toList: string): Promise<string> {
    const movableLists = ['Business Ideas', 'Safety Ideas', 'Near Miss'];
    if (!movableLists.includes(fromList) || !movableLists.includes(toList)) {
      throw new Error('Items can only be moved between Business Ideas, Safety Ideas and Near Miss lists');
    }
    if (fromList === toList) {
      throw new Error('The item is already on that list');
    }

    const fromConfig = SharePointListsService.LIST_CONFIGS[fromList];
    const toConfig = SharePointListsService.LIST_CONFIGS[toList];
    if (!fromConfig || !toConfig) {
      throw new Error('List configuration not found');
    }

    const fromBaseUrl = fromConfig.siteUrl
      ? `${fromConfig.siteUrl}/_api/web/lists/getbytitle`
      : this.baseUrl;
    const numericId = itemId.replace(`${fromList.toLowerCase().replace(' ', '-')}-`, '');
    const sourceUrl = `${fromBaseUrl}('${fromConfig.listTitle}')/items(${numericId})`;

    try {
      // 1. Read the raw source item (and its etag for the later delete)
      const getResponse = await fetch(sourceUrl, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose'
        }
      });
      if (!getResponse.ok) {
        throw new Error(`Failed to read source item: ${getResponse.status}`);
      }
      const sourceData = (await getResponse.json()).d;
      const etag = sourceData.__metadata.etag;

      // Standardise the fields we want to carry across
      const title = sourceData.Title || '';
      const description = sourceData[fromConfig.contentField] || '';
      const status = sourceData.Status || 'Submitted';
      const meetingNotes = sourceData[fromConfig.meetingNotesField] || '';
      const meetingDate = sourceData.MeetingDate || null;
      const submissionDate = sourceData[fromConfig.submissionDateField] || sourceData.Created;
      const secondaryDescription = sourceData['Howdidthishappen_x003f_'] || '';

      // 2. Create the item in the destination list
      const toBaseUrl = toConfig.siteUrl
        ? `${toConfig.siteUrl}/_api/web/lists/getbytitle`
        : this.baseUrl;
      const createUrl = `${toBaseUrl}('${toConfig.listTitle}')/items`;

      const listMetaResponse = await fetch(`${toBaseUrl}('${toConfig.listTitle}')?$select=ListItemEntityTypeFullName`, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose'
        }
      });
      const listMeta = await listMetaResponse.json();
      const entityType = listMeta.d?.ListItemEntityTypeFullName
        || `SP.Data.${toConfig.listTitle.replace(/\s/g, '_x0020_').replace(/[^a-zA-Z0-9_]/g, '_x005f_')}ListItem`;

      const isoOrNow = (value: any) => value ? new Date(value).toISOString() : new Date().toISOString();

      let payload: any = {
        '__metadata': { 'type': entityType },
        'Title': title,
        'Status': status
      };
      if (meetingDate) payload['MeetingDate'] = new Date(meetingDate).toISOString();

      if (toList === 'Business Ideas') {
        payload['BusinessIdea1'] = description;
        payload['Idea_x0020_Type'] = 'Process Improvement';
        payload['Idea_x0020_Date'] = isoOrNow(submissionDate);
        payload['MeetingNotes'] = meetingNotes;
      } else if (toList === 'Safety Ideas') {
        payload['SafetyIdea1'] = description;
        payload['Idea_x0020_Type'] = 'Safety Improvement';
        payload['Idea_x0020_Date'] = isoOrNow(submissionDate);
        payload['Meeting_x0020_Notes1'] = meetingNotes;
      } else if (toList === 'Near Miss') {
        payload['Canyoubrieflyexplainwhathappened'] = description;
        payload['Howdidthishappen_x003f_'] = secondaryDescription;
        payload['EventType'] = 'Near Miss';
        payload['Date'] = isoOrNow(submissionDate);
        payload['MeetingNotes'] = meetingNotes;
      }

      const createResponse = await fetch(createUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        body: JSON.stringify(payload)
      });
      if (!createResponse.ok) {
        const errorText = await createResponse.text();
        // Nothing was deleted yet, so the original item is safe.
        throw new Error(`Failed to copy item to ${toList}: ${createResponse.status} - ${errorText}`);
      }
      const newItemId = (await createResponse.json()).d.ID;

      // 3. Delete the original item now that the copy succeeded
      const deleteResponse = await fetch(sourceUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose',
          'If-Match': etag,
          'X-HTTP-Method': 'DELETE'
        }
      });
      if (!deleteResponse.ok) {
        const errorText = await deleteResponse.text();
        throw new Error(`Item was copied to ${toList} but the original could not be removed from ${fromList}. Please delete it in SharePoint. (${deleteResponse.status} - ${errorText})`);
      }

      return `${toList.toLowerCase().replace(' ', '-')}-${newItemId}`;
    } catch (error) {
      logger.error({ err: error, fromList, toList, itemId }, 'Failed to move SharePoint item');
      throw error;
    }
  }

  /**
   * Get SharePoint Choice field options for status dropdowns
   */
  async getChoiceFieldOptions(listType: string, fieldName: string): Promise<string[]> {
    const config = SharePointListsService.LIST_CONFIGS[listType];
    if (!config) {
      throw new Error(`List configuration not found for: ${listType}`);
    }

    // Use custom site URL if specified, otherwise use main site
    const baseUrl = config.siteUrl 
      ? `${config.siteUrl}/_api/web/lists/getbytitle`
      : this.baseUrl;
    
    const fieldUrl = `${baseUrl}('${config.listTitle}')/fields/getbytitle('${fieldName}')`;
    
    try {
      const response = await fetch(fieldUrl, {
        headers: {
          'Authorization': `Bearer ${this.accessToken}`,
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (!response.ok) {
        console.warn(`Could not fetch choice options for ${fieldName} in ${listType}`);
        // Return default options based on what we know exists
        if (fieldName.toLowerCase().includes('status')) {
          return ['Submitted', 'In Discussion', 'Actions', 'Closed'];
        }
        return [];
      }
      
      const fieldData = await response.json();
      const choices = fieldData.d.Choices?.results || [];
      
      return choices;
      
    } catch (error) {
      console.error(` Failed to get choice options for ${fieldName}:`, error);
      // Return sensible defaults
      if (fieldName.toLowerCase().includes('status')) {
        return ['Submitted', 'In Discussion', 'Actions', 'Closed'];
      }
      return [];
    }
  }

  /**
   * Check if service is properly configured
   */
  isConfigured(): boolean {
    return !!this.accessToken;
  }

  /**
   * Get list of pending field integrations that need internal names
   */
  static getPendingFieldIntegrations(): { listType: string; fields: string[]; instructions: string }[] {
    return [
      {
        listType: 'Near Miss',
        fields: ['Meeting Notes', 'Report Link'],
        instructions: 'Add sample data to these fields in SharePoint, then re-run Power Automate JSON extraction to capture internal field names'
      }
    ];
  }
}