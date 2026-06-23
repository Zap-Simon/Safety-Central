/**
 * Markdown-based PDF Export System
 * 
 * Generates structured markdown documents with proper heading hierarchy
 * for PDF export with navigation bookmarks (similar to Typora approach)
 */

import { buildActionRequiredLines, getDisplayItemStatus, type ReadyToCloseAction } from "./meeting-export-shared";

interface MeetingItem {
  id: string;
  title: string;
  description: string;
  type: 'Business Ideas' | 'Safety Ideas' | 'Near Miss' | 'Actions';
  status: string;
  meetingDate: string;
  meetingNotes: string;
  submittedBy: string;
  submittedDate: string;
  assignedTo?: string;
  ideaType?: string;
  secondaryDescription?: string;
  // Action-related properties
  actionAssignedTo?: string;
  actionStatus?: string;
  actionPriority?: string;
  actionStartDate?: string;
  actionDueDate?: string;
  actionNotes?: string;
}

interface MeetingAnalytics {
  totalItems: number;
  statusBreakdown: Record<string, number>;
  typeBreakdown: Record<string, number>;
  submissionsByPerson: Record<string, number>;
}

export class MarkdownMeetingGenerator {
  
  /**
   * Generate structured markdown document with proper heading hierarchy
   * This creates PDF bookmarks when converted: # = Level 1, ## = Level 2, etc.
   */
  static generateMeetingMarkdown(
    filteredData: MeetingItem[], 
    meetingDate: string, 
    currentDate: string,
    meetingAttendance?: Record<string, string[]>,
    selectedMeeting?: string,
    readyToCloseActions: ReadyToCloseAction[] = []
  ): string {
    
    const analytics = this.generateAnalytics(filteredData);
    
    let markdown = `# CRANFIELD GLASS CHRISTCHURCH
## Health & Safety Committee Meeting Minutes

**Meeting Date:** ${meetingDate}  
**Document Generated:** ${currentDate}  
**Total Items:** ${analytics.totalItems}

---

## Table of Contents
- [Meeting Overview](#meeting-overview)
- [Attendance Record](#attendance-record)  
- [Meeting Analytics](#meeting-analytics)
- [Agenda Items](#agenda-items)
  - [Near Miss Reports](#near-miss-reports)
  - [Safety Ideas](#safety-ideas)
  - [Business Ideas](#business-ideas)
  - [Action Items](#action-items)
- [Actions Ready to Close](#actions-ready-to-close)
- [Meeting Approval](#meeting-approval)

---

## Meeting Overview

This document contains the complete record of items discussed during the Health & Safety Committee meeting held on ${meetingDate}. All submissions have been reviewed and categorized according to company policy.

### Meeting Statistics
- **Total Submissions:** ${analytics.totalItems}
- **Active Items:** ${analytics.statusBreakdown['Submitted'] || 0}
- **In Discussion:** ${analytics.statusBreakdown['In Discussion'] || 0}
- **Actioned Items:** ${analytics.statusBreakdown['Actions'] || 0}
- **Closed Items:** ${analytics.statusBreakdown['Closed'] || 0}

---

## Attendance Record

${this.generateAttendanceSection(meetingAttendance, selectedMeeting, meetingDate)}

---

## Meeting Analytics

### Submission Breakdown by Type
${Object.entries(analytics.typeBreakdown).map(([type, count]) => 
  `- **${type}:** ${count} items`
).join('\n')}

### Status Distribution
${Object.entries(analytics.statusBreakdown).map(([status, count]) => 
  `- **${status}:** ${count} items`
).join('\n')}

### Top Contributors
${Object.entries(analytics.submissionsByPerson)
  .sort(([,a], [,b]) => b - a)
  .slice(0, 5)
  .map(([person, count]) => `- **${person}:** ${count} submissions`)
  .join('\n')}

---

## Agenda Items

The following items were presented for discussion and review during this meeting. Items are organized by type and priority for systematic review.

`;

    // Group items by type with proper heading hierarchy
    const groupedItems = this.groupItemsByType(filteredData);
    
    // Near Miss Reports (Highest Priority)
    if (groupedItems['Near Miss']?.length > 0) {
      markdown += this.generateTypeSection('Near Miss', groupedItems['Near Miss'], '### 🚨 Near Miss Reports\n\n*Critical safety incidents requiring immediate attention and follow-up*\n\n');
    }
    
    // Safety Ideas
    if (groupedItems['Safety Ideas']?.length > 0) {
      markdown += this.generateTypeSection('Safety Ideas', groupedItems['Safety Ideas'], '### 🔒 Safety Ideas\n\n*Staff suggestions for improving workplace safety and procedures*\n\n');
    }
    
    // Business Ideas  
    if (groupedItems['Business Ideas']?.length > 0) {
      markdown += this.generateTypeSection('Business Ideas', groupedItems['Business Ideas'], '### 💡 Business Ideas\n\n*Staff suggestions for business improvements and operational efficiency*\n\n');
    }
    
    // Action Items
    if (groupedItems['Actions']?.length > 0) {
      markdown += this.generateTypeSection('Actions', groupedItems['Actions'], '### ✅ Action Items\n\n*Follow-up actions and implementation tracking*\n\n');
    }

    // Actions Ready to Close — drawn from the whole backlog, not just this meeting
    markdown += this.generateReadyToCloseSection(readyToCloseActions);

    // Meeting Approval Section
    markdown += `

---

## Meeting Approval

### Committee Signatures

**Health & Safety Coordinator**  
Name: Simon Hubbard  
Signature: _________________________  
Date: _________________________

**Health & Safety Manager**  
Name: _________________________  
Signature: _________________________  
Date: _________________________

**General Manager**  
Name: Kevin Young  
Signature: _________________________  
Date: _________________________

### Next Meeting

**Scheduled Date:** ${this.calculateNextMeetingDate(meetingDate)}  
**Location:** Cranfield Glass Offices  
**Agenda Items:** Review of action items and new submissions

---

*This document was automatically generated from the Cranfield Glass Health & Safety Management System on ${currentDate}.*

`;

    return markdown;
  }

  /**
   * Generate type-specific sections with consistent formatting
   */
  private static generateTypeSection(type: string, items: MeetingItem[], headerMarkdown: string): string {
    let section = headerMarkdown;
    
    section += `**${items.length} ${type.toLowerCase()} item${items.length !== 1 ? 's' : ''} for review**\n\n`;
    
    items.forEach((item, index) => {
      const itemNumber = index + 1;
      
      // Each item gets its own level 4 heading for PDF bookmarks
      section += `#### ${itemNumber}. ${item.title || `${type} Submission`}\n\n`;
      
      // Item details in structured format
      section += `**Submitted by:** ${item.submittedBy}  \n`;
      section += `**Date:** ${new Date(item.submittedDate).toLocaleDateString('en-GB')}  \n`;
      section += `**Status:** ${getDisplayItemStatus(item)}  \n`;
      if (item.assignedTo) {
        section += `**Assigned to:** ${item.assignedTo}  \n`;
      }
      if (item.ideaType) {
        section += `**Category:** ${item.ideaType}  \n`;
      }
      section += '\n';
      
      // Main content
      if (type === 'Near Miss' && item.secondaryDescription) {
        section += `**What Happened:**\n${item.description}\n\n`;
        section += `**How It Happened:**\n${item.secondaryDescription}\n\n`;
      } else {
        section += `**Description:**\n${item.description}\n\n`;
      }
      
      // Meeting notes if available
      if (item.meetingNotes && item.meetingNotes.trim()) {
        section += `**Meeting Discussion:**\n${item.meetingNotes}\n\n`;
      }
      
      // Action required section
      section += `**Action Required:**\n`;
      section += this.generateActionText(item) + '\n\n';
      
      section += '---\n\n';
    });
    
    return section;
  }

  /**
   * Group items by type with priority sorting
   */
  private static groupItemsByType(items: MeetingItem[]): Record<string, MeetingItem[]> {
    const grouped: Record<string, MeetingItem[]> = {};
    
    items.forEach(item => {
      if (!grouped[item.type]) {
        grouped[item.type] = [];
      }
      grouped[item.type].push(item);
    });
    
    // Sort within each type by submission date (newest first)
    Object.keys(grouped).forEach(type => {
      grouped[type].sort((a, b) => 
        new Date(b.submittedDate).getTime() - new Date(a.submittedDate).getTime()
      );
    });
    
    return grouped;
  }

  /**
   * Generate analytics from meeting data
   */
  private static generateAnalytics(items: MeetingItem[]): MeetingAnalytics {
    const statusBreakdown: Record<string, number> = {};
    const typeBreakdown: Record<string, number> = {};
    const submissionsByPerson: Record<string, number> = {};
    
    items.forEach(item => {
      // Status breakdown
      statusBreakdown[item.status] = (statusBreakdown[item.status] || 0) + 1;
      
      // Type breakdown
      typeBreakdown[item.type] = (typeBreakdown[item.type] || 0) + 1;
      
      // Submissions by person
      submissionsByPerson[item.submittedBy] = (submissionsByPerson[item.submittedBy] || 0) + 1;
    });
    
    return {
      totalItems: items.length,
      statusBreakdown,
      typeBreakdown,
      submissionsByPerson
    };
  }

  /**
   * Generate attendance section based on UI selections - synced with HTML export
   */
  private static generateAttendanceSection(
    meetingAttendance?: Record<string, string[]>, 
    selectedMeeting?: string,
    meetingDate?: string
  ): string {
    // Use the same attendee structure as HTML export for consistency
    const allAttendees = {
      management: [
        { name: 'Hoani Hunt', role: 'Company Director' },
        { name: 'Simon Hubbard', role: 'Health & Safety Coordinator' },
        { name: 'James Waites', role: 'Glazing Supervisor' },
        { name: 'Emma White', role: 'Administrator' }
      ],
      glaziers: [
        { name: 'Kevin Young', role: 'Glazier' },
        { name: 'Ryan Newman', role: 'Glazier' },
        { name: 'Isaac Ensor', role: 'Glazier' },
        { name: 'Struan O\'Donnell', role: 'Glazier' },
        { name: 'Sam Chang', role: 'Glazier' }
      ]
    };

    // Function to check if attendee is present (matches HTML logic)
    const isAttending = (attendeeName: string): boolean => {
      if (!meetingAttendance || !selectedMeeting || selectedMeeting === 'all') {
        return true; // Default to present if no specific meeting data
      }
      
      const attendeesForMeeting = meetingAttendance[selectedMeeting];
      if (!attendeesForMeeting) {
        return true; // Default to present if no attendance data
      }
      
      return attendeesForMeeting.includes(attendeeName);
    };

    if (!meetingAttendance || !selectedMeeting || selectedMeeting === 'all') {
      return `**Attendance will be recorded during the meeting**

### Management Team
| Name | Position | Present | Absent |
|------|----------|---------|--------|
${allAttendees.management.map(attendee => 
  `| ${attendee.name} | ${attendee.role} | ☐ | ☐ |`
).join('\n')}

### Glaziers
| Name | Position | Present | Absent |
|------|----------|---------|--------|
${allAttendees.glaziers.map(attendee => 
  `| ${attendee.name} | ${attendee.role} | ☐ | ☐ |`
).join('\n')}

`;
    }

    // Generate attendance with actual UI data
    let attendanceMarkdown = `**Meeting Attendance Record**

### Management Team
| Name | Position | Present | Absent |
|------|----------|---------|--------|
${allAttendees.management.map(attendee => {
  const isPresent = isAttending(attendee.name);
  return `| ${attendee.name} | ${attendee.role} | ${isPresent ? '☑' : '☐'} | ${isPresent ? '☐' : '☑'} |`;
}).join('\n')}

### Glaziers
| Name | Position | Present | Absent |
|------|----------|---------|--------|
${allAttendees.glaziers.map(attendee => {
  const isPresent = isAttending(attendee.name);
  return `| ${attendee.name} | ${attendee.role} | ${isPresent ? '☑' : '☐'} | ${isPresent ? '☐' : '☑'} |`;
}).join('\n')}

`;

    return attendanceMarkdown;
  }

  /**
   * Generate action text based on item type and status - improved to avoid generic text
   */
  private static generateActionText(item: MeetingItem): string {
    // Sourced strictly from the real actioned system, shared with every other
    // export format. No fabricated type/status boilerplate.
    return buildActionRequiredLines(item)
      .map(line => (line.label ? `**${line.label}:** ${line.value}` : line.value))
      .join('  \n');
  }

  /**
   * Actions parked at "Ready to Close" across the whole backlog (not just this
   * meeting). They need a group review + sign-off to be formally closed, so the
   * same action legitimately re-appears in consecutive minutes — the due date is
   * shown with each so it can be tracked even when it sits in the future.
   */
  private static generateReadyToCloseSection(readyToCloseActions: ReadyToCloseAction[]): string {
    let section = `\n\n---\n\n## Actions Ready to Close\n\n`;

    if (!readyToCloseActions || readyToCloseActions.length === 0) {
      section += `*No actions are currently ready to close.*\n\n`;
      return section;
    }

    section += `| Item | Type | Actioned By | Due Date | What Was Done |\n`;
    section += `| --- | --- | --- | --- | --- |\n`;
    readyToCloseActions.forEach((action) => {
      const cell = (v: string) => (v || '—').replace(/\|/g, '\\|').replace(/\n+/g, ' ');
      section += `| ${cell(action.title)} | ${cell(action.type)} | ${cell(action.assignedTo)} | ${cell(action.dueDate)} | ${cell(action.outcome)} |\n`;
    });
    section += `\n`;

    return section;
  }

  /**
   * Calculate next meeting date (fortnightly Tuesday)
   */
  private static calculateNextMeetingDate(meetingDate: string): string {
    if (meetingDate === 'All Meetings') {
      return 'TBD';
    }
    
    const date = new Date(meetingDate);
    date.setDate(date.getDate() + 14); // Add 2 weeks
    return date.toLocaleDateString('en-GB');
  }
}