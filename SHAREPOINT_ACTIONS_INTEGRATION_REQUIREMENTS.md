# SharePoint Actions Integration Requirements

## Overview
This document outlines the SharePoint list modifications needed to create a seamless Actions workflow where items progress from Ideas → Actions → Completion with automatic status synchronization.

## Current Workflow vs Required Workflow

### Current State
- Ideas are marked "Actions" in the original list
- No automatic creation of Action items
- No completion tracking or status synchronization

### Required Workflow
1. **Idea Submission** → Ideas marked as "Submitted" in respective lists
2. **Meeting Discussion** → Ideas marked as "In Discussion" during meetings
3. **Action Creation** → Ideas marked as "Actioned" + Action item created automatically
4. **Action Completion** → Action marked as "Finished" + Original idea marked as "Closed"

## SharePoint List Modifications Required

### 1. Business Ideas List
**List Name:** `Business Ideas`
**SharePoint URL:** `/lists/Business Ideas`

#### Required Column Additions:
```
ActionItemId (Number)
- Purpose: Links to corresponding Action item ID
- Required: No
- Default: Blank

ActionCreatedDate (Date and Time) 
- Purpose: When the Action was created
- Required: No
- Default: Blank

ActionAssignedTo (Person or Group)
- Purpose: Who is assigned to complete the action
- Required: No  
- Default: Blank

ActionDueDate (Date and Time)
- Purpose: When the action is due for completion
- Required: No
- Default: Blank

ActionStatus (Choice)
- Purpose: Track action progress independently  
- Required: No
- Choices: ["Not Started", "In Progress", "Completed", "On Hold"]
- Default: Blank

ActionNotes (Multiple lines of text)
- Purpose: Progress updates and notes on action completion
- Required: No
- Default: Blank
```

#### Required Status Choice Updates:
**Existing Column:** `Status`
**Current Choices:** ["Submitted", "In Discussion", "Actions", "Closed"]
**Required Changes:**
- Change "Actions" to "Actioned" 
- Keep "Closed" (set automatically when Action is Finished)

**Final Status Choices:** `["Submitted", "In Discussion", "Actioned", "Closed"]`

### 2. Safety Ideas List  
**List Name:** `Safety Ideas`
**SharePoint URL:** `/lists/Safety Ideas`

#### Required Column Additions:
```
ActionItemId (Number)
- Purpose: Links to corresponding Action item ID
- Required: No
- Default: Blank

ActionCreatedDate (Date and Time)
- Purpose: When the Action was created  
- Required: No
- Default: Blank

ActionAssignedTo (Person or Group)
- Purpose: Who is assigned to complete the action
- Required: No
- Default: Blank

ActionDueDate (Date and Time)
- Purpose: When the action is due for completion
- Required: No
- Default: Blank

ActionStatus (Choice)
- Purpose: Track action progress independently
- Required: No  
- Choices: ["Not Started", "In Progress", "Completed", "On Hold"]
- Default: Blank

ActionNotes (Multiple lines of text)
- Purpose: Progress updates and notes on action completion
- Required: No
- Default: Blank
```

#### Required Status Choice Updates:
**Existing Column:** `Status`
**Current Choices:** ["Submitted", "In Discussion", "Actions", "Closed"]
**Required Changes:**
- Change "Actions" to "Actioned"
- Keep "Closed" (set automatically when Action is Finished)

**Final Status Choices:** `["Submitted", "In Discussion", "Actioned", "Closed"]`

### 3. Near Miss - Accident Safety Register List
**List Name:** `Near Miss - Accident Safety Register`  
**SharePoint URL:** `/sites/IncidentsReports/lists/Near Miss - Accident Safety Register`

#### Required Column Additions:
```
ActionItemId (Number)
- Purpose: Links to corresponding Action item ID
- Required: No
- Default: Blank

ActionCreatedDate (Date and Time)
- Purpose: When the Action was created
- Required: No  
- Default: Blank

ActionAssignedTo (Person or Group)
- Purpose: Who is assigned to complete the action
- Required: No
- Default: Blank

ActionDueDate (Date and Time)
- Purpose: When the action is due for completion
- Required: No
- Default: Blank

ActionStatus (Choice)
- Purpose: Track action progress independently
- Required: No
- Choices: ["Not Started", "In Progress", "Completed", "On Hold"]  
- Default: Blank

ActionNotes (Multiple lines of text)
- Purpose: Progress updates and notes on action completion
- Required: No
- Default: Blank
```

#### Required Status Choice Updates:
**Existing Column:** `Status`
**Current Choices:** ["Submitted", "In Discussion", "Actions", "Closed"]
**Required Changes:**
- Change "Actions" to "Actioned"
- Add "Closed" (set automatically when Action is Finished)

**Final Status Choices:** `["Submitted", "In Discussion", "Actioned", "Closed"]`

### 4. Actions List (Enhanced)
**List Name:** `Actions`
**SharePoint URL:** `/lists/Actions`

#### Required Column Additions:
```
ActionCompletionDate (Date and Time)
- Purpose: When the action was actually completed
- Required: No
- Default: Blank

OutcomeResult (Multiple lines of text)  
- Purpose: Final outcome and results description
- Required: No
- Default: Blank

OriginalListType (Choice)
- Purpose: Which list the original item came from
- Required: Yes
- Choices: ["Business Ideas", "Safety Ideas", "Near Miss"]
- Default: Blank

OriginalItemId (Number)
- Purpose: ID of the original item in source list  
- Required: Yes
- Default: Blank

AutoCreated (Yes/No)
- Purpose: Flag indicating if created automatically vs manually
- Required: No
- Default: Yes
```

#### Required Status Choice Updates:
**Existing Column:** `Status`  
**Current Choices:** ["Not Started", "In Progress", "Completed"]
**Required Changes:**
- Add "Finished" status (triggers original item to "Closed")
- Keep existing statuses for internal action tracking

**Final Status Choices:** `["Not Started", "In Progress", "Completed", "Finished"]`

## Power Automate Flows Required

### Flow 1: Auto-Create Actions
**Trigger:** When item status changes to "Actioned" in any source list
**Actions:**
1. Create new Action item with data from source item
2. Set OriginalListType and OriginalItemId
3. Set ActionItemId in source item  
4. Set ActionCreatedDate in source item
5. Copy ActionAssignedTo if set in source item

### Flow 2: Auto-Close Original Items  
**Trigger:** When Action status changes to "Finished"
**Actions:**
1. Lookup original item using OriginalListType and OriginalItemId
2. Set original item status to "Closed"
3. Set ActionCompletionDate in Action item
4. Update source item's ActionStatus to "Completed"

## API Integration Changes

### New Backend Endpoints Required:
```
POST /api/actions/create-from-item
- Creates Action item from source item
- Links Action back to source item

PUT /api/actions/:id/finish  
- Marks Action as Finished
- Triggers original item closure

GET /api/actions/by-source/:listType/:itemId
- Gets Actions for specific source item
```

### Modified Existing Endpoints:
```
PUT /api/sharepoint/items/:listType/:id
- Enhanced to handle embedded Action fields
- Updates ActionStatus when item status changes

GET /api/meeting-history  
- Returns embedded Action data with items
- Shows Action progress without separate API calls
```

## UI Changes Required

### Status Dropdown Logic:
- **Regular Items:** Show `["Submitted", "In Discussion", "Actioned"]`
- **Actions Tab:** Show `["Not Started", "In Progress", "Completed", "Finished"]` 
- **Closed Items:** Read-only, no status changes allowed

### Item Details Enhancement:
- Show embedded Action fields when item status is "Actioned"
- Display Action progress without separate Actions tab
- Allow Action status updates directly in item details

### Action Creation Flow:
1. User sets item status to "Actioned"  
2. System prompts for Action assignment and due date
3. Action created automatically with proper linking
4. Item shows embedded Action tracking fields

## Data Consistency Rules

### Validation Rules:
1. Items can only be marked "Closed" by system (never manually)
2. ActionItemId must exist when status is "Actioned" 
3. ActionStatus must be "Completed" when parent Action is "Finished"
4. ActionDueDate required when creating Actions

### Synchronization Rules:
1. Action "Finished" → Original item "Closed" (automatic)
2. Action assigned person → ActionAssignedTo in source item (automatic)
3. Action due date → ActionDueDate in source item (automatic)
4. Action progress notes → ActionNotes in source item (automatic)

This integration creates a seamless workflow where Actions and Ideas stay perfectly synchronized without requiring separate list management.