# SharePoint List API Integration Instructions (Near Miss - Accident Safety Register)

## 🔄 PENDING FIELD INTEGRATION

**Missing Fields (Need Internal Names):**
- **Meeting Notes**: Likely `Meeting_x0020_Notes1`, `MeetingNotes`, or similar pattern
- **Report Link**: Likely `{Link}`, `ReportLink`, or URL field pattern

**Action Required:**
1. Add sample data to these fields in SharePoint
2. Re-run Power Automate JSON extraction to capture internal field names
3. Update `selectFields` in `sharepoint-lists-service.ts` Near Miss configuration
4. Update `meetingNotesField` configuration once identified

**Quick Integration Checklist (when field names are found):**
```javascript
// In sharepoint-lists-service.ts, Near Miss selectFields array, add:
'ACTUAL_MEETING_NOTES_FIELD_NAME',
'ACTUAL_REPORT_LINK_FIELD_NAME',

// Update meetingNotesField:
meetingNotesField: 'ACTUAL_MEETING_NOTES_FIELD_NAME',
```

**Current Status:** Fields exist but are empty, so internal names not captured in schema

## 1. How to Find Internal Names for Your Flow

Look for the internal field names in your schema (`NearMissAccidentSafetyRegister-Schema-Tidied.json`).

**Key fields:**
- **Title**: `Title`
- **Date**: `Date`
- **Event Type**: `EventType` (Choice field)
- **Brief Explanation**: `Canyoubrieflyexplainwhathappened`
- **How Did This Happen**: `Howdidthishappen_x003f_`
- **Investigation Required**: `InvestigationRequired_x003f_` (Boolean)
- **Status**: `Status` (Choice field)
- **Meeting Date**: `MeetingDate`
- **Submitted By**: `Name` (Person field)
- **Created By**: `Author` (Person)
- **Modified By**: `Editor` (Person)
- **Created**: `Created`
- **Modified**: `Modified`

---

## 2. Power Automate HTTP Request Setup

- **Do NOT use `$expand` for Status or EventType** (they are Choice fields, not lookup fields).
- **Only use `$expand` for Person fields**: `Name`, `Author`, `Editor`.

### Example API Call

```
_api/web/lists/getbytitle('Near Miss - Accident Safety Register')/items
?$select=ID,Title,Date,EventType,EventType/Value,Canyoubrieflyexplainwhathappened,Howdidthishappen_x003f_,InvestigationRequired_x003f_,Status,Status/Value,MeetingDate,Name/Id,Name/Title,Name/EMail,Name/JobTitle,Author/Id,Author/Title,Author/EMail,Author/JobTitle,Editor/Id,Editor/Title,Editor/EMail,Editor/JobTitle,Created,Modified
&$expand=Name,Author,Editor
```

---

## 3. Summary Table

| Display Name                | Internal Name                      | Type      | Expand? |
|-----------------------------|------------------------------------|-----------|---------|
| Title                       | Title                              | Text      | No      |
| Date                        | Date                               | DateTime  | No      |
| Event Type                  | EventType                          | Choice    | No      |
| Brief Explanation           | Canyoubrieflyexplainwhathappened   | Note      | No      |
| How Did This Happen         | Howdidthishappen_x003f_            | Note      | No      |
| Investigation Required      | InvestigationRequired_x003f_       | Boolean   | No      |
| Status                      | Status                             | Choice    | No      |
| Meeting Date                | MeetingDate                        | DateTime  | No      |
| Submitted By                | Name                               | Person    | Yes     |
| Created By                  | Author                             | Person    | Yes     |
| Modified By                 | Editor                             | Person    | Yes     |
| Created                     | Created                            | DateTime  | No      |
| Modified                    | Modified                           | DateTime  | No      |

---

## 4. Power Automate HTTP Action Example

```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://cranfieldglass.sharepoint.com",
      "parameters/method": "GET",
      "parameters/uri": "_api/web/lists/getbytitle('Near Miss - Accident Safety Register')/items?$select=ID,Title,Date,EventType,EventType/Value,Canyoubrieflyexplainwhathappened,Howdidthishappen_x003f_,InvestigationRequired_x003f_,Status,Status/Value,MeetingDate,Name/Id,Name/Title,Name/EMail,Name/JobTitle,Author/Id,Author/Title,Author/EMail,Author/JobTitle,Editor/Id,Editor/Title,Editor/EMail,Editor/JobTitle,Created,Modified&$expand=Name,Author,Editor"
    },
    "host": {
      "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline",
      "connection": "shared_sharepointonline",
      "operationId": "HttpRequest"
    }
  },
  "runAfter": {}
}
```

---

## 5. Using the Output in Power Automate

- **Compose (after HTTP request):**
  ```
  body('Send_an_HTTP_request_to_SharePoint')?['d']?['results']
  ```
- **Apply to each:**  Use the output of the Compose above.

### Inside Apply to each, reference fields:

- `item()?['Title']`
- `item()?['Date']`
- `item()?['EventType']?['Value']`
- `item()?['Canyoubrieflyexplainwhathappened']`
- `item()?['Howdidthishappen_x003f_']`
- `item()?['InvestigationRequired_x003f_']`
- `item()?['Status']?['Value']`
- `item()?['MeetingDate']`
- `item()?['Name']?['Title']` (Submitted By)
- `item()?['Author']?['Title']` (Created By)
- `item()?['Editor']?['Title']` (Modified By)

---

## 6. Removing HTML from Fields

- Use the **HTML to text** action (Data Operations > HTML to text) in Power Automate.
- Input: `item()?['Canyoubrieflyexplainwhathappened']`
- Output: Plain text without HTML tags.

---

## 7. Paging and Performance

- For large lists, enable **Pagination** in the HTTP action settings and set a high threshold.
- For testing, use `$top=5` or `$top=10` to limit results.

---

## 8. Microsoft Graph API (React + Vite Example)

- **Permission:** `Sites.Read.All` (delegated)
- **Auth library:** MSAL 2.0 (`@azure/msal-react`)
- **API endpoint:** `/sites/{tenant}.sharepoint.com:/sites/{site-name}:/lists/{list-id}/items?expand=fields`
- **Consent:** User/admin must consent in Azure

```javascript
import React, { useEffect, useState } from "react";
import { useMsal } from "@azure/msal-react";

const GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0/sites/{tenant}.sharepoint.com:/sites/{site-name}:/lists/{list-id}/items?expand=fields";

export default function SharePointList() {
  const { instance, accounts } = useMsal();
  const [items, setItems] = useState([]);

  useEffect(() => {
    const getItems = async () => {
      const response = await instance.acquireTokenSilent({
        scopes: ["Sites.Read.All"],
        account: accounts[0],
      });
      const accessToken = response.accessToken;

      const res = await fetch(GRAPH_ENDPOINT, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });
      const data = await res.json();
      setItems(data.value);
    };

    if (accounts.length > 0) {
      getItems();
    }
  }, [accounts, instance]);

  return (
    <div>
      <h2>SharePoint List Items</h2>
      <ul>
        {items.map((item) => (
          <li key={item.id}>{item.fields.Title}</li>
        ))}
      </ul>
    </div>
  );
}
```

---

## 9. Troubleshooting

- **Expand errors:** Make sure every `/` field in `$select` is also in `$expand`.
- **Field does not exist:** Double-check the internal name in your schema.
- **Need more user info (like Claims):** Use a second HTTP request to `/siteusers/getbyid({Id})`.

---

**Let me know if you want a step-by-step with screenshots or a sample Parse JSON schema!**
---

## 10. How to Get All Unique Meeting Dates with Replit (Node.js Example)

You can use Replit to fetch all Meeting Dates from your SharePoint list using the REST API and Node.js. Here’s how:

### Example: Node.js (with axios)

```javascript
// Install axios in Replit: npm install axios
const axios = require('axios');

const siteUrl = 'https://cranfieldglass.sharepoint.com';
const listName = 'Near Miss - Accident Safety Register';
const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=MeetingDate`;
const headers = {
  'Accept': 'application/json;odata=verbose',
  'Authorization': 'Bearer YOUR_ACCESS_TOKEN' // Replace with your access token
};

async function getUniqueMeetingDates() {
  let url = apiUrl;
  let allDates = [];

  while (url) {
    const res = await axios.get(url, { headers });
    const results = res.data.d.results;
    allDates.push(...results.map(item => item.MeetingDate));
    url = res.data.d.__next || null; // Paging
  }

  // Remove duplicates and filter out empty/null dates
  const uniqueDates = [...new Set(allDates.filter(Boolean))];
  console.log('Unique Meeting Dates:', uniqueDates);
}

getUniqueMeetingDates();
```

---

## 11. **Power Automate Approach**

1. **Initialize an object variable** (e.g., `dateToIds`) before your loop:
   - **Type:** Object
   - **Value:** `{}`

2. **Apply to each** (over your items array):

   - **Get the date and ID:**
     - `item()?['MeetingDate']`
     - `item()?['ID']`

   - **Add a "Set variable" action** (use the **Expression** tab):
     - **Name:** `dateToIds`
     - **Value:**
       ```json
       union(
         variables('dateToIds'),
         json(concat(
           '{ "', 
           item()?['MeetingDate'], 
           '": ', 
           if(equals(variables('dateToIds')[item()?['MeetingDate']], null), 
             concat('[', string(item()?['ID']), ']'), 
             concat('[', join(variables('dateToIds')[item()?['MeetingDate']], ','), ',', string(item()?['ID']), ']')
           ), 
           '}'
         ))
       )
       ```
     - This expression builds up an object where each key is a date and the value is an array of IDs for that date.

   - **Alternative (simpler):**
     - Use a **"Append to array variable"** for each date, but this requires a separate array per date, which is less dynamic.

3. **After the loop**, your `dateToIds` variable will look like:
   ```json
   {
     "2025-02-05": [8, 9, 10],
     "2025-02-12": [11, 12]
   }
   ```

---

## **Node.js (Replit) Approach**

Here’s how you can do it in Node.js:

```javascript
const dateToIds = {};
allItems.forEach(item => {
  const date = item.MeetingDate;
  const id = item.ID;
  if (!date) return;
  if (!dateToIds[date]) dateToIds[date] = [];
  dateToIds[date].push(id);
});
console.log(dateToIds);
```

---

## **Summary Table**

| Platform        | Step                                                                 |
|-----------------|----------------------------------------------------------------------|
| Power Automate  | Use an object variable, build up arrays of IDs for each date         |
| Node.js/Replit  | Use a JS object, push IDs into arrays keyed by date                  |

---

**This will give you a mapping of each MeetingDate to the list of item IDs for that date.**  
Let me know if you want a full Power Automate step-by-step for this!
