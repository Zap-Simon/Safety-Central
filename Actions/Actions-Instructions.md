# SharePoint List API Integration Instructions: Actions List

## 1. How to Find Internal Names for Your Flow

Use the `Actions-Schema-Tidied.json` file to look up the `InternalName` for each field you want to use in Power Automate or code. Example fields:

- **Title**: `Title`
- **Action/Event**: `Idea_x002f_Event`
- **Status**: `Status` (Choice field)
- **Priority**: `Priority` (Choice field)
- **List Origin**: `ListOrigin` (Choice field)
- **Link**: `Link` (Hyperlink)
- **Action Start Date**: `ActionStartDate` (DateTime)
- **Assigned To**: `Assigend_x0020_to` (Person field)
- **Submitted By**: `SubmittedBy` (Person field)
- **Idea Type**: `IdeaType` (Choice field)
- **Category**: `Catagorie` (Choice field)
- **Idea Date**: `IdeaDate` (DateTime)
- **Notes/Comments**: `Notes_x0020__x002f__x0020_Commen`
- **Created By**: `Author` (Person)
- **Modified By**: `Editor` (Person)

---

## 2. Power Automate HTTP Request Setup

- **Do NOT use `$expand` for Status, IdeaType, Catagorie, Priority, ListOrigin, Link, or ActionStartDate** (they are not Person fields).
- **Only use `$expand` for Person fields**: `Assigend_x0020_to`, `SubmittedBy`, `Author`, `Editor`.

### Example API Call

```
_api/web/lists/getbytitle('Actions')/items
?$select=ID,Title,Idea_x002f_Event,Status,Priority,ListOrigin,Link,ActionStartDate,Assigend_x0020_to/Id,Assigend_x0020_to/Title,Assigend_x0020_to/EMail,Assigend_x0020_to/JobTitle,SubmittedBy/Id,SubmittedBy/Title,SubmittedBy/EMail,SubmittedBy/JobTitle,IdeaType,Catagorie,IdeaDate,Notes_x0020__x002f__x0020_Commen,Author/Id,Author/Title,Author/EMail,Author/JobTitle,Editor/Id,Editor/Title,Editor/EMail,Editor/JobTitle,Created,Modified
&$expand=Assigend_x0020_to,SubmittedBy,Author,Editor
```

- **Base site URL:** `https://cranfieldglass.sharepoint.com/sites/Team_Actions_Meetings`
- **Full API endpoint:**
  ```
  https://cranfieldglass.sharepoint.com/sites/Team_Actions_Meetings/_api/web/lists/getbytitle('Actions')/items?... 
  ```

---

## 3. Summary Table

| Display Name      | Internal Name                  | Type      | Expand? |
|-------------------|-------------------------------|-----------|---------|
| Title             | Title                         | Text      | No      |
| Action/Event      | Idea_x002f_Event              | Note      | No      |
| Status            | Status                        | Choice    | No      |
| Priority          | Priority                      | Choice    | No      |
| List Origin       | ListOrigin                    | Choice    | No      |
| Link              | Link                          | Hyperlink | No      |
| Action Start Date | ActionStartDate               | DateTime  | No      |
| Assigned To       | Assigend_x0020_to             | Person    | Yes     |
| Submitted By      | SubmittedBy                   | Person    | Yes     |
| Idea Type         | IdeaType                      | Choice    | No      |
| Category          | Catagorie                     | Choice    | No      |
| Idea Date         | IdeaDate                      | DateTime  | No      |
| Notes/Comments    | Notes_x0020__x002f__x0020_Commen | Note  | No      |
| Created By        | Author                        | Person    | Yes     |
| Modified By       | Editor                        | Person    | Yes     |

---

## 4. Power Automate HTTP Action Example

```json
{
  "type": "OpenApiConnection",
  "inputs": {
    "parameters": {
      "dataset": "https://cranfieldglass.sharepoint.com/sites/Team_Actions_Meetings",
      "parameters/method": "GET",
      "parameters/uri": "_api/web/lists/getbytitle('Actions')/items?$select=ID,Title,Idea_x002f_Event,Status,Priority,ListOrigin,Link,ActionStartDate,Assigend_x0020_to/Id,Assigend_x0020_to/Title,Assigend_x0020_to/EMail,Assigend_x0020_to/JobTitle,SubmittedBy/Id,SubmittedBy/Title,SubmittedBy/EMail,SubmittedBy/JobTitle,IdeaType,Catagorie,IdeaDate,Notes_x0020__x002f__x0020_Commen,Author/Id,Author/Title,Author/EMail,Author/JobTitle,Editor/Id,Editor/Title,Editor/EMail,Editor/JobTitle,Created,Modified&$expand=Assigend_x0020_to,SubmittedBy,Author,Editor"
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
- `item()?['Idea_x002f_Event']`
- `item()?['Status']`
- `item()?['Priority']`
- `item()?['ListOrigin']`
- `item()?['Link']`
- `item()?['ActionStartDate']`
- `item()?['Assigend_x0020_to']?['Title']`
- `item()?['SubmittedBy']?['Title']`
- `item()?['IdeaType']`
- `item()?['Catagorie']`
- `item()?['IdeaDate']`
- `item()?['Notes_x0020__x002f__x0020_Commen']`
- `item()?['Author']?['Title']`
- `item()?['Editor']?['Title']`

---

## 6. Removing HTML from Fields

- Use the **HTML to text** action (Data Operations > HTML to text) in Power Automate.
- Input: `item()?['Idea_x002f_Event']` or `item()?['Notes_x0020__x002f__x0020_Commen']`
- Output: Plain text without HTML tags.

---

## 7. Paging and Performance

- For large lists, enable **Pagination** in the HTTP action settings and set a high threshold.
- For testing, use `$top=5` or `$top=10` to limit results.

---

## 8. Grouping IDs by Date (Power Automate & Node.js)

You can group item IDs by a date field (e.g., `IdeaDate` or `ActionStartDate`) in both Power Automate and Node.js. See the Business Ideas instructions for detailed steps—replace `MeetingDate` with your chosen field.

---

## 9. Troubleshooting

- **Expand errors:** Make sure every `/` field in `$select` is also in `$expand`.
- **Field does not exist:** Double-check the internal name in your schema.
- **Need more user info (like Claims):** Use a second HTTP request to `/siteusers/getbyid({Id})`.

---

**Let me know if you want a step-by-step with screenshots or a sample Parse JSON schema for the Actions list!**
