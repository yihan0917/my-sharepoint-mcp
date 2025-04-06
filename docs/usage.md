# SharePoint MCP Server - Usage Guide

> **DISCLAIMER**: This is an unofficial, community-developed project that is not affiliated with, endorsed by, or related to Microsoft Corporation. This documentation describes interaction with Microsoft services but is not approved by or associated with Microsoft.

This guide demonstrates how to use your SharePoint MCP server with Claude and other LLM applications.

## Setup

1. First, edit the `.env` file with your SharePoint connection details:
   - Get your tenant ID from the Azure Portal
   - Register an application in Azure AD to get client ID and secret
   - Set your SharePoint site URL 

2. Run the server:
   ```bash
   python src/server.py
   ```

3. For development and testing, use MCP Inspector:
   ```bash
   mcp dev src/server.py
   ```

4. To install in Claude Desktop:
   ```bash
   mcp install src/server.py --name "SharePoint Assistant"
   ```

## Available Resources

Access SharePoint content through these resource URIs:

1. Site Information:
   ```
   sharepoint://site-info
   ```

2. Document Library Contents:
   ```
   sharepoint://documents/{library_name}
   ```
   Example: `sharepoint://documents/Shared Documents`

3. List Data:
   ```
   sharepoint://list/{list_name}
   ```
   Example: `sharepoint://list/Tasks`

4. Document Content:
   ```
   sharepoint://document/{library_name}/{file_path}
   ```
   Example: `sharepoint://document/Shared Documents/report.docx`

## Available Tools

Your SharePoint MCP provides these tools to Claude:

1. `search_sharepoint(query: str)`: Search across your SharePoint site
   - Example: "Weather forecast document"

2. `download_document(library_name: str, file_path: str)`: Download a specific document
   - Example: download_document("Shared Documents", "quarterly-report.pdf")

3. `get_sharepoint_lists()`: Get all lists in the SharePoint site

4. `create_list_item(list_name: str, item_data: Dict[str, Any])`: Create a new item in a list
   - Example: create_list_item("Tasks", {"Title": "Review document", "Status": "Not Started"})

## Example Prompts

Here are some examples of how to interact with the SharePoint MCP in Claude:

### Example 1: Browse Document Libraries
```
Can you tell me what documents are in the "Marketing Materials" library on my SharePoint site?
```

### Example 2: Search for Content
```
Can you search my SharePoint site for documents related to "quarterly revenue"?
```

### Example 3: Get List Items
```
Show me all the items in the "Tasks" list on my SharePoint site.
```

### Example 4: Create a New Task
```
Please create a new task in the "Tasks" list with the title "Prepare monthly report" and status "Not Started".
```

### Example 5: Analyze a Document
```
Can you download and analyze the "sales_data.xlsx" file from my "Reports" document library?
```