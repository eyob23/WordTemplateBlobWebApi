# Word Template Blob Web API

ASP.NET Core Web API sample that:

- Downloads a Word `.docx` template from Azure Blob Storage
- Supports placeholder replacement (`{{CustomerName}}`) and bookmark/tag replacement
- Clones a formatted table row for line items (for both approaches)
- Uploads the generated Word document back to Azure Blob Storage

## Endpoints

- `POST /api/documents/generate`
  Uses placeholder-based replacement (`{{CustomerName}}`, `{{ItemName}}`, etc.).
- `POST /api/documents/generate-with-tags`
  Uses bookmark/tag replacement with a generic key/value payload.

## Bookmark/tag template (recommended)

Create a Word file named `template.docx`.

Add bookmarks for header fields, for example:

```text
CustomerName
ProjectName
PreparedBy
GeneratedDate
```

Add a table with a header row and one template row.
Inside the template row, add bookmarks for item fields, for example:

```text
ItemName
ItemDescription
ItemAmount
```

Format the template row however you want in Word. The code clones that row for each item object in the request.

## Placeholder template (legacy)

Create a Word file named `proposal-template.docx`.

Add normal placeholders:

```text
Customer: {{CustomerName}}
Project: {{ProjectName}}
Prepared By: {{PreparedBy}}
Generated Date: {{GeneratedDate}}
```

Add a table with a header row and one template row:

| Item | Description | Amount |
|---|---|---|
| {{ItemName}} | {{ItemDescription}} | {{ItemAmount}} |

Format that template row however you want in Word. The code clones that row, so the generated rows keep the same formatting.

## Azure Blob Storage setup

Create two containers:

```text
word-templates
generated-documents
```

Upload your template to:

```text
word-templates/template.docx
```

## Configuration

Update `appsettings.json`:

```json
{
  "AzureStorage": {
    "AccountUrl": "https://wordtemplates.blob.core.windows.net",
    "TemplateContainer": "word-templates",
    "OutputContainer": "generated-documents",
    "TemplateBlobName": "template.docx"
  }
}
```

## Local auth

This sample uses `DefaultAzureCredential`.

For local development:

```bash
az login
```

Your user needs Storage Blob Data Reader on the template container and Storage Blob Data Contributor on the output container.

For Azure App Service or Azure Functions, enable managed identity and grant it the same RBAC permissions.

## Run

```bash
cd WordTemplateBlobWebApi
dotnet restore
ASPNETCORE_ENVIRONMENT=Development dotnet run
```

Open Swagger:

```text
http://localhost:5000/swagger/index.html
```

or use the HTTP file in the project.

## Test request (bookmark/tag endpoint)

Use with `POST /api/documents/generate-with-tags`:

```json
{
  "tags": {
    "CustomerName": "Phoenix Mechanical CA",
    "ProjectName": "HVAC Proposal",
    "PreparedBy": "Eyobe",
    "GeneratedDate": "2026-05-14"
  },
  "items": [
    {
      "ItemName": "AC Installation",
      "ItemDescription": "Install new high-efficiency AC system",
      "ItemAmount": "$4,500"
    },
    {
      "ItemName": "Duct Work",
      "ItemDescription": "Repair and seal duct system",
      "ItemAmount": "$1,200"
    }
  ]
}
```

## Test request (placeholder endpoint)

Use with `POST /api/documents/generate`:

```json
{
  "customerName": "Phoenix Mechanical CA",
  "projectName": "HVAC Proposal",
  "preparedBy": "Eyobe",
  "items": [
    {
      "itemName": "AC Installation",
      "itemDescription": "Install new high-efficiency AC system",
      "amount": 4500
    },
    {
      "itemName": "Duct Work",
      "itemDescription": "Repair and seal duct system",
      "amount": 1200
    }
  ]
}
```

## Important production note

Plain text placeholders can be split across multiple Word XML runs if edited heavily in Microsoft Word. For production templates, bookmark/tag replacement is usually more reliable.
