# Word Template Blob Web API

ASP.NET Core Web API sample that:

- Downloads a Word `.docx` template from Azure Blob Storage
- Replaces placeholders like `{{CustomerName}}`
- Clones a formatted table row for line items
- Applies example bold formatting
- Uploads the generated Word document back to Azure Blob Storage

## Placeholder template

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
word-templates/proposal-template.docx
```

## Configuration

Update `appsettings.json`:

```json
{
  "AzureStorage": {
    "AccountUrl": "https://wordtemplates.blob.core.windows.net",
    "TemplateContainer": "word-templates",
    "OutputContainer": "generated-documents",
    "TemplateBlobName": "proposal-template.docx"
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
dotnet run
```

Open Swagger:

```text
https://localhost:5001/swagger
```

or use the HTTP file in the project.

## Test request

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

Plain text placeholders can sometimes be split across multiple Word XML runs if you edit the text heavily in Microsoft Word. For production templates, keep placeholders typed in one pass, or use content controls/bookmarks for more reliable replacement.
