# Word Template Blob Web API

ASP.NET Core Web API that generates Word `.docx` documents by populating templates stored in Azure Blob Storage. Supports three approaches: placeholder text replacement, bookmark/tag replacement, and Word content controls (SDT).

## Endpoints

| Method | Route | Description |
|--------|-------|-------------|
| `POST` | `/api/documents/generate` | Placeholder-based replacement (`{{CustomerName}}` etc.) |
| `POST` | `/api/documents/generate-with-tags` | Generic key/value tag replacement (same `{{}}` format, flexible schema) |
| `POST` | `/api/documents/generate-with-content-controls` | Word content control (SDT) replacement |
| `POST` | `/api/documents/sdt-template/upload` | Generates and uploads the blank SDT template to blob storage |

---

## Approach 1 — Placeholder replacement (`generate`)

Template file: `template.docx` (configured via `TemplateBlobName`)

Use `{{Placeholder}}` text anywhere in the document:

```text
Customer: {{CustomerName}}
Project:  {{ProjectName}}
Prepared By: {{PreparedBy}}
Generated Date: {{GeneratedDate}}
```

Add a table with a header row and one template row:

| Item | Description | Amount |
|---|---|---|
| `{{ItemName}}` | `{{ItemDescription}}` | `{{ItemAmount}}` |

**Request body:**
```json
{
  "customerName": "Acme Corp",
  "projectName": "Cloud Migration",
  "preparedBy": "Sales Team",
  "items": [
    { "itemName": "Discovery", "itemDescription": "Initial analysis", "amount": 5000 }
  ]
}
```

---

## Approach 2 — Tag replacement (`generate-with-tags`)

Uses the same `template.docx` and `{{}}` placeholder format as approach 1, but accepts a generic key/value payload — useful when field names vary per request.

**Request body:**
```json
{
  "tags": {
    "CustomerName": "Northwind Traders",
    "ProjectName": "ERP Modernization",
    "PreparedBy": "Enterprise Solutions Team",
    "GeneratedDate": "2026-05-15"
  },
  "items": [
    { "ItemName": "Requirements Gathering", "ItemDescription": "Gap analysis", "ItemAmount": "8500" },
    { "ItemName": "System Architecture", "ItemDescription": "Technical blueprint", "ItemAmount": "15000" }
  ]
}
```

Tag keys are automatically wrapped as `{{Key}}` before matching, so `"CustomerName"` replaces `{{CustomerName}}` in the template.

---

## Approach 3 — Word content controls (`generate-with-content-controls`)

Uses a separate SDT template (`sdt-template.docx`) with structured Word content controls. Content controls are visible and labelled in the Word editor.

### Step 1 — Upload the SDT template (once)

```bash
POST /api/documents/sdt-template/upload
```

This generates the blank template and uploads it to `word-templates/sdt-template.docx`.

### Step 2 — Generate documents

**Request body** (same schema as `generate`):
```json
{
  "customerName": "Northwind Traders",
  "projectName": "ERP Modernization",
  "preparedBy": "Enterprise Solutions Team",
  "items": [
    { "itemName": "Requirements Gathering", "itemDescription": "Gap analysis", "amount": 8500 },
    { "itemName": "System Architecture", "itemDescription": "Technical blueprint", "amount": 15000 }
  ]
}
```

Content controls are matched by their `Tag` value (`CustomerName`, `ProjectName`, `PreparedBy`, `GeneratedDate`, `ItemName`, `ItemDescription`, `ItemAmount`).

---

## Azure Blob Storage setup

Create two containers:

```text
word-templates        ← templates live here
generated-documents   ← generated files are written here
```

Upload your placeholder template to `word-templates/template.docx` (or call `POST /api/documents/sdt-template/upload` for the SDT template).

---

## Configuration

```json
{
  "AzureStorage": {
    "AccountUrl": "https://<your-account>.blob.core.windows.net",
    "TemplateContainer": "word-templates",
    "OutputContainer": "generated-documents",
    "TemplateBlobName": "template.docx",
    "SdtTemplateBlobName": "sdt-template.docx"
  }
}
```

---

## Authentication

Uses `DefaultAzureCredential`. For local development:

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
