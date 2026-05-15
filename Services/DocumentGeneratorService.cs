using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Options;
using WordTemplateBlobWebApi.Models;
using WordTemplateBlobWebApi.Options;
using WordTemplateBlobWebApi.Utilities;

namespace WordTemplateBlobWebApi.Services;

public sealed class DocumentGeneratorService : IDocumentGeneratorService
{
    private readonly BlobServiceClient _blobServiceClient;
    private readonly AzureStorageOptions _options;
    private readonly ILogger<DocumentGeneratorService> _logger;

    public DocumentGeneratorService(
        BlobServiceClient blobServiceClient,
        IOptions<AzureStorageOptions> options,
        ILogger<DocumentGeneratorService> logger)
    {
        _blobServiceClient = blobServiceClient;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<GenerateDocumentResponse> GenerateAsync(
        GenerateDocumentRequest request,
        CancellationToken cancellationToken = default)
    {
        ValidateOptions();

        var templateContainer = _blobServiceClient.GetBlobContainerClient(_options.TemplateContainer);
        var outputContainer = _blobServiceClient.GetBlobContainerClient(_options.OutputContainer);

        await outputContainer.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

        var templateBlob = templateContainer.GetBlobClient(_options.TemplateBlobName);

        if (!await templateBlob.ExistsAsync(cancellationToken))
        {
            throw new FileNotFoundException(
                $"Template blob '{_options.TemplateBlobName}' was not found in container '{_options.TemplateContainer}'.");
        }

        await using var templateStream = new MemoryStream();
        await templateBlob.DownloadToAsync(templateStream, cancellationToken);
        templateStream.Position = 0;

        await using var outputStream = new MemoryStream();
        await templateStream.CopyToAsync(outputStream, cancellationToken);
        outputStream.Position = 0;

        PopulateWordTemplate(outputStream, request);
        outputStream.Position = 0;

        var outputBlobName = $"generated/proposal-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}.docx";
        var outputBlob = outputContainer.GetBlobClient(outputBlobName);

        await outputBlob.UploadAsync(
            outputStream,
            new BlobHttpHeaders
            {
                ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            cancellationToken: cancellationToken);

        _logger.LogInformation("Generated document uploaded to blob {BlobName}", outputBlobName);

        return new GenerateDocumentResponse
        {
            Message = "Document generated successfully.",
            BlobName = outputBlobName,
            BlobUrl = outputBlob.Uri.ToString()
        };
    }

    public async Task<GenerateDocumentResponse> GenerateWithTagsAsync(
        GenerateDocumentWithTagsRequest request,
        CancellationToken cancellationToken = default)
    {
        ValidateOptions();

        var templateContainer = _blobServiceClient.GetBlobContainerClient(_options.TemplateContainer);
        var outputContainer = _blobServiceClient.GetBlobContainerClient(_options.OutputContainer);

        await outputContainer.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

        var templateBlob = templateContainer.GetBlobClient(_options.TemplateBlobName);

        if (!await templateBlob.ExistsAsync(cancellationToken))
        {
            throw new FileNotFoundException(
                $"Template blob '{_options.TemplateBlobName}' was not found in container '{_options.TemplateContainer}'.");
        }

        await using var templateStream = new MemoryStream();
        await templateBlob.DownloadToAsync(templateStream, cancellationToken);
        templateStream.Position = 0;

        await using var outputStream = new MemoryStream();
        await templateStream.CopyToAsync(outputStream, cancellationToken);
        outputStream.Position = 0;

        PopulateWordTemplateWithTags(outputStream, request);
        outputStream.Position = 0;

        var outputBlobName = $"generated/proposal-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}.docx";
        var outputBlob = outputContainer.GetBlobClient(outputBlobName);

        await outputBlob.UploadAsync(
            outputStream,
            new BlobHttpHeaders
            {
                ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            cancellationToken: cancellationToken);

        _logger.LogInformation("Generated document with tags uploaded to blob {BlobName}", outputBlobName);

        return new GenerateDocumentResponse
        {
            Message = "Document generated successfully using word tags.",
            BlobName = outputBlobName,
            BlobUrl = outputBlob.Uri.ToString()
        };
    }

    public async Task<GenerateDocumentResponse> GenerateWithSdtAsync(
        GenerateDocumentRequest request,
        CancellationToken cancellationToken = default)
    {
        ValidateOptions();

        if (string.IsNullOrWhiteSpace(_options.SdtTemplateBlobName))
            throw new InvalidOperationException("AzureStorage:SdtTemplateBlobName is required.");

        var templateContainer = _blobServiceClient.GetBlobContainerClient(_options.TemplateContainer);
        var outputContainer = _blobServiceClient.GetBlobContainerClient(_options.OutputContainer);

        await outputContainer.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

        var templateBlob = templateContainer.GetBlobClient(_options.SdtTemplateBlobName);

        if (!await templateBlob.ExistsAsync(cancellationToken))
        {
            throw new FileNotFoundException(
                $"SDT template blob '{_options.SdtTemplateBlobName}' was not found in container '{_options.TemplateContainer}'.");
        }

        await using var templateStream = new MemoryStream();
        await templateBlob.DownloadToAsync(templateStream, cancellationToken);
        templateStream.Position = 0;

        await using var outputStream = new MemoryStream();
        await templateStream.CopyToAsync(outputStream, cancellationToken);
        outputStream.Position = 0;

        PopulateWordTemplateSdt(outputStream, request);
        outputStream.Position = 0;

        var outputBlobName = $"generated/proposal-sdt-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}.docx";
        var outputBlob = outputContainer.GetBlobClient(outputBlobName);

        await outputBlob.UploadAsync(
            outputStream,
            new BlobHttpHeaders
            {
                ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            cancellationToken: cancellationToken);

        _logger.LogInformation("Generated SDT document uploaded to blob {BlobName}", outputBlobName);

        return new GenerateDocumentResponse
        {
            Message = "Document generated successfully using content controls.",
            BlobName = outputBlobName,
            BlobUrl = outputBlob.Uri.ToString()
        };
    }

    private void ValidateOptions()
    {
        if (string.IsNullOrWhiteSpace(_options.TemplateContainer))
            throw new InvalidOperationException("AzureStorage:TemplateContainer is required.");

        if (string.IsNullOrWhiteSpace(_options.OutputContainer))
            throw new InvalidOperationException("AzureStorage:OutputContainer is required.");

        if (string.IsNullOrWhiteSpace(_options.TemplateBlobName))
            throw new InvalidOperationException("AzureStorage:TemplateBlobName is required.");
    }

    private static void PopulateWordTemplate(Stream documentStream, GenerateDocumentRequest request)
    {
        documentStream.Position = 0;

        using var wordDoc = WordprocessingDocument.Open(documentStream, true);

        var body = wordDoc.MainDocumentPart?.Document.Body
                   ?? throw new InvalidOperationException("Word document body was not found.");

        var replacements = new Dictionary<string, string>
        {
            ["{{CustomerName}}"] = request.CustomerName,
            ["{{ProjectName}}"] = request.ProjectName,
            ["{{PreparedBy}}"] = request.PreparedBy,
            ["{{GeneratedDate}}"] = DateTime.UtcNow.ToString("yyyy-MM-dd")
        };

        ReplaceTextInElement(body, replacements);
        ReplaceTableRows(body, request.Items);

        // Example: bold specific generated text after replacement.
        BoldTextInElement(body, request.CustomerName);
        BoldTextInElement(body, request.ProjectName);

        wordDoc.MainDocumentPart!.Document.Save();
    }

    private static void ReplaceTableRows(Body body, IReadOnlyCollection<DocumentLineItem> items)
    {
        var tables = body.Descendants<Table>();

        foreach (var table in tables)
        {
            var templateRow = table
                .Descendants<TableRow>()
                .FirstOrDefault(row => row.InnerText.Contains("{{ItemName}}", StringComparison.OrdinalIgnoreCase));

            if (templateRow is null)
                continue;

            foreach (var item in items)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);

                ReplaceTextInElement(newRow, new Dictionary<string, string>
                {
                    ["{{ItemName}}"] = item.ItemName,
                    ["{{ItemDescription}}"] = item.ItemDescription,
                    ["{{ItemAmount}}"] = item.Amount.ToString("C")
                });

                table.InsertBefore(newRow, templateRow);
            }

            templateRow.Remove();
        }
    }

    private static void ReplaceTextInElement(OpenXmlElement element, Dictionary<string, string> replacements)
    {
        foreach (var text in element.Descendants<Text>())
        {
            foreach (var replacement in replacements)
            {
                if (text.Text.Contains(replacement.Key, StringComparison.Ordinal))
                {
                    text.Text = text.Text.Replace(replacement.Key, replacement.Value ?? string.Empty, StringComparison.Ordinal);
                }
            }
        }
    }

    private static void BoldTextInElement(OpenXmlElement element, string searchText)
    {
        if (string.IsNullOrWhiteSpace(searchText))
            return;

        foreach (var text in element.Descendants<Text>())
        {
            if (!text.Text.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                continue;

            var run = text.Ancestors<Run>().FirstOrDefault();
            if (run is null)
                continue;

            run.RunProperties ??= new RunProperties();
            run.RunProperties.Bold = new Bold();
        }
    }

    private static void PopulateWordTemplateWithTags(Stream documentStream, GenerateDocumentWithTagsRequest request)
    {
        documentStream.Position = 0;

        using var wordDoc = WordprocessingDocument.Open(documentStream, true);

        var mainPart = wordDoc.MainDocumentPart
                       ?? throw new InvalidOperationException("Word document main part was not found.");

        var body = mainPart.Document.Body
                   ?? throw new InvalidOperationException("Word document body was not found.");

        // Wrap each key with {{ }} to match the {{Placeholder}} text format in the template
        var textReplacements = request.Tags.ToDictionary(
            kvp => $"{{{{{kvp.Key}}}}}",
            kvp => kvp.Value ?? string.Empty,
            StringComparer.Ordinal);

        ReplaceTextInElement(body, textReplacements);

        if (request.Items.Count > 0)
        {
            ReplaceTableRowsWithDictItems(body, request.Items);
        }

        mainPart.Document.Save();
    }

    private static void ReplaceTableRowsWithDictItems(Body body, IReadOnlyCollection<Dictionary<string, string>> items)
    {
        if (items.Count == 0)
            return;

        // Find the template row by looking for any {{key}} placeholder from the first item
        var firstItemPlaceholders = items.First().Keys
            .Select(k => $"{{{{{k}}}}}")
            .ToList();

        foreach (var table in body.Descendants<Table>().ToList())
        {
            var templateRow = table
                .Descendants<TableRow>()
                .FirstOrDefault(row => firstItemPlaceholders.Any(placeholder =>
                    row.InnerText.Contains(placeholder, StringComparison.OrdinalIgnoreCase)));

            if (templateRow is null)
                continue;

            foreach (var item in items)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);

                var rowReplacements = item.ToDictionary(
                    kvp => $"{{{{{kvp.Key}}}}}",
                    kvp => kvp.Value ?? string.Empty,
                    StringComparer.Ordinal);

                ReplaceTextInElement(newRow, rowReplacements);
                table.InsertBefore(newRow, templateRow);
            }

            templateRow.Remove();
            break; // only process the first matching table
        }
    }

    public async Task<string> UploadSdtTemplateAsync(CancellationToken cancellationToken = default)
    {
        ValidateOptions();

        if (string.IsNullOrWhiteSpace(_options.SdtTemplateBlobName))
            throw new InvalidOperationException("AzureStorage:SdtTemplateBlobName is required.");

        var templateContainer = _blobServiceClient.GetBlobContainerClient(_options.TemplateContainer);
        await templateContainer.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

        var tempPath = Path.Combine(Path.GetTempPath(), $"sdt-template-{Guid.NewGuid():N}.docx");
        try
        {
            WordTemplateGeneratorSdt.GenerateTemplate(tempPath);

            await using var stream = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            var blobClient = templateContainer.GetBlobClient(_options.SdtTemplateBlobName);

            await blobClient.UploadAsync(
                stream,
                new Azure.Storage.Blobs.Models.BlobHttpHeaders
                {
                    ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                },
                cancellationToken: cancellationToken);

            _logger.LogInformation("SDT template uploaded to blob {BlobName}", _options.SdtTemplateBlobName);
            return blobClient.Uri.ToString();
        }
        finally
        {
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }
    }

    // -------------------------------------------------------------------------
    // SDT (Content Control) population
    // -------------------------------------------------------------------------

    private static void PopulateWordTemplateSdt(Stream documentStream, GenerateDocumentRequest request)
    {
        documentStream.Position = 0;

        using var wordDoc = WordprocessingDocument.Open(documentStream, true);

        var body = wordDoc.MainDocumentPart?.Document.Body
                   ?? throw new InvalidOperationException("Word document body was not found.");

        var fieldValues = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["CustomerName"]  = request.CustomerName,
            ["ProjectName"]   = request.ProjectName,
            ["PreparedBy"]    = request.PreparedBy,
            ["GeneratedDate"] = DateTime.UtcNow.ToString("yyyy-MM-dd")
        };

        // Populate block-level content controls (SdtBlock)
        foreach (var sdt in body.Descendants<SdtBlock>().ToList())
        {
            var tag = sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
            if (tag is null || !fieldValues.TryGetValue(tag, out var value))
                continue;

            var textElement = sdt.Descendants<Text>().FirstOrDefault();
            if (textElement is not null)
                textElement.Text = value;

            // Remove ShowingPlaceholder flag so it renders as real content
            sdt.SdtProperties?.GetFirstChild<ShowingPlaceholder>()?.Remove();
        }

        // Populate table rows with inline content controls (SdtRun)
        if (request.Items.Count > 0)
            PopulateSdtTableRows(body, request.Items);

        wordDoc.MainDocumentPart!.Document.Save();
    }

    private static void PopulateSdtTableRows(Body body, IReadOnlyCollection<DocumentLineItem> items)
    {
        foreach (var table in body.Descendants<Table>().ToList())
        {
            // Find the template row — the one that contains SdtRun controls tagged with item field names
            var templateRow = table
                .Descendants<TableRow>()
                .FirstOrDefault(row =>
                    row.Descendants<SdtRun>()
                       .Any(sdt => sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value is "ItemName"
                                                                                        or "ItemDescription"
                                                                                        or "ItemAmount"));

            if (templateRow is null)
                continue;

            foreach (var item in items)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);

                SetSdtRunText(newRow, "ItemName",        item.ItemName);
                SetSdtRunText(newRow, "ItemDescription", item.ItemDescription);
                SetSdtRunText(newRow, "ItemAmount",      item.Amount.ToString("C"));

                table.InsertBefore(newRow, templateRow);
            }

            templateRow.Remove();
        }
    }

    private static void SetSdtRunText(OpenXmlElement element, string tag, string value)
    {
        var sdt = element
            .Descendants<SdtRun>()
            .FirstOrDefault(s => s.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == tag);

        if (sdt is null)
            return;

        var text = sdt.Descendants<Text>().FirstOrDefault();
        if (text is not null)
            text.Text = value;

        sdt.SdtProperties?.GetFirstChild<ShowingPlaceholder>()?.Remove();
    }

}
