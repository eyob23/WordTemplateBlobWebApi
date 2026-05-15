using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Options;
using WordTemplateBlobWebApi.Models;
using WordTemplateBlobWebApi.Options;
using BookmarkStart = DocumentFormat.OpenXml.Wordprocessing.BookmarkStart;
using BookmarkEnd = DocumentFormat.OpenXml.Wordprocessing.BookmarkEnd;

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

        var replacements = new Dictionary<string, string>(request.Tags, StringComparer.Ordinal);

        // Replace bookmarks in the main document
        ReplaceBookmarks(mainPart, replacements);

        if (request.Items.Count > 0)
        {
            ReplaceTableRowsWithTagItems(mainPart.Document.Body!, request.Items);
        }

        mainPart.Document.Save();
    }

    private static void ReplaceBookmarks(MainDocumentPart mainPart, Dictionary<string, string> replacements)
    {
        var body = mainPart.Document.Body;

        foreach (var kvp in replacements)
        {
            ReplaceBookmarkContent(body!, kvp.Key, kvp.Value ?? string.Empty);
        }
    }

    private static void ReplaceBookmarkContent(OpenXmlElement element, string bookmarkName, string replacementText)
    {
        // Find BookmarkStart and BookmarkEnd elements
        var bookmarkStart = element
            .Descendants<BookmarkStart>()
            .FirstOrDefault(b => b.Name == bookmarkName);

        if (bookmarkStart is null)
            return;

        // Find the corresponding BookmarkEnd
        var bookmarkEnd = element
            .Descendants<BookmarkEnd>()
            .FirstOrDefault(b => b.Id == bookmarkStart.Id);

        if (bookmarkEnd is null)
            return;

        // Remove existing content between bookmark start and end
        var nodesToRemove = new List<OpenXmlElement>();
        var currentNode = bookmarkStart.NextSibling();

        while (currentNode is not null && currentNode != bookmarkEnd)
        {
            if (currentNode is Run || currentNode is Paragraph)
            {
                nodesToRemove.Add(currentNode);
            }
            currentNode = currentNode.NextSibling();
        }

        // Remove the collected nodes
        foreach (var node in nodesToRemove)
        {
            node.Remove();
        }

        // Create and insert new Run with the replacement text
        var newRun = new Run(new Text(replacementText));

        // Insert the new run after the bookmark start
        bookmarkStart.InsertAfterSelf(newRun);
    }

    private static void ReplaceTableRowsWithTagItems(Body body, IReadOnlyCollection<Dictionary<string, string>> items)
    {
        if (items.Count == 0)
            return;

        var itemKeys = items
            .SelectMany(item => item.Keys)
            .ToHashSet(StringComparer.Ordinal);

        var tables = body.Descendants<Table>();

        foreach (var table in tables)
        {
            var templateRow = table
                .Descendants<TableRow>()
                .FirstOrDefault(row =>
                    row.Descendants<BookmarkStart>().Any(b => b.Name?.Value is string name && itemKeys.Contains(name)));

            if (templateRow is null)
                continue;

            foreach (var item in items)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);

                foreach (var entry in item)
                {
                    ReplaceBookmarkContent(newRow, entry.Key, entry.Value ?? string.Empty);
                }

                table.InsertBefore(newRow, templateRow);
            }

            templateRow.Remove();
        }
    }

}
