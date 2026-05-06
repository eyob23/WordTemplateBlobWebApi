namespace WordTemplateBlobWebApi.Models;

public sealed class GenerateDocumentResponse
{
    public string Message { get; set; } = string.Empty;
    public string BlobName { get; set; } = string.Empty;
    public string BlobUrl { get; set; } = string.Empty;
}
