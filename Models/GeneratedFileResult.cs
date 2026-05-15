namespace WordTemplateBlobWebApi.Models;

public sealed class GeneratedFileResult
{
    public byte[] Content { get; set; } = [];
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
}
