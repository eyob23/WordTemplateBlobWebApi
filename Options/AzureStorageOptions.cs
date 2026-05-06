namespace WordTemplateBlobWebApi.Options;

public sealed class AzureStorageOptions
{
    public const string SectionName = "AzureStorage";

    public string AccountUrl { get; set; } = string.Empty;
    public string TemplateContainer { get; set; } = string.Empty;
    public string OutputContainer { get; set; } = string.Empty;
    public string TemplateBlobName { get; set; } = string.Empty;
}
