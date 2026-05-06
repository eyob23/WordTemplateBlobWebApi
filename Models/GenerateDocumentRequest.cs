using System.ComponentModel.DataAnnotations;

namespace WordTemplateBlobWebApi.Models;

public sealed class GenerateDocumentRequest
{
    [Required]
    public string CustomerName { get; set; } = string.Empty;

    [Required]
    public string ProjectName { get; set; } = string.Empty;

    public string PreparedBy { get; set; } = string.Empty;

    public List<DocumentLineItem> Items { get; set; } = [];
}

public sealed class DocumentLineItem
{
    [Required]
    public string ItemName { get; set; } = string.Empty;

    public string ItemDescription { get; set; } = string.Empty;

    public decimal Amount { get; set; }
}
