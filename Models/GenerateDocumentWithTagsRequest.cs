using System.ComponentModel.DataAnnotations;

namespace WordTemplateBlobWebApi.Models;

public sealed class GenerateDocumentWithTagsRequest
{
    [Required]
    public Dictionary<string, string> Tags { get; set; } = new();

    public List<Dictionary<string, string>> Items { get; set; } = [];
}
