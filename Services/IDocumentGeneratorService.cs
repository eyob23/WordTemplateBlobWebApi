using WordTemplateBlobWebApi.Models;

namespace WordTemplateBlobWebApi.Services;

public interface IDocumentGeneratorService
{
    Task<GenerateDocumentResponse> GenerateAsync(GenerateDocumentRequest request, CancellationToken cancellationToken = default);
}
