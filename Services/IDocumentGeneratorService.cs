using WordTemplateBlobWebApi.Models;

namespace WordTemplateBlobWebApi.Services;

public interface IDocumentGeneratorService
{
    Task<GenerateDocumentResponse> GenerateAsync(GenerateDocumentRequest request, CancellationToken cancellationToken = default);

    Task<GenerateDocumentResponse> GenerateWithTagsAsync(GenerateDocumentWithTagsRequest request, CancellationToken cancellationToken = default);
    Task<GenerateDocumentResponse> GenerateWithSdtAsync(GenerateDocumentWithTagsRequest request, CancellationToken cancellationToken = default);
    Task<string> UploadSdtTemplateAsync(CancellationToken cancellationToken = default);
}
