using Microsoft.AspNetCore.Mvc;
using WordTemplateBlobWebApi.Models;
using WordTemplateBlobWebApi.Services;

namespace WordTemplateBlobWebApi.Controllers;

[ApiController]
[Route("api/documents")]
public sealed class DocumentsController : ControllerBase
{
    private readonly IDocumentGeneratorService _documentGeneratorService;

    public DocumentsController(IDocumentGeneratorService documentGeneratorService)
    {
        _documentGeneratorService = documentGeneratorService;
    }

    [HttpPost("generate")]
    [ProducesResponseType(typeof(GenerateDocumentResponse), StatusCodes.Status200OK)]
    public async Task<ActionResult<GenerateDocumentResponse>> Generate(
        [FromBody] GenerateDocumentRequest request,
        CancellationToken cancellationToken)
    {
        var response = await _documentGeneratorService.GenerateAsync(request, cancellationToken);
        return Ok(response);
    }

    [HttpPost("generate-with-tags")]
    [ProducesResponseType(typeof(GenerateDocumentResponse), StatusCodes.Status200OK)]
    public async Task<ActionResult<GenerateDocumentResponse>> GenerateWithTags(
        [FromBody] GenerateDocumentWithTagsRequest request,
        CancellationToken cancellationToken)
    {
        var response = await _documentGeneratorService.GenerateWithTagsAsync(request, cancellationToken);
        return Ok(response);
    }

    [HttpPost("generate-with-content-controls")]
    [ProducesResponseType(typeof(GenerateDocumentResponse), StatusCodes.Status200OK)]
    public async Task<ActionResult<GenerateDocumentResponse>> GenerateWithContentControls(
        [FromBody] GenerateDocumentRequest request,
        CancellationToken cancellationToken)
    {
        var response = await _documentGeneratorService.GenerateWithSdtAsync(request, cancellationToken);
        return Ok(response);
    }

    /// <summary>
    /// Generates the blank SDT (content control) Word template and uploads it to blob storage.
    /// Call this once before using generate-with-content-controls.
    /// </summary>
    [HttpPost("sdt-template/upload")]
    [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
    public async Task<IActionResult> UploadSdtTemplate(CancellationToken cancellationToken)
    {
        var blobUrl = await _documentGeneratorService.UploadSdtTemplateAsync(cancellationToken);
        return Ok(new { message = "SDT template uploaded successfully.", blobUrl });
    }
}
