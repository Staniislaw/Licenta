using Burse.Models.TemplatePDF;

namespace Burse.Services.Abstractions
{
    public interface IPdfGeneratorService
    {
        Task<MemoryStream> GeneratePdfAsync(PdfRequest request);
    }

}
