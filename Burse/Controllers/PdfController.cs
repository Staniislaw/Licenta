using Microsoft.AspNetCore.Mvc;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using QuestPDF.Drawing;
using System.Text.Json;
using Burse.Models.TemplatePDF;
using Burse.Services.Abstractions;
using System.Text.RegularExpressions;
using System;
using Burse.Data;
using Microsoft.EntityFrameworkCore;
using Burse.Services;
namespace Burse.Controllers
{
    [ApiController]
    [Route("api/pdf")]
    public class PdfController : ControllerBase
    {
        private readonly IFondBurseService _fondBurseService;
        private readonly BurseDBContext _context;
        private readonly IPdfGeneratorService _pdfGeneratorService;
        private readonly IGrupuriService _grupuriService;
        private readonly IStudentService _studentService;

        public PdfController(IFondBurseService fondBurseService, BurseDBContext context, IPdfGeneratorService pdfGeneratorService, IGrupuriService grupuriService, IStudentService studentService)
        {
            _fondBurseService = fondBurseService;
            _context = context;
            _pdfGeneratorService = pdfGeneratorService;
            _grupuriService = grupuriService;
            _studentService = studentService;
        }



        [HttpPost("generate")]
        public async Task<IActionResult> GeneratePdf([FromBody] PdfRequest request)
        {
            try
            {
                var stream = await _pdfGeneratorService.GeneratePdfAsync(request);
                return File(stream, "application/pdf", "generated.pdf");
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("generate-all-pdfs")]
        public async Task<IActionResult> GenerateAllPdfs([FromBody] PdfRequest request)
        {
            // Preia toate grupurile cu valorile lor
            var grupuriPdf = await _grupuriService.GetGrupuriPdfAsync();

            var pdfStreams = new List<(MemoryStream Stream, string FileName)>();

            foreach (var grup in grupuriPdf)
            {
                // Construim un DynamicFields cu valorile grupului, separate prin virgulă
                var dynamicFields = new Dictionary<string, string>
                {
                    { "ProgramStudiu.Dynamic", string.Join(", ", grup.Value) }
                };

                var pdfRequest = new PdfRequest
                {
                    Elements = request.Elements, 
                    DynamicFields = dynamicFields
                };


                var stream = await _pdfGeneratorService.GeneratePdfAsync(pdfRequest);
                stream.Position = 0;

                pdfStreams.Add((stream, $"{grup.Key}.pdf"));
            }

            var zipStream = new MemoryStream();
            using (var archive = new System.IO.Compression.ZipArchive(zipStream, System.IO.Compression.ZipArchiveMode.Create, true))
            {
                foreach (var (stream, fileName) in pdfStreams)
                {
                    var zipEntry = archive.CreateEntry(fileName);
                    using (var entryStream = zipEntry.Open())
                    {
                        await stream.CopyToAsync(entryStream);
                    }
                }
            }
            zipStream.Position = 0;

            return File(zipStream, "application/zip", "all-pdfs.zip");
        }

        [HttpGet("export-excel-studenti")]
        public async Task<IActionResult> ExportExcelStudenti()
        {
            var fileBytes = await _studentService.ExportStudentiExcelAsync();
            var fileName = $"studenti_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            return File(fileBytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);
        }




    }
}
