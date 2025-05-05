using Azure.AI.FormRecognizer.DocumentAnalysis;
using Azure;

using Burse.Data;
using Burse.Models.TemplatePDF;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;
using static QuestPDF.Helpers.Colors;
using System.Drawing;
using System.Drawing.Imaging;

namespace Burse.Controllers
{
    [ApiController]
    [Route("api/template")]
    public class TemplateController : ControllerBase
    {
        private readonly BurseDBContext _context;
        private readonly string endpoint = "https://formrecognizerlicenta.cognitiveservices.azure.com/";
        private readonly string apiKey = "42DilgHZfu2yj9b6pfNcjpAymzJLCQSrAK1p2C0qjQUOMiMkhvvnJQQJ99BDACYeBjFXJ3w3AAALACOGdG2j";

        public TemplateController(BurseDBContext context)
        {
            _context = context;
        }

        [HttpPost("SaveTemplate")]
        public async Task<IActionResult> SaveTemplate([FromBody] TemplateEntity template)
        {
            if (string.IsNullOrWhiteSpace(template.Name) || string.IsNullOrWhiteSpace(template.ElementsJson))
                return BadRequest("Template Name și ElementsJson sunt obligatorii.");

            template.CreatedAt = DateTime.UtcNow;
            _context.TemplateEntity.Add(template);
            await _context.SaveChangesAsync();

            return Ok(template);
        }

        [HttpGet("GetTemplates")]
        public async Task<IActionResult> GetTemplates()
        {
            var templates = await _context.TemplateEntity
                .OrderByDescending(t => t.CreatedAt)
                .ToListAsync();

            return Ok(templates);
        }

        [HttpGet("GetTemplate")]
        public async Task<IActionResult> GetTemplate([FromQuery] int id)
        {
            var template = await _context.TemplateEntity.FindAsync(id);

            if (template == null)
                return NotFound();

            return Ok(template);
        }
        [HttpDelete("DeleteTemplate{id}")]
        public async Task<IActionResult> DeleteTemplate(int id)
        {
            var template = await _context.TemplateEntity.FindAsync(id);
            if (template == null)
                return NotFound();

            _context.TemplateEntity.Remove(template);
            await _context.SaveChangesAsync();

            return Ok(new { message = "Template șters cu succes!" });
        }
        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateTemplate(int id, [FromBody] TemplateEntity updatedTemplate)
        {
            var existingTemplate = await _context.TemplateEntity.FindAsync(id);
            if (existingTemplate == null)
                return NotFound();

            existingTemplate.Name = updatedTemplate.Name;
            existingTemplate.ElementsJson = updatedTemplate.ElementsJson;
            existingTemplate.CreatedAt = DateTime.UtcNow; // sau păstrezi data veche dacă vrei

            await _context.SaveChangesAsync();

            return Ok(existingTemplate);
        }

        [HttpPost("AnalyzeDocument")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> AnalyzeDocument([FromForm] DocumentUploadRequest request)
        {
            if (request.File == null || request.File.Length == 0)
                return BadRequest("Fișierul este necesar.");

            var credential = new AzureKeyCredential(apiKey);
            var client = new DocumentAnalysisClient(new Uri(endpoint), credential);

            using var stream = request.File.OpenReadStream();
            var operation = await client.AnalyzeDocumentAsync(WaitUntil.Completed, request.ModelId, stream);
            var result = operation.Value;

            // Extragem datele normale (field-uri)
            var extractedData = result.Documents[0].Fields.ToDictionary(
                f => f.Key,
                f => (object)f.Value.Content
);
            // Extragem tabelele (inclusiv "Programe", dacă e tabel)
            var extractedTables = new List<List<string>>();

            foreach (var table in result.Tables)
            {
                foreach (var cell in table.Cells)
                {
                    // Asigurăm structura rândurilor
                    while (extractedTables.Count <= cell.RowIndex)
                    {
                        extractedTables.Add(new List<string>());
                    }

                    var row = extractedTables[cell.RowIndex];
                    while (row.Count <= cell.ColumnIndex)
                    {
                        row.Add(string.Empty);
                    }

                    row[cell.ColumnIndex] = cell.Content;
                }
            }

            extractedData["Tabel_Programe"] = extractedTables;


            return Ok(extractedData);

        }
        [HttpPost("UpscalingImage")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> UpscalingImage(IFormFile img)
        {
            if (img == null || img.Length == 0)
                return BadRequest("Fișierul este necesar.");

            Bitmap bitmap;
            using (var stream = img.OpenReadStream())
            {
                bitmap = new Bitmap(stream);
            }

            // 🔥 Fix important: Redimensionăm la multiplu de 4
            bitmap = ResizeToMultiple(bitmap, 2);

            var inputTensor = ImageToTensor(bitmap);

            var session = new InferenceSession("C:\\Licenta\\realesrgan-x4.onnx"); // asigură-te că e ONNX, nu fp16 onnx!
            var inputs = new List<NamedOnnxValue>
    {
        NamedOnnxValue.CreateFromTensor("input", inputTensor)
    };

            using var results = session.Run(inputs);
            var output = results.First().AsTensor<float>();
            var upscaled = TensorToBitmap(output);

            // 🔁 Redimensionare corectă și salvare
            var resized = ResizeImage(upscaled, 3000, 3000);
            using var ms = new MemoryStream();
            resized.Save(ms, ImageFormat.Png);
            ms.Seek(0, SeekOrigin.Begin);

            return File(ms.ToArray(), "image/png", "upscaled.png");
        }
        private Bitmap ResizeToMultiple(Bitmap bmp, int multiple)
        {
            int newWidth = bmp.Width - (bmp.Width % multiple);
            int newHeight = bmp.Height - (bmp.Height % multiple);

            var resized = new Bitmap(newWidth, newHeight);
            using (var g = Graphics.FromImage(resized))
            {
                g.DrawImage(bmp, 0, 0, newWidth, newHeight);
            }

            return resized;
        }

        Bitmap ResizeImage(Bitmap image, int maxWidth, int maxHeight)
        {
            var ratioX = (double)maxWidth / image.Width;
            var ratioY = (double)maxHeight / image.Height;
            var ratio = Math.Min(ratioX, ratioY);

            var newWidth = (int)(image.Width * ratio);
            var newHeight = (int)(image.Height * ratio);

            var resized = new Bitmap(newWidth, newHeight);
            using var g = Graphics.FromImage(resized);
            g.DrawImage(image, 0, 0, newWidth, newHeight);
            return resized;
        }

        private DenseTensor<float> ImageToTensor(Bitmap bmp)
        {
            int width = bmp.Width;
            int height = bmp.Height;
            var tensor = new DenseTensor<float>(new[] { 1, 3, height, width });

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    var pixel = bmp.GetPixel(x, y);
                    tensor[0, 0, y, x] = pixel.R / 255f;
                    tensor[0, 1, y, x] = pixel.G / 255f;
                    tensor[0, 2, y, x] = pixel.B / 255f;
                }
            }

            return tensor;
        }

        private Bitmap TensorToBitmap(Tensor<float> tensor)
        {
            int height = tensor.Dimensions[2];
            int width = tensor.Dimensions[3];
            var bmp = new Bitmap(width, height);

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    int r = ClampToByte(tensor[0, 0, y, x] * 255f);
                    int g = ClampToByte(tensor[0, 1, y, x] * 255f);
                    int b = ClampToByte(tensor[0, 2, y, x] * 255f);
                    bmp.SetPixel(x, y, Color.FromArgb(r, g, b));
                }
            }

            return bmp;
        }

        private int ClampToByte(float value)
        {
            return (int)Math.Max(0, Math.Min(255, value));
        }
        public class DocumentUploadRequest
        {
            public IFormFile File { get; set; }
            public string ModelId { get; set; }
        }

    }
}
