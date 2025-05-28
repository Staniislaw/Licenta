using Microsoft.AspNetCore.Mvc;

using System.Text.RegularExpressions;

namespace Burse.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class LogsController : ControllerBase
    {
        private readonly string _logsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Logs");

        // Modified route for clarity
        [HttpGet("GetLogCategories")] // Changed from "GetLogCategories" for consistency with Angular
        public IActionResult GetLogCategories()
        {
            if (!Directory.Exists(_logsFolder))
            {
                return Ok(new List<string>());
            }

            // NOUA Expresie Regulată: Permite litere, cifre și cratime în partea 'type'
            var filePattern = new Regex(@"^(?<type>[a-zA-Z0-9-]+)-(?<date>\d{8})\.txt$", RegexOptions.IgnoreCase);

            var categories = Directory.GetFiles(_logsFolder, "*.txt")
                .Select(f => Path.GetFileName(f))
                .Where(fileName => filePattern.IsMatch(fileName)) // Acest Where e crucial
                .Select(fileName => filePattern.Match(fileName).Groups["type"].Value)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(c => c)
                .ToList();

            var userFriendlyCategories = categories.Select(c =>
            {
                switch (c.ToLowerInvariant())
                {
                    case "errors": return "Errors";
                    case "students": return "studenti aceeasi bursa";
                    case "formatii": return "Formatii";
                    case "excel-import": return "Excel Import";
                    case "excel": return "Excel Import";
                    // Important: "students-excels" aici trebuie să se potrivească EXACT cu ce extrage regex-ul
                    case "students-excels": return "Students Excels"; // Denumirea user-friendly pentru UI
                    default: return char.ToUpper(c[0]) + c.Substring(1);
                }
            }).ToList();

            return Ok(userFriendlyCategories);
        }

        // Endpoint 2: Obținerea datelor disponibile pentru o categorie de log specifică
        // GET /api/Logs/availabledates?category=Errors
        [HttpGet("availabledates")]
        public IActionResult GetAvailableLogDates([FromQuery] string logType) // Numele parametrului schimbat AICI
        {
            if (string.IsNullOrWhiteSpace(logType)) // Verificăm noul nume al parametrului
            {
                return BadRequest("LogType parameter is required.");
            }

            if (!Directory.Exists(_logsFolder))
            {
                return Ok(new List<string>());
            }

            // Aici, folosește `logType` în loc de `category`
            string filePrefix = logType.ToLowerInvariant(); // Folosim `logType` direct
            switch (filePrefix)
            {
                case "errors": filePrefix = "errors"; break;
                case "studenti aceeasi bursa": filePrefix = "students"; break;
                case "formatii": filePrefix = "formatii"; break;
                case "excel import": filePrefix = "excel-import"; break;
                case "students excels": filePrefix = "students-excels"; break;
                case "excel": filePrefix = "excel"; break;
                default:
                    return BadRequest($"Unknown log category: {logType}");
            }

            // ... restul logicii din această metodă rămâne la fel, folosind 'filePrefix' ...
            var filePattern = new Regex($@"^{filePrefix}-(?<date>\d{{8}})\.txt$", RegexOptions.IgnoreCase);

            var dates = Directory.GetFiles(_logsFolder, $"{filePrefix}-*.txt")
                .Select(f => Path.GetFileName(f))
                .Where(fileName => filePattern.IsMatch(fileName))
                .Select(fileName =>
                {
                    var match = filePattern.Match(fileName);
                    if (match.Success)
                    {
                        string datePart = match.Groups["date"].Value;
                        if (datePart.Length == 8 && DateTime.TryParseExact(datePart, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                        {
                            return parsedDate.ToString("yyyy-MM-dd");
                        }
                    }
                    return null;
                })
                .Where(d => d != null)
                .Distinct()
                .OrderByDescending(d => d)
                .ToList();

            return Ok(dates);
        }


        // Endpoint 3: Obținerea conținutului unui log specific (Tip + Dată)
        // GET /api/Logs/content?category=Errors&date=2025-05-28
        [HttpGet("content")]
        public async Task<IActionResult> GetLogContent([FromQuery] string logType, [FromQuery] string date)
        {
            if (string.IsNullOrWhiteSpace(logType) || string.IsNullOrWhiteSpace(date))
            {
                return BadRequest("Category and date parameters are required.");
            }

            // Mapează categoria user-friendly înapoi la prefixul din numele fișierului
            string filePrefix = logType.ToLowerInvariant();
            switch (filePrefix)
            {
                case "errors": filePrefix = "errors"; break;
                case "studenti aceeasi bursa": filePrefix = "students"; break;
                case "student": filePrefix = "students"; break; // pentru cazul in care AppLogger e singular
                case "formatii": filePrefix = "formatii"; break;
                case "excel import": filePrefix = "excel-import"; break;
                case "students excels": filePrefix = "students-excels"; break;
                case "excel": filePrefix = "excel"; break;
                default:
                    return BadRequest($"Unknown log category: {logType}");
            } 

            // Ajustează formatul datei la cel din numele fișierului (e.g., "2025-05-28" -> "20250528")
            string fileDatePart = date.Replace("-", "");

            var filePath = Path.Combine(_logsFolder, $"{filePrefix}-{fileDatePart}.txt");

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound($"Log file for category '{logType}' on date '{date}' not found at '{filePath}'.");
            }
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(stream))
            {
                var content = await reader.ReadToEndAsync();
                var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                return Ok(lines);
            }
        }
    }
}
