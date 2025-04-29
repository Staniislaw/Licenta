using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace Burse.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelUploadController : ControllerBase
    {
        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(List<IFormFile> files)
        {
            if (files == null || files.Count == 0)
                return BadRequest("No files uploaded.");

            var results = new List<object>();

            foreach (var file in files)
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                using var workbook = new XLWorkbook(stream);
                var worksheet = workbook.Worksheets.First();

                var tableData = new List<Dictionary<string, object>>();
                
                var headers = new List<string>();
                foreach (var cell in worksheet.Row(1).CellsUsed())
                {
                    headers.Add(cell.GetValue<string>());
                }

                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    var rowData = new Dictionary<string, object>();
                    int i = 0;
                    foreach (var cell in row.Cells(1, headers.Count))
                    {
                        rowData[headers[i]] = cell.Value;
                        i++;
                    }
                    tableData.Add(rowData);
                }

                results.Add(new
                {
                    FileName = file.FileName,
                    Data = tableData
                });
            }

            return Ok(results);
        }
    }
}
