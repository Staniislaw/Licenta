using Burse.Data;
using Burse.Models;
using Burse.Services.Abstractions;

using ClosedXML.Excel;

using Microsoft.EntityFrameworkCore;

namespace Burse.Services
{
    public class StudentService : IStudentService
    {
        private readonly BurseDBContext _context;

        public StudentService(BurseDBContext context)
        {
            _context = context;
        }

        public async Task<List<StudentRecord>> GetAllAsync()
        {
            return await _context.StudentRecord
                .Include(s => s.FondBurseMeritRepartizat)
                 .Include(s => s.IstoricBursa)
                .ToListAsync();
        }
        public async Task<StudentRecord> UpdateBursaAsync(int id, string bursa)
        {
            var student = await _context.StudentRecord
                .Include(s => s.FondBurseMeritRepartizat)
                .Include(s => s.IstoricBursa)
                .FirstOrDefaultAsync(s => s.Id == id);

            if (student == null)
                return null;

            var actiune = "Modificare bursa";
            if(string.IsNullOrEmpty(bursa))
            {
                bursa = null;
            }
            student.Bursa = bursa;

            var historyEntry = new BursaIstoric
            {
                StudentRecordId = student.Id,
                TipBursa = bursa,
                Motiv = $"Schimbare bursă prin interfață",
                Actiune = actiune,
                Etapa = null,               
                Suma = 0m,              
                Comentarii = "Modificat din interfață",
                DataModificare = DateTime.UtcNow
            };
            student.IstoricBursa.Add(historyEntry);

            await _context.SaveChangesAsync();

            return student;
        }

        public async Task<byte[]> ExportStudentiExcelAsync()
        {
            var studenti = await GetAllAsync(); // sau repository/metoda ta
            var studentiCuBursa = studenti
                .Where(s => !string.IsNullOrEmpty(s.Bursa) && s.Bursa != "NU") // adaptează dacă ai alte valori
                .ToList();

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Studenți");

            // Antet
            var headers = new[]
                {
                    "Nr. crt.", "Emplid", "CNP", "Nume Student", "Țară Cetățenie",
                    "An", "Media", "Punctaj An", "CO", "RO", "TC", "TR",
                    "Sursa de finanțare", "Domeniu", "Bursa", "Suma Bursă"
                };
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = headers[i];
                worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                worksheet.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(1, i + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            }

            // Conținut
            int row = 2;
            int nrCrt = 1;

            foreach (var s in studentiCuBursa)
            {
                worksheet.Cell(row, 1).Value = nrCrt++;
                worksheet.Cell(row, 2).Value = s.Emplid;
                worksheet.Cell(row, 3).Value = s.CNP;
                worksheet.Cell(row, 4).Value = s.NumeStudent;
                worksheet.Cell(row, 5).Value = "RO"; // sau s.Tara, dacă ai câmpul
                worksheet.Cell(row, 6).Value = s.An + 1;
                worksheet.Cell(row, 7).Value = s.Media;
                worksheet.Cell(row, 8).Value = s.PunctajAn;
                worksheet.Cell(row, 9).Value = s.CO;
                worksheet.Cell(row, 10).Value = s.RO;
                worksheet.Cell(row, 11).Value = s.TC;
                worksheet.Cell(row, 12).Value = s.TR;
                worksheet.Cell(row, 13).Value = s.SursaFinantare; // sau logica ta: s.Bursa == "DA" ? "BUGET" : "TAXĂ"
                worksheet.Cell(row, 14).Value = s.FondBurseMeritRepartizat?.domeniu ?? ""; // nou
                worksheet.Cell(row, 15).Value = s.Bursa ?? ""; // nou
                worksheet.Cell(row, 16).Value = s.SumaBursa.ToString("0.00") ?? ""; // nou


                // Stilizare: Bordură pt fiecare celulă din rând
                for (int col = 1; col <= headers.Length; col++)
                {
                    var cell = worksheet.Cell(row, col);
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                row++;
            }

            // Auto-size coloane
            worksheet.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }
    }
}
