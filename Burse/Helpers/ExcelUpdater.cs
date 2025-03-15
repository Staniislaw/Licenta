using Burse.Models;
using OfficeOpenXml;


public class ExcelUpdater
{
    public static void UpdateScholarshipCounts(string filePath, List<StudentScholarshipData> studentiClasificati)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];

            Console.WriteLine($"📄 Numele foii: {worksheet.Name}");

            // 🟢 Detectăm automat unde începe tabelul
            int headerRow = FindHeaderRow(worksheet, "Program de studiu");


            Console.WriteLine($"🔎 Rândul antetului detectat: {headerRow}");
            Console.WriteLine("📌 Conținutul rândului de antet:");

            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                string columnText = worksheet.Cells[headerRow, col].Text.Trim();
                string columnValue = worksheet.Cells[headerRow, col].Value?.ToString().Trim() ?? "";

                Console.WriteLine($"Coloana {col}: Text='{columnText}', Value='{columnValue}'");
            }


            if (headerRow == -1)
            {
                Console.WriteLine("⚠️ Nu s-a găsit rândul antetului pentru tabel.");
                return;
            }

            // 🟢 Găsim indexul coloanelor, inclusiv în celule fuzionate
            int programStudiuCol = FindColumnIndex(worksheet, "Program de studiu", headerRow);
            int bp1Col = FindColumnIndex(worksheet, "BM1", headerRow);
            if (bp1Col == -1) bp1Col = FindColumnIndex(worksheet, "B.Perf.1", headerRow);
            if (bp1Col == -1) bp1Col = FindColumnIndex(worksheet, "BM1 (B.Perf.1)", headerRow);

            int bp2Col = FindColumnIndex(worksheet, "BM2", headerRow);
            if (bp2Col == -1) bp2Col = FindColumnIndex(worksheet, "B.Perf.2", headerRow);
            if (bp2Col == -1) bp2Col = FindColumnIndex(worksheet, "BM2 (B.Perf.2)", headerRow);


            if (programStudiuCol == -1 || bp1Col == -1 || bp2Col == -1)
            {
                Console.WriteLine("⚠️ Nu s-au găsit toate coloanele necesare în foaia selectată.");
                return;
            }

            int lastRow = worksheet.Dimension.End.Row;

            // 🟢 Actualizăm datele în fișier
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string domeniu = worksheet.Cells[row, programStudiuCol].Text.Trim();
                var entry = studentiClasificati.FirstOrDefault(s => s.Domeniu == domeniu);

                if (entry != null)
                {
                    worksheet.Cells[row, bp1Col].Value = entry.BP1Count;
                    worksheet.Cells[row, bp2Col].Value = entry.BP2Count;
                }
            }

            package.Save();
            Console.WriteLine("✅ Datele au fost actualizate în fișierul Excel.");
        }
    }


    private static int FindHeaderRow(ExcelWorksheet worksheet, string headerName)
    {
        int lastRow = worksheet.Dimension.End.Row;

        for (int row = 1; row <= lastRow; row++) // Căutăm până la finalul fișierului
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                string cellText = worksheet.Cells[row, col].Text.Trim().ToLower();
                string cellValue = worksheet.Cells[row, col].Value?.ToString().Trim().ToLower() ?? "";

                if (cellText.Contains(headerName.ToLower()) || cellValue.Contains(headerName.ToLower()))
                {
                    Console.WriteLine($"✅ Antet găsit pe rândul {row} la coloana {col}");
                    return row;
                }
            }
        }
        return -1; // Nu am găsit antetul
    }


    private static int FindColumnIndex(ExcelWorksheet worksheet, string columnName, int headerRow)
    {
        int lastCol = worksheet.Dimension.End.Column;
        columnName = columnName.ToLower().Trim();

        // 🔎 Căutăm în rândurile 16, 17 și 18
        for (int row = headerRow; row <= headerRow + 2; row++)
        {
            for (int col = 1; col <= lastCol; col++)
            {
                string cellText = worksheet.Cells[row, col].Text.Trim().ToLower();
                string cellValue = worksheet.Cells[row, col].Value?.ToString().Trim().ToLower() ?? "";

                if (cellText.Contains(columnName) || cellValue.Contains(columnName))
                {
                    Console.WriteLine($"✅ Coloana '{columnName}' găsită pe index {col} în rândul {row}");
                    return col;
                }
            }
        }

        Console.WriteLine($"❌ Coloana '{columnName}' nu a fost găsită.");
        return -1;
    }
    


}
