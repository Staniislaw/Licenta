using System.Collections.Generic;
using System.IO;

using Burse.Models;

using ExcelDataReader;

public class FondBurseExcelReader
{
    public List<FondBurse> ReadFondBurseFromExcel(Stream stream)
    {
        var fonduriBurse = new List<FondBurse>();

        // Configurează ExcelDataReader să suporte fișiere de tip .xls și .xlsx
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            while (reader.Read())
            {
                // Sari peste rândul de antet
                if (reader.Depth == 0) continue;

                // Creează un nou obiect FondBurse
                var fondBurse = new FondBurse
                {
                    CategorieBurse = reader.GetString(0) // Coloana 1
                };

                // Validare și conversie pentru coloana Valorea Lunara
                string valoareText = reader.GetValue(1)?.ToString(); // Asigură-te că obții text
                if (decimal.TryParse(valoareText, out decimal valoareDecimal))
                {
                    fondBurse.ValoreaLunara = valoareDecimal;
                }
                else
                {
                    // Valoarea nu poate fi convertită, setează 0
                    fondBurse.ValoreaLunara = 0;
                }

                // Verifică dacă rândul este gol (CategorieBurse gol sau ValoreaLunara 0)
                if (!string.IsNullOrWhiteSpace(fondBurse.CategorieBurse) && fondBurse.ValoreaLunara != 0)
                {
                    // Adaugă obiectul în listă doar dacă are date valabile
                    fonduriBurse.Add(fondBurse);
                }
            }
        }

        return fonduriBurse;
    }


}
