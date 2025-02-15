using Burse.Models;

using ExcelDataReader;

public class FormatiiStudiiFromExcel
{
    public List<FormatiiStudii> ReadFormatiiStudiiFromExcel(string filePath)
    {
        var formatiiStudii = new List<FormatiiStudii>();

        // Configurează ExcelDataReader să suporte fișiere de tip .xls și .xlsx
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                int idCounter = 1;
                bool citimDate = false; // Flag pentru a ști când să citim datele
                bool sarRandAntet = false; // Pentru a sări peste antet după ce detectăm titlul tabelului
                string[] ultimaInregistrare = new string[16]; // Salvăm ultima înregistrare pentru a completa datele lipsă
                string ultimulProgramDeStudiu = "";
                string facultatea = "";
                // Iterăm rândurile
                while (reader.Read())
                {
                    var firstCell = reader.GetValue(0)?.ToString()?.Trim();

                    // Identificăm titlul tabelului
                    if (string.Equals(firstCell, "STUDII UNIVERSITARE DE LICENȚĂ, ÎNVĂȚĂMÂNT CU FRECVENȚĂ", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(firstCell, "STUDII UNIVERSITARE DE MASTERAT, ÎNVĂȚĂMÂNT CU FRECVENȚĂ", StringComparison.OrdinalIgnoreCase))
                    {
                        citimDate = true; // Activăm citirea datelor
                        sarRandAntet = true; // Următorul rând trebuie să fie antet, deci îl ignorăm
                        ultimaInregistrare = new string[16]; // Resetăm ultima înregistrare
                        continue;
                    }

                    // Dacă am întâlnit un alt tip de tabel sau informații irelevante, oprim citirea
                    if (citimDate && string.IsNullOrWhiteSpace(firstCell))
                    {
                        // Ignorăm rândurile complet goale
                        if (IsRowEmpty(reader))
                        {
                            citimDate = false;
                            continue;
                        }
                    }

                    // Ignorăm rândul de antet
                    if (sarRandAntet)
                    {
                        sarRandAntet = false;
                        continue;
                    }
                    // Citim datele relevante doar dacă suntem în secțiunea tabelului
                    if (citimDate)
                    {
                        // Procesăm rândul și completăm valorile lipsă
                        var valori = new string[16];
                        for (int i = 0; i < valori.Length; i++)
                        {
                            valori[i] = reader.GetValue(i)?.ToString()?.Trim();

                            // Dacă valoarea curentă este `null` sau goală
                            if (string.IsNullOrWhiteSpace(valori[i]))
                            {
                                // Setăm 0 în loc de valoarea anterioară
                                valori[i] = "0";
                            }
                        }
                        if (valori[1]!="0")
                        {
                            ultimulProgramDeStudiu=valori[1];
                        }
                        if (valori[0]!="0")
                        {
                            facultatea = valori[0];
                        }
                        // Creez un obiect nou FormatiiStudii
                        string an = ConvertesteAnulInString(valori[2]);
                        var formatie = new FormatiiStudii
                        {
                            id = 0,
                            Facultatea = facultatea, // Facultatea
                            ProgramDeStudiu = ultimulProgramDeStudiu, // Program de studiu
                            An = an, // An
                            FaraTaxaRomani = valori[6], // Fără taxă români
                            FaraTaxaRp = valori[7], // Fără taxă RP
                            FaraTaxaUECEE = valori[8], // Fără taxă UE CEE
                            CuTaxaRomani = valori[9], // Cu taxă români
                            ElibiliB = valori[10], // Eligibili B
                            CuTaxaRM = valori[11], // Cu taxă RM
                            RMEligibil = valori[12], // RM Eligibil
                            CuTaxaUECEE = valori[13], // Cu taxă UE CEE
                            BursieriAIStatuluiRoman = valori[14], // Bursieri AI statului român
                            CPV = valori[15] // CPV
                        };

                        // Adaugă doar dacă există valori relevante în "Facultatea"
                        if (!string.IsNullOrWhiteSpace(formatie.Facultatea))
                        {
                            formatiiStudii.Add(formatie);
                        }
                    }
                }
            }
        }

        return formatiiStudii;
    }
    private string ConvertesteAnulInString(string anInLitere)
    {
        switch (anInLitere.Trim().ToUpper())
        {
            case "I": return "1";  // Anul 1
            case "II": return "2"; // Anul 2
            case "III": return "3"; // Anul 3
            case "IV": return "4";  // Anul 4
            default: return "An invalid"; // Dacă valoarea nu se potrivește (opțional, poți să alegi ce să returnezi în acest caz)
        }
    }
    // Metodă pentru a verifica dacă un rând este gol
    private bool IsRowEmpty(IExcelDataReader reader)
    {
        for (int i = 0; i < reader.FieldCount; i++)
        {
            if (!string.IsNullOrWhiteSpace(reader.GetValue(i)?.ToString()))
            {
                return false; // Rândul are cel puțin o valoare validă
            }
        }
        return true;
    }
}
