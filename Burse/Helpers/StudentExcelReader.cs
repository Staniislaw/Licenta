using Burse.Models;
using DocumentFormat.OpenXml.Drawing;

using ExcelDataReader;
using System.Data;

namespace Burse.Helpers
{
    public class StudentExcelReader
    {
        public Dictionary<string, List<StudentRecord>> ReadStudentRecordsFromExcel(string filePath)
        {
            var studentRecordsByDomain = new Dictionary<string, List<StudentRecord>>();
            string domeniu = ""; // Salvează domeniul o singură dată
            var generator = new AcronymGenerator();
            // Activează suportul pentru encoding-ul necesar fișierelor Excel
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    bool isHeaderPassed = false; // Folosit pentru a ști când încep datele studenților

                    bool isTableStarted = false;
                    int anStudiu = 0; // Inițializăm anul studiului

                    while (reader.Read())
                    {
                        int totalColumns = reader.FieldCount;
                        string firstCellValue = reader.GetValue(0)?.ToString()?.Trim();

                        // 🔍 Debugging: Afișează fiecare rând citit
                        Console.WriteLine($"Row {reader.Depth}: {totalColumns} columns detected");
                        for (int i = 0; i < totalColumns; i++)
                        {
                            Console.Write($"{reader.GetValue(i)?.ToString()} | ");
                        }
                        Console.WriteLine(); // Linie nouă pentru fiecare rând

                        // ✅ Detectează domeniul corect
                        if (firstCellValue == "Domeniul:" && totalColumns > 1)
                        {
                            domeniu = reader.GetValue(1)?.ToString()?.Trim() ?? "";
                            continue;
                        }

                        // ✅ Detectează anul studiului din "An școlar:"
                        if (firstCellValue == "An școlar:" && totalColumns > 2)
                        {
                            bool anValid = int.TryParse(reader.GetValue(2)?.ToString(), out anStudiu);
                            if (!anValid) anStudiu = 0; // Dacă nu este valid, setăm 0
                            continue;
                        }

                        // ❌ Dacă tabelul nu a început, ignoră rândurile
                        if (!isTableStarted && firstCellValue == "Nr. crt.")
                        {
                            isTableStarted = true;
                            continue;
                        }

                        // ❌ Dacă tabelul nu a început sau domeniul nu este setat, continuăm
                        if (!isTableStarted || string.IsNullOrEmpty(domeniu)) continue;
             
                        string domeniuAcronim = generator.GenerateAcronym(domeniu, anStudiu.ToString());
                        // 🔹 Adaugăm domeniul în dictionary
                        if (!studentRecordsByDomain.ContainsKey(domeniuAcronim))
                        {
                            studentRecordsByDomain[domeniuAcronim] = new List<StudentRecord>();
                        }

                        Console.WriteLine($"✅ Domeniu găsit: {domeniu} -> {domeniuAcronim}");

                        // ✅ Procesăm doar rândurile cu studenți (prima coloană trebuie să fie un număr valid)
                        if (int.TryParse(firstCellValue, out int nrCrt))
                        {
                            string emplid = totalColumns > 1 ? reader.GetValue(1)?.ToString()?.Trim() ?? "" : "";
                            string cnp = totalColumns > 2 ? reader.GetValue(2)?.ToString()?.Trim() ?? "" : "";
                            string numeStudent = totalColumns > 3 ? reader.GetValue(3)?.ToString()?.Trim() ?? "[NECUNOSCUT]" : "";
                            string taraCetatenie = totalColumns > 4 ? reader.GetValue(4)?.ToString()?.Trim() ?? "" : "";

                            int an = (totalColumns > 5 && int.TryParse(reader.GetValue(5)?.ToString(), out int tempAn)) ? tempAn : 0;
                            decimal media = (totalColumns > 6 && decimal.TryParse(reader.GetValue(6)?.ToString(), out decimal tempMedia)) ? tempMedia : 0;
                            int punctajAn = (totalColumns > 7 && int.TryParse(reader.GetValue(7)?.ToString(), out int tempPunctaj)) ? tempPunctaj : 0;
                            int co = (totalColumns > 8 && int.TryParse(reader.GetValue(8)?.ToString(), out int tempCo)) ? tempCo : 0;
                            int ro = (totalColumns > 9 && int.TryParse(reader.GetValue(9)?.ToString(), out int tempRo)) ? tempRo : 0;
                            int tc = (totalColumns > 10 && int.TryParse(reader.GetValue(10)?.ToString(), out int tempTc)) ? tempTc : 0;
                            int tr = (totalColumns > 11 && int.TryParse(reader.GetValue(11)?.ToString(), out int tempTr)) ? tempTr : 0;

                            // 🔹 Normalizează "Sursa de finanțare" (elimină spații și caractere speciale)
                            string sursaFinantare = totalColumns > 12 ? reader.GetValue(12)?.ToString()?.Trim().Replace("\n", "").Replace("\r", "") ?? "" : "";

                            // ✅ Adaugă studentul în listă
                            var studentRecord = new StudentRecord
                            {
                                NrCrt = nrCrt,
                                Emplid = emplid,
                                CNP = cnp,
                                NumeStudent = numeStudent,
                                TaraCetatenie = taraCetatenie,
                                An = an,
                                Media = media,
                                PunctajAn = punctajAn,
                                CO = co,
                                RO = ro,
                                TC = tc,
                                TR = tr,
                                SursaFinantare = sursaFinantare
                            };

                            studentRecordsByDomain[domeniuAcronim].Add(studentRecord);
                        }
                    }

                }
            }

            return studentRecordsByDomain;
        }

    }
}
