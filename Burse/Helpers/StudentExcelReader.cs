using Burse.Models;
using DocumentFormat.OpenXml.Drawing;

using ExcelDataReader;
using System.Data;

namespace Burse.Helpers
{
    public class StudentExcelReader
    {
        public Dictionary<string, List<StudentRecord>> ReadStudentRecordsFromExcel(Stream stream, string fisier)
        {
            var studentRecordsByDomain = new Dictionary<string, List<StudentRecord>>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            string fileName = System.IO.Path.GetFileNameWithoutExtension(fisier);
            var columnMappings = LoadColumnMappingsFromDatabase(); // Încărcăm mapping-ul
            
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    string formattedSheetName = reader.Name;
                    string domeniu = $"{fileName.ToUpper()} ({formattedSheetName})";

                    // Detectăm "Xcdual" și transformăm în "C (X)-DUAL"
                    var matchDual = System.Text.RegularExpressions.Regex.Match(formattedSheetName, @"^(\d+)\w*dual$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (matchDual.Success)
                    {
                        formattedSheetName = $"{fileName} ({matchDual.Groups[1].Value})-DUAL";
                        domeniu = formattedSheetName;
                    }
                    else
                    {
                        // Detectăm orice format de tip "1scc", "2rcc", "2sc" și îl transformăm în "X (Y)", excluzând "dual"
                        var matchGeneric = System.Text.RegularExpressions.Regex.Match(formattedSheetName, @"^(\d+)([a-zA-Z]+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        if (matchGeneric.Success && !matchGeneric.Groups[2].Value.ToLower().Contains("dual"))
                        {
                            formattedSheetName = $"{matchGeneric.Groups[2].Value.ToUpper()} ({matchGeneric.Groups[1].Value})";
                            domeniu = formattedSheetName;
                        }
                    }
                       
                    if (fileName.ToUpper() == "IETTI")
                    {
                        if (formattedSheetName == "1")
                        {
                            domeniu = "IETTI (1)";
                        }
                        else if (formattedSheetName == "2")
                        {
                            domeniu = "IETTI (2)";
                        }
                        else if (formattedSheetName == "3")
                        {
                            domeniu = "RST (3)";
                        }
                        else if (formattedSheetName == "4")
                        {
                            domeniu = "RST (4)";
                        }
                    }

                    bool isTableStarted = false;
                    Dictionary<string, int> columnMapping = new Dictionary<string, int>();

                    Console.WriteLine($"📄 Citim foaia: {reader.Name}");

                    while (reader.Read())
                    {
                        int totalColumns = reader.FieldCount;
                        string firstCellValue = reader.GetValue(0)?.ToString()?.Trim() ?? "";

                        if (!isTableStarted && firstCellValue.ToLower().Contains("nr. crt"))
                        {
                            isTableStarted = true;
                            for (int i = 0; i < totalColumns; i++)
                            {
                                string colName = reader.GetValue(i)?.ToString()?.Trim().ToLower() ?? "";
                                if (!string.IsNullOrEmpty(colName))
                                {
                                    columnMapping[colName] = i;
                                }
                            }
                            Console.WriteLine($"📌 Antet tabel detectat, mapăm coloanele...");
                            continue;
                        }

                        if (!isTableStarted || columnMapping.Count == 0) continue;

                        if (!studentRecordsByDomain.ContainsKey(domeniu))
                        {
                            studentRecordsByDomain[domeniu] = new List<StudentRecord>();
                        }

                        if (int.TryParse(firstCellValue, out int nrCrt))
                        {
                            var student = new StudentRecord
                            {
                                NrCrt = nrCrt,
                                Emplid = GetColumnValue(reader, columnMapping, columnMappings["Emplid"]),
                                CNP = GetColumnValue(reader, columnMapping, columnMappings["CNP"]),
                                NumeStudent = GetColumnValue(reader, columnMapping, columnMappings["NumeStudent"]),
                                TaraCetatenie = GetColumnValue(reader, columnMapping, columnMappings["TaraCetatenie"]),
                                An = GetColumnValueAsInt(reader, columnMapping, columnMappings["An"]),
                                Media = GetColumnValueAsDecimal(reader, columnMapping, columnMappings["Media"]),
                                PunctajAn = GetColumnValueAsInt(reader, columnMapping, columnMappings["PunctajAn"]),
                                CO = GetColumnValueAsInt(reader, columnMapping, columnMappings["CO"]),
                                RO = GetColumnValueAsInt(reader, columnMapping, columnMappings["RO"]),
                                TC = GetColumnValueAsInt(reader, columnMapping, columnMappings["TC"]),
                                TR = GetColumnValueAsInt(reader, columnMapping, columnMappings["TR"]),
                                SursaFinantare = GetColumnValue(reader, columnMapping, columnMappings["SursaFinantare"])
                            };

                            Console.WriteLine($"👨‍🎓 Student detectat: {student.NumeStudent} - Media: {student.Media}");
                            studentRecordsByDomain[domeniu].Add(student);
                        }
                    }
                } while (reader.NextResult());
            }

            return studentRecordsByDomain;
        }

        private static string GetColumnValue(IDataReader reader, Dictionary<string, int> columnMapping, List<string> keys)
        {
            foreach (var key in keys)
            {
                if (columnMapping.TryGetValue(key.ToLower(), out var index))
                {
                    var value = reader.GetValue(index)?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(value))
                    {
                        return value;
                    }
                }
            }
            return "";
        }


        private static int GetColumnValueAsInt(IDataReader reader, Dictionary<string, int> columnMapping, List<string> keys)
        {
            foreach (var key in keys)
            {
                if (columnMapping.ContainsKey(key.ToLower()) && int.TryParse(reader.GetValue(columnMapping[key.ToLower()])?.ToString(), out int value))
                {
                    return value;
                }
            }
            return 0;
        }

        private static decimal GetColumnValueAsDecimal(IDataReader reader, Dictionary<string, int> columnMapping, List<string> keys)
        {
            foreach (var key in keys)
            {
                if (columnMapping.ContainsKey(key.ToLower()) && decimal.TryParse(reader.GetValue(columnMapping[key.ToLower()])?.ToString(), out decimal value))
                {
                    return value;
                }
            }
            return 0;
        }

        public static Dictionary<string, List<string>> LoadColumnMappingsFromDatabase()
        {
            var columnMappings = new Dictionary<string, List<string>>();
            columnMappings["Emplid"] = new List<string> { "Nr. matricol", "Emplid" };
            columnMappings["CNP"] = new List<string> { "CNP" };
            columnMappings["NumeStudent"] = new List<string> { "Nume Student", "Nume și Prenume" };
            columnMappings["TaraCetatenie"] = new List<string> { "Țară Cetățenie" };
            columnMappings["An"] = new List<string> { "An" };
            columnMappings["Media"] = new List<string> { "Media", "Media de admitere", };
            columnMappings["PunctajAn"] = new List<string> { "Punctaj An" };
            columnMappings["CO"] = new List<string> { "CO" };
            columnMappings["RO"] = new List<string> { "RO" };
            columnMappings["TC"] = new List<string> { "TC" };
            columnMappings["TR"] = new List<string> { "TR" };
            columnMappings["SursaFinantare"] = new List<string> { "Sursa Finanțare" };

            return columnMappings;
        }


        public Dictionary<string, List<StudentRecord>> ReadStudentRecordsFromExcel2(string filePath)
        {
            var studentRecordsByDomain = new Dictionary<string, List<StudentRecord>>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do  // 🔹 Citim fiecare foaie din fișier
                    {
                        string domeniu = "";
                        int anStudiu = 0;
                        bool isTableStarted = false;

                        Console.WriteLine($"📄 Citim foaia: {reader.Name}");

                        while (reader.Read()) // 🔹 Citim fiecare rând
                        {
                            int totalColumns = reader.FieldCount;
                            string firstCellValue = reader.GetValue(0)?.ToString()?.Trim() ?? "";

                            // 🔍 Debugging
                            Console.WriteLine($"Row {reader.Depth}: {totalColumns} columns detected");
                            for (int i = 0; i < totalColumns; i++)
                            {
                                Console.Write($"{reader.GetValue(i)?.ToString()} | ");
                            }
                            Console.WriteLine();

                            // ✅ Detectează domeniul
                            if (firstCellValue == "Domeniul:" && totalColumns > 1)
                            {
                                domeniu = reader.GetValue(1)?.ToString()?.Trim() ?? "";
                                Console.WriteLine($"🔹 Domeniu detectat: {domeniu}");
                                continue;
                            }

                            // ✅ Detectează anul studiului
                            if (firstCellValue == "An școlar:" && totalColumns > 2)
                            {
                                bool anValid = int.TryParse(reader.GetValue(2)?.ToString(), out anStudiu);
                                anStudiu = anValid ? anStudiu : 0;
                                Console.WriteLine($"🔹 An școlar detectat: {anStudiu}");
                                continue;
                            }

                            // ❌ Ignorăm rândurile până începe tabelul
                            if (!isTableStarted && firstCellValue == "Nr. crt.")
                            {
                                isTableStarted = true;
                                continue;
                            }

                            // ❌ Ignorăm rândurile dacă nu avem domeniu
                            if (!isTableStarted || string.IsNullOrEmpty(domeniu)) continue;

                            // 🔹 Generăm acronimul pentru dicționar
                            string domeniuAcronim = $"{domeniu}_{reader.Name}_{anStudiu}".Trim();

                            // 🛑 Evităm cheile goale
                            if (string.IsNullOrEmpty(domeniuAcronim) || domeniuAcronim == "_0")
                            {
                                Console.WriteLine($"⚠️ Domeniu invalid pentru foaia {reader.Name}, continuăm...");
                                continue;
                            }

                            // ✅ Asigurăm că cheia există în dicționar
                            if (!studentRecordsByDomain.ContainsKey(domeniuAcronim))
                            {
                                studentRecordsByDomain[domeniuAcronim] = new List<StudentRecord>();
                            }

                            Console.WriteLine($"✅ Domeniu procesat: {domeniuAcronim}");

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

                                // 🔹 Normalizează "Sursa de finanțare"
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
                    } while (reader.NextResult()); // 🔹 Trecem la următoarea foaie din Excel
                }
            }

            return studentRecordsByDomain;
        }

    }
}
