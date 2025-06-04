using System.Globalization;
using System.Text;

namespace Burse.Helpers
{
    public class AcronymGenerator
    {
        // Lista de cuvinte de legătură (stop words) care sunt ignorate
        private static readonly HashSet<string> StopWords = new HashSet<string>(
            new[] { "ȘI", "SI", "DE", "LA", "DIN", "CU", "ÎN","IN", "INVĂȚĂMÂNT","PENTRU","INVATAMANT","DUAL" },
            StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Generează un acronim pentru denumirea unui program.
        /// Exemplu: "Managementul energiei" => "ME",
        ///          "Automatică și informatică aplicată" => "AIA",
        ///          "Calculatoare" => "C",
        ///          "Calculatoare învățământ dual" => "C-DUAL"
        ///          "Inginerie energetică" => "IEN"
        /// </summary>
        public string GenerateAcronym(string programName, string an)
        {
            if (string.IsNullOrWhiteSpace(programName))
                return string.Empty;

            // Convertim la majuscule pentru consistență
            string upperProgram = AcronymGenerator.RemoveDiacritics(programName).ToUpperInvariant();

            // Dacă programul conține "ÎNVĂȚĂMÂNT DUAL", tratăm separat
            if (upperProgram.Contains("INVATAMANT DUAL"))
            {
                // Extragem partea dinaintea "ÎNVĂȚĂMÂNT DUAL"
                string coreProgram = upperProgram.Replace("INVATAMANT DUAL", "").Trim();
                string coreAcronym = GenerateAcronym(coreProgram, an); // Aplicăm algoritmul pe restul
                return $"{coreAcronym}-DUAL";
            }

            // Împărțim textul în cuvinte (se pot folosi spații, liniuțe etc.)
            var words = upperProgram.Split(new char[] { ' ', '-', ',' }, StringSplitOptions.RemoveEmptyEntries);

            // Eliminăm cuvintele de legătură (stop words)
            var filteredWords = words.Where(w => !StopWords.Contains(w.Trim())).ToList();

            if (filteredWords.Count == 0)
                return string.Empty;

            string acronym;

            // Dacă avem doar un cuvânt, returnăm prima literă
            if (filteredWords.Count == 1)
            {
                acronym = filteredWords[0].Substring(0, 1);
            }
            // Dacă avem două cuvinte, să luăm prima literă a fiecărui cuvânt,
            // dar introducem o excepție pentru "INGINERIE ENERGETICĂ"
            else if (filteredWords.Count == 2)
            {
                if (filteredWords[0] == "INGINERIE" && filteredWords[1].Contains("ENERGETIC"))
                {
                    acronym = "IEN";
                }
                else
                {
                    acronym = filteredWords[0].Substring(0, 1) + filteredWords[1].Substring(0, 1);
                }
            }
            // Pentru trei sau mai multe cuvinte, returnăm prima literă a fiecărui cuvânt
            else
            {
                acronym = string.Join("", filteredWords.Select(w => w.Substring(0, 1)));
            }

            return $"{acronym} ({an})";
        }
        public static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);

                if (unicodeCategory != UnicodeCategory.NonSpacingMark &&
                    !char.IsPunctuation(c) &&
                    !char.IsSymbol(c))
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }


    }
}
