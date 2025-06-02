using Burse.Models;

namespace Burse.Helpers
{
    public class StudentScholarshipComparer : IComparer<StudentRecord>
    {
        private readonly FondBurseMeritRepartizat _fondBurseMeritRepartizat;

        public StudentScholarshipComparer(FondBurseMeritRepartizat fondBurseMeritRepartizat)
        {
            _fondBurseMeritRepartizat = fondBurseMeritRepartizat;
        }

        public int Compare(StudentRecord s1, StudentRecord s2)
        {
            // Primary sorting: descending by Media
            // Students with higher Media come first.
            int mediaComparison = s2.Media.CompareTo(s1.Media);
            if (mediaComparison != 0)
            {
                return mediaComparison;
            }

            // If Media is equal, apply tie-breaking criteria based on program and year
            if (_fondBurseMeritRepartizat.programStudiu.Equals("licenta", StringComparison.OrdinalIgnoreCase))
            {
                if (s1.An == 1) // a) Anul I de studiu, studii universitare de licență
                {
                    // 1. Media generală de la Examenul de Bacalaureat; (MediaBac)
                    // We use null-coalescing with -1m to ensure null values sort lower.
                    int bacMediaComparison = s2.MediaBac.CompareTo(s1.MediaBac);
                    if (bacMediaComparison != 0) return bacMediaComparison;

                    // 2. Media probei Matematică de la Examenul de Bacalaureat; (MediaBacMat)
                    int matematicaComparison = s2.MediaBacMat.CompareTo(s1.MediaBacMat);
                    if (matematicaComparison != 0) return matematicaComparison;

                    // 3. Rezultatul probei de evaluare a competențelor digitale, de la Examenul de Bacalaureat.
                    // This field is not in your StudentRecord. If you add it (e.g., string RezultatCompetenteDigitale),
                    // you'd need custom comparison logic (e.g., mapping A > B > C).
                    // Example placeholder:
                    // int digitalCompetenceComparison = string.Compare(s2.RezultatCompetenteDigitale, s1.RezultatCompetenteDigitale);
                    // if (digitalCompetenceComparison != 0) return digitalCompetenceComparison;
                }
                else // c) Anii II-IV, studii universitare de licență
                {
                    // 1. Media ponderată a semestrului II a anului universitar anterior;
                    // This field is not directly present in your StudentRecord.
                    // You might need to derive it or add a specific property.

                    // 2. Cea mai mare notă obținută la disciplinele cu cel mai mare număr de credite din anul universitar precedent;
                    // This requires a list of past disciplines with grades and credits, which is not in your StudentRecord.

                    // 3. A doua notă în ordinea descrescătoare a numărului de credite ale disciplinelor și a notelor obținute în anul universitar precedent;
                    // This also requires a list of past disciplines with grades and credits.

                    // 4. Media de admitere; (Could be MediaBac, or MEDG_ASL if applicable for this context)
                    // Assuming MediaBac can serve as "Media de admitere" for these years if no other specific field.
                    int admitereComparison = s2.MediaBac.CompareTo(s1.MediaBac);
                    if (admitereComparison != 0) return admitereComparison;
                }
            }
            else if (_fondBurseMeritRepartizat.programStudiu.Equals("master", StringComparison.OrdinalIgnoreCase))
            {
                if (s1.An == 1) // b) Anul I de studiu, studii universitare de masterat
                {
                    // 1. Nota obținută la interviul pentru testarea cunoștințelor și a capacităților cognitive din tematica specifică fiecărui masterat; (MediaInterviu)
                    int interviuComparison = s2.MediaInterviu.CompareTo(s1.MediaInterviu);
                    if (interviuComparison != 0) return interviuComparison;

                    // 2. Media generală ponderată a anilor de studiu, studii universitare de licență; (MediaDL)
                    int mediaLicentaComparison = s2.MediaDL.CompareTo(s1.MediaDL);
                    if (mediaLicentaComparison != 0) return mediaLicentaComparison;

                    // 3. Media generală de la Examenul de Bacalaureat. (MediaBac)
                    int bacMediaComparison = s2.MediaBac.CompareTo(s1.MediaBac);
                    if (bacMediaComparison != 0) return bacMediaComparison;
                }
                else // c) Anul II, studii universitare de masterat
                {
                    // 1. Media ponderată a semestrului II a anului universitar anterior;
                    // This field is not directly present in your StudentRecord.

                    // 2. Cea mai mare notă obținută la disciplinele cu cel mai mare număr de credite din anul universitar precedent;
                    // This requires a list of past disciplines with grades and credits.

                    // 3. A doua notă în ordinea descrescătoare a numărului de credite ale disciplinelor și a notelor obținute în anul universitar precedent;
                    // This also requires a list of past disciplines with grades and credits.

                    // 4. Media de admitere; (Could be MediaBac or MEDG_ASL)
                    // Assuming MediaBac can serve as "Media de admitere".
                    int admitereComparison = s2.MediaBac.CompareTo(s1.MediaBac);
                    if (admitereComparison != 0) return admitereComparison;

                    // 5. Pentru studii universitare de masterat - media generală ponderată a anilor de studiu, studii universitare de licență. (MediaDL)
                    int mediaLicentaComparison = s2.MediaDL.CompareTo(s1.MediaDL);
                    if (mediaLicentaComparison != 0) return mediaLicentaComparison;
                }
            }

            // If all available criteria are equal, maintain original order (stable sort).
            // If there's an inherent order you want to preserve if all criteria are the same,
            // you might use an Id or another stable identifier here, but 0 is standard for equal.
            return 0;
        }
    }

}
