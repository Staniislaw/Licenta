namespace Burse.Models
{
    public class GrupDomeniuEntry
    {
        public int Id { get; set; }
        public string Grup { get; set; }       // ex: "IEN/ME/ETI"
        public string Domeniu { get; set; }    // ex: "C"
    }

    public class GrupBursaEntry
    {
        public int Id { get; set; }
        public string GrupBursa { get; set; }  // ex: "G1"
        public string Domeniu { get; set; }    // ex: "C"
    }

    public class GrupProgramStudiiEntry
    {
        public int Id { get; set; }
        public string Grup { get; set; } = string.Empty;
        public string Domeniu { get; set; } = string.Empty;
    }
    public class GrupPdfEntry
    {
        public int Id { get; set; }
        public string Grup { get; set; } = null!;
        public string Valoare { get; set; } = null!;
    }

    public class GrupAcronimEntry
    {
        public int Id { get; set; }
        public string Grup { get; set; } = string.Empty;
        public string Valoare { get; set; } = string.Empty;

    }


}
