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
}
