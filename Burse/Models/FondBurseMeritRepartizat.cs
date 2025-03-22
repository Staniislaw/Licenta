namespace Burse.Models
{
    public class FondBurseMeritRepartizat
    {
        public int ID { get; set; } 
        public string domeniu { get; set; }
        public decimal bursaAlocatata { get;set; }
        public string programStudiu { get; set; }
        public string Grupa { get; set; }
        public decimal SumaRamasa { get; set; }

        public List<StudentRecord> Studenti { get; set; } = new List<StudentRecord>();
    }
}
