namespace Burse.Models.DTO
{
    public class StudentDto
    {
        // câmpuri din StudentRecord
        public int Id { get; set; }
        public string Emplid { get; set; }
        public string CNP { get; set; }
        public string NumeStudent { get; set; }
        public int An { get; set; }
        public decimal Media { get; set; }
        public int PunctajAn { get; set; }
        public int CO { get; set; }
        public int RO { get; set; }
        public int TC { get; set; }
        public int TR { get; set; }
        public string Bursa { get; set; }

        // câmpuri preluate din FondBurseMeritRepartizat
        public string Domeniu { get; set; }
        public string ProgramStudiu { get; set; }
        public string Grupa { get; set; }
        public decimal? SumaRamasa { get; set; }
        public List<BursaIstoricDto> IstoricBursa { get; set; } = new();
    }
    public class BursaIstoricDto
    {
        public string TipBursa { get; set; }
        public string Motiv { get; set; }
        public string Actiune { get; set; }
        public decimal Suma { get; set; }
        public string SursaFinantare { get; set; }
        public string Comentarii { get; set; }
        public string UserModificare { get; set; }
        public DateTime DataModificare { get; set; }
    }
}
