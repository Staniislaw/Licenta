using System.ComponentModel.DataAnnotations.Schema;

namespace Burse.Models
{
    public class StudentRecord
    {
        public int Id { get; set; }
        public string Emplid { get; set; }  // ID angajat
        public string CNP { get; set; }  // Cod Numeric Personal
        public string NumeStudent { get; set; }  // Nume student
        public string TaraCetatenie { get; set; }  // Țară Cetățenie
        public int An { get; set; }  // Anul de studiu
        public decimal Media { get; set; }  // Media generală
        public int PunctajAn { get; set; }  // Punctajul anual
        public int CO { get; set; }  // Coloana CO CO – credite obţinute in anul curent
        public int RO { get; set; }  // RO – restanţe anul curent
        public int TC { get; set; }  // TC – creditele obţinute pe anii anteriori+ credite anul curent
        public int TR { get; set; }  // TR – restanţele anii precedenti + restante anul curent
        public string SursaFinantare { get; set; }  // Sursa de finanțare
        public string Bursa { get; set; }
        public decimal SumaBursa { get; set; }

        [ForeignKey("FondBurseMeritRepartizat")]
        public int FondBurseMeritRepartizatId { get; set; }

        public FondBurseMeritRepartizat FondBurseMeritRepartizat { get; set; }
        public virtual ICollection<BursaIstoric> IstoricBursa { get; set; }


    }
}
