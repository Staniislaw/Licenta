using System.ComponentModel.DataAnnotations.Schema;

namespace Burse.Models
{
    public class StudentRecord
    {
        public int Id { get; set; }
        public string Emplid { get; set; } = string.Empty;// ID angajat
        public string CNP { get; set; } = string.Empty;  // Cod Numeric Personal
        public string NumeStudent { get; set; } = string.Empty;  // Nume student
        public string TaraCetatenie { get; set; } = string.Empty;  // Țară Cetățenie
        public int An { get; set; }  // Anul de studiu
        public decimal Media { get; set; }  // Media generală
        public decimal MediaBac { get; set; } //Media BAC pentru anii 1
        public decimal MediaBacMat { get; set; }//media BAC matematica
        public decimal MediaInterviu { get; set; }
        public decimal MediaDL { get; set; }
        public decimal MEDG_ASL { get; set; }
        public int PunctajAn { get; set; }  // Punctajul anual
        public int CO { get; set; }  // Coloana CO CO – credite obţinute in anul curent
        public int RO { get; set; }  // RO – restanţe anul curent
        public int TC { get; set; }  // TC – creditele obţinute pe anii anteriori+ credite anul curent
        public int TR { get; set; }  // TR – restanţele anii precedenti + restante anul curent
        public string SursaFinantare { get; set; } = string.Empty;  // Sursa de finanțare
        public string Bursa { get; set; } = string.Empty;
        public decimal SumaBursa { get; set; }

        [ForeignKey("FondBurseMeritRepartizat")]
        public int FondBurseMeritRepartizatId { get; set; }
        public FondBurseMeritRepartizat FondBurseMeritRepartizat { get; set; }
        public virtual ICollection<BursaIstoric> IstoricBursa { get; set; }
        public string TipInconsistenta { get; set; } = string.Empty;
        [Column(TypeName = "longtext")]
        public string Observatii { get; set; } = string.Empty;
    }
}
