using System.ComponentModel.DataAnnotations.Schema;

namespace Burse.Models
{
    public class BursaIstoric
    {
        public int Id { get; set; }

        [ForeignKey("StudentRecord")]
        public int StudentRecordId { get; set; }

        public string TipBursa { get; set; }  // Ex: BP1 ,BP2
        [Column(TypeName = "longtext")]
        public string Motiv { get; set; } = string.Empty;  // Ex: Media > 9.50
        public string Actiune { get; set; } = string.Empty;  // Ex: Acordare, Retragere
        public string Etapa { get; set; } = string.Empty;//Etapa 1,2,3;
        [Column(TypeName = "decimal(18, 2)")]
        public decimal Suma { get; set; }  //Suma acordata -> 
        [Column(TypeName = "longtext")]
        public string Comentarii { get; set; } = string.Empty;
        public DateTime DataModificare { get; set; }
        [Column(TypeName = "longtext")]
        public string ComentariiAI { get; set; } = string.Empty;
        public StudentRecord StudentRecord { get; set; }
    }

}
