using System.ComponentModel.DataAnnotations.Schema;

namespace Burse.Models
{
    public class BursaIstoric
    {
        public int Id { get; set; }

        [ForeignKey("StudentRecord")]
        public int StudentRecordId { get; set; }

        public string TipBursa { get; set; }  // Ex: BP1 ,BP2
        public string Motiv { get; set; }  // Ex: Media > 9.50
        public string Actiune { get; set; }  // Ex: Acordare, Retragere
        public string Etapa { get;set; } //Etapa 1,2,3;
        public decimal Suma { get; set; } //Suma acordata -> 
        public string Comentarii { get; set; }
        public DateTime DataModificare { get; set; }

        public StudentRecord StudentRecord { get; set; }
    }

}
