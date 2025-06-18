using System.ComponentModel.DataAnnotations.Schema;

namespace Burse.Models
{
    public class FondBurse
    {
        public int Id { get; set; }
        public string CategorieBurse { get; set; }
        [Column(TypeName = "decimal(18, 2)")]
        public Decimal ValoreaLunara { get; set; }
    }
}
