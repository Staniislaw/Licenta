using System.ComponentModel.DataAnnotations.Schema;

namespace Burse.Models.TemplatePDF
{
    public class TemplateEntity
    {
        public int Id { get; set; } 
        public string Name { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [Column(TypeName = "longtext")] 
        public string ElementsJson { get; set; } = string.Empty;
    }

}
