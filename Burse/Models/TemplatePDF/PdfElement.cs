namespace Burse.Models.TemplatePDF
{
    public class PdfElement
    {
        public string Type { get; set; }
        public string Content { get; set; }
        public PdfStyle Style { get; set; }
        public List<string> Domenii { get; set; } // nou!

    }

    public class PdfStyle
    {
        public int FontSize { get; set; } = 14;
        public string TextAlign { get; set; } = "left";
        public string Color { get; set; } = "#000000";
    }

    public class PdfRequest
    {
        public List<PdfElement> Elements { get; set; }
    }

}
