namespace Burse.Models
{
    public class DictionaryPayload
    {
        public Dictionary<string, List<string>> GrupuriBurse { get; set; }
        public Dictionary<string, List<string>> Grupuri { get; set; }
        public Dictionary<string, List<string>> GrupProgramStudii { get; set; }
    }

}
