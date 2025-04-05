using Burse.Data;

using Microsoft.EntityFrameworkCore;

namespace Burse.Helpers
{
        /*  private static readonly Dictionary<string, List<string>> Grupuri = new()
          {
              { "IEN/ME/ETI", new() { "IEN", "ME", "ETI" } },
              { "SE", new() { "SE" } },
              { "SE-DUAL", new() { "SE-DUAL" } },
              { "AIA", new() { "AIA" } },
              { "AIA-DUAL", new() { "AIA-DUAL" } },
              { "C", new() { "C" } },
              { "C-DUAL", new() { "C-DUAL" } },
              { "IETTI/RST", new() { "IETTI", "RST" } },
              { "ESM", new() { "ESM" } },
              { "ESCCPA", new() { "ESCCPA" } },
              { "SIC", new() { "SIC" } },
              { "SMPCPE", new() { "SMPCPE" } },
              { "TAIMAE", new() { "TAIMAE" } },
              { "RCC", new() { "RCC" } },
              { "SC", new() { "SC" } },
          };

          public static string GetGrupa(string domeniu)
          {
              string domeniuSimplificat = domeniu.Split(' ')[0];

              if (domeniu.Contains("-DUAL"))
                  domeniuSimplificat += "-DUAL";

              foreach (var grup in Grupuri)
              {
                  if (grup.Value.Contains(domeniuSimplificat))
                      return grup.Key;
              }

              return "Necunoscut";
          }

          public static readonly Dictionary<string, List<string>> GrupuriBurse = new()
          {
              { "G1", new List<string> { "C" } }, 

              { "G2", new List<string> { "AIA", "IETTI", "RST" } },

              { "G3", new List<string> { "ESSCA", "SE", "ESM", "IEN", "ETI", "ME" } }, 
          };*/
        public class GrupuriDomeniiHelper
        {
            private readonly BurseDBContext _context;

            public GrupuriDomeniiHelper(BurseDBContext context)
            {
                _context = context;
            }
            public async Task<string> GetGrupaAsync(string domeniu)
            {
                var domeniuSimplificat = domeniu.Split(' ')[0];

                if (domeniu.Contains("-DUAL"))
                    domeniuSimplificat += "-DUAL";

                var entries = await _context.GrupDomeniu.ToListAsync();

                var grupuri = entries
                    .GroupBy(g => g.Grup)
                    .ToDictionary(g => g.Key, g => g.Select(x => x.Domeniu).ToList());


            foreach (var grup in grupuri)
                {
                    if (grup.Value.Contains(domeniuSimplificat))
                        return grup.Key;
                }

                return "Necunoscut";
            }
            public async Task<Dictionary<string, List<string>>> GetGrupuriAsync()
            {
                return await _context.GrupDomeniu
                    .GroupBy(g => g.Grup)
                    .ToDictionaryAsync(g => g.Key, g => g.Select(x => x.Domeniu).ToList());
            }
            public async Task<Dictionary<string, List<string>>> GetGrupuriBurseAsync()
            {
                var entries = await _context.GrupBursa.ToListAsync();

                return entries
                    .GroupBy(g => g.GrupBursa)
                    .ToDictionary(g => g.Key, g => g.Select(x => x.Domeniu).ToList());
            }
    }
}
