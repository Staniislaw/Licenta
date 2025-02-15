using Burse.Models;

namespace Burse.Services.Abstractions
{
    public interface IFondBurseService
    {
        Task<List<FondBurse>> GetDateFromBursePerformanteAsync();
        Task<List<FormatiiStudii>> GetAllFromFormatiiStudiiAsync();
    }
}
