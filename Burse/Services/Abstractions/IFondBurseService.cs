using Burse.Models;

namespace Burse.Services.Abstractions
{
    public interface IFondBurseService
    {
        Task<List<FondBurse>> GetDateFromBursePerformanteAsync();
        Task<List<FormatiiStudii>> GetAllFromFormatiiStudiiAsync();
        Task<byte[]> GenerateCustomLayout2(string filePath, List<FondBurse> fonduri, List<FormatiiStudii> formatiiStudii, decimal disponibilBM);
        Task SaveNewStudentsAsync(List<StudentRecord> students);
        Task<List<StudentRecord>> GetStudentsWithBursaFromDatabaseAsync();
    }
}
