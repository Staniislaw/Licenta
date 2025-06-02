using Burse.Models;

namespace Burse.Services.Abstractions
{
    public interface IFondBurseMeritRepartizatService
    {
        Task<List<FondBurseMeritRepartizat>> GetAllAsync();
        Task<FondBurseMeritRepartizat> GetByDomeniuAsync(string domeniu);
        Task<bool> AddAsync(FondBurseMeritRepartizat newFond);
        Task<bool> UpdateAsync(string domeniu, decimal newAmount);
        Task<bool> DeleteAsync(string domeniu);
        Task UpdateAsync(FondBurseMeritRepartizat fond);
        Task<bool> UpdateFondAsync(int id, decimal suma, bool scade);

    }
}
