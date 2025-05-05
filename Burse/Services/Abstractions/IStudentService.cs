using Burse.Models;

namespace Burse.Services.Abstractions
{
    public interface IStudentService
    {
        Task<List<StudentRecord>> GetAllAsync();
        Task<StudentRecord> UpdateBursaAsync(int id, string bursa);

    }
}
