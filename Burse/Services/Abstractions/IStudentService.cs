using Burse.Models;

namespace Burse.Services.Abstractions
{
    public interface IStudentService
    {
        Task<List<StudentRecord>> GetAllAsync();
        Task<StudentRecord> GetByIdAsync(int id);
        Task<StudentRecord> UpdateBursaAsync(int id, string bursa);
        Task<byte[]> ExportStudentiExcelAsync();

    }
}
