using Burse.Models;

namespace Burse.Services.Abstractions
{
    public interface IBurseIstoricService
    {
        Task AddToIstoricAsync(StudentRecord student, string actiune, decimal suma, string motiv);
    }
}
