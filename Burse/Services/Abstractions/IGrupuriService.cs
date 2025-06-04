using Burse.Models;

using System.Collections.Generic;
using System.Threading.Tasks;

public interface IGrupuriService
{
    // Grupuri Bursa
    Task<bool> AddDomeniuToGrupBursaAsync(GrupBursaEntry payload);
    Task RemoveDomeniuFromGrupBursaAsync(string grup, string domeniu);
    Task<Dictionary<string, List<string>>> GetGrupuriBurseAsync();

    // Grupuri Domeniu
    Task<bool> AddDomeniuToGrupAsync(GrupDomeniuEntry payload);
    Task RemoveDomeniuFromGrupAsync(string grup, string domeniu);
    Task<Dictionary<string, List<string>>> GetGrupuriAsync();
    Task<List<string>> GetGrupuriAsyncByGrup(string grup);


    // Grupuri Program Studii
    Task<bool> AddDomeniuToGrupProgramStudiiAsync(GrupProgramStudiiEntry payload);
    Task RemoveDomeniuFromGrupProgramStudiiAsync(string grup, string domeniu);
    Task<Dictionary<string, List<string>>> GetGrupuriProgramStudiiAsync();

    // Grupuri PDF
    Task<bool> AddValToPdfGroupAsync(GrupPdfEntry payload);
    Task RemoveValFromPdfGroupAsync(string grup, string valoare);
    Task<Dictionary<string, List<string>>> GetGrupuriPdfAsync();

    // Grupuri Acronime
    Task<bool> AddValToAcronimGroupAsync(GrupAcronimEntry payload);
    Task RemoveValFromAcronimGroupAsync(string grup, string valoare);
    Task<Dictionary<string, List<string>>> GetGrupuriAcronimeAsync();
}
