using Burse.Models;
using Microsoft.AspNetCore.Mvc;
using System.Text.Json;
using System;
using Burse.Data;
using Microsoft.EntityFrameworkCore;

namespace Burse.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SettingsController : ControllerBase
    {
        private readonly IGrupuriService _grupuriService;

        public SettingsController(IGrupuriService grupuriService)
        {
            _grupuriService = grupuriService;
        }

        // Grupuri Bursa
        [HttpPost("grupuri-burse/add")]
        public async Task<IActionResult> AddDomeniuToGrupBursa([FromBody] GrupBursaEntry payload)
        {
            var added = await _grupuriService.AddDomeniuToGrupBursaAsync(payload);
            if (added)
                return Ok();
            return BadRequest("Entry already exists");
        }

        [HttpDelete("grupuri-burse/remove")]
        public async Task<IActionResult> RemoveDomeniuFromGrupBursa([FromQuery] string grup, [FromQuery] string domeniu)
        {
            await _grupuriService.RemoveDomeniuFromGrupBursaAsync(grup, domeniu);
            return Ok();
        }

        [HttpGet("grupuri-burse")]
        public async Task<IActionResult> GetGrupuriBurse()
        {
            var result = await _grupuriService.GetGrupuriBurseAsync();
            return Ok(result);
        }

        // Grupuri Domeniu
        [HttpGet("grupuri")]
        public async Task<IActionResult> GetGrupuri()
        {
            var result = await _grupuriService.GetGrupuriAsync();
            return Ok(result);
        }

        [HttpPost("grupuri/add")]
        public async Task<IActionResult> AddDomeniuToGrup([FromBody] GrupDomeniuEntry payload)
        {
            var added = await _grupuriService.AddDomeniuToGrupAsync(payload);
            if (added)
                return Ok();
            return BadRequest("Entry already exists");
        }

        [HttpDelete("grupuri/remove")]
        public async Task<IActionResult> RemoveDomeniuFromGrup([FromQuery] string grup, [FromQuery] string domeniu)
        {
            await _grupuriService.RemoveDomeniuFromGrupAsync(grup, domeniu);
            return Ok();
        }

        // Grupuri Program Studii
        [HttpPost("program-studii/add")]
        public async Task<IActionResult> AddDomeniuToGrupProgramStudii([FromBody] GrupProgramStudiiEntry payload)
        {
            var added = await _grupuriService.AddDomeniuToGrupProgramStudiiAsync(payload);
            if (added)
                return Ok();
            return BadRequest("Entry already exists");
        }

        [HttpDelete("program-studii/remove")]
        public async Task<IActionResult> RemoveDomeniuFromGrupProgramStudii([FromQuery] string grup, [FromQuery] string domeniu)
        {
            await _grupuriService.RemoveDomeniuFromGrupProgramStudiiAsync(grup, domeniu);
            return Ok();
        }

        [HttpGet("program-studii")]
        public async Task<IActionResult> GetGrupuriProgramStudii()
        {
            var result = await _grupuriService.GetGrupuriProgramStudiiAsync();
            return Ok(result);
        }

        // Grupuri PDF
        [HttpGet("grupuri-pdf")]
        public async Task<IActionResult> GetGrupuriPdf()
        {
            var result = await _grupuriService.GetGrupuriPdfAsync();
            return Ok(result);
        }

        [HttpPost("grupuri-pdf/add")]
        public async Task<IActionResult> AddValToPdfGroup([FromBody] GrupPdfEntry payload)
        {
            var added = await _grupuriService.AddValToPdfGroupAsync(payload);
            if (added)
                return Ok();
            return BadRequest("Entry already exists");
        }

        [HttpDelete("grupuri-pdf/remove")]
        public async Task<IActionResult> RemoveValFromPdfGroup([FromQuery] string grup, [FromQuery] string valoare)
        {
            await _grupuriService.RemoveValFromPdfGroupAsync(grup, valoare);
            return Ok();
        }
        [HttpGet("grupuri-acronime")]
        public async Task<IActionResult> GetGrupuriAcronime()
        {
            var result = await _grupuriService.GetGrupuriAcronimeAsync();
            return Ok(result);
        }

        [HttpPost("grupuri-acronime/add")]
        public async Task<IActionResult> AddValToAcronimGroup([FromBody] GrupAcronimEntry payload)
        {
            var added = await _grupuriService.AddValToAcronimGroupAsync(payload);
            if (added)
                return Ok();
            return BadRequest("Entry already exists");
        }

        [HttpDelete("grupuri-acronime/remove")]
        public async Task<IActionResult> RemoveValFromAcronimGroup([FromQuery] string grup, [FromQuery] string valoare)
        {
            await _grupuriService.RemoveValFromAcronimGroupAsync(grup, valoare);
            return Ok();
        }
    }
}
