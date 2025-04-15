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
        private readonly BurseDBContext _context;

        public SettingsController(BurseDBContext context)
        {
            _context = context;
        }

        [HttpPost("grupuri-burse/add")]
        public async Task<IActionResult> AddDomeniuToGrup([FromBody] GrupBursaEntry payload)
        {
            // Verifici dacă există deja
            var exists = await _context.GrupBursa
                .AnyAsync(e => e.GrupBursa == payload.GrupBursa && e.Domeniu == payload.Domeniu);

            if (!exists)
            {
                _context.GrupBursa.Add(payload);
                await _context.SaveChangesAsync();
            }

            return Ok();
        }
        [HttpDelete("grupuri-burse/remove")]
        public async Task<IActionResult> RemoveDomeniuFromGrupBursa([FromQuery] string grup, [FromQuery] string domeniu)
        {
            var entry = await _context.GrupBursa
                .FirstOrDefaultAsync(e => e.GrupBursa == grup && e.Domeniu == domeniu);

            if (entry != null)
            {
                _context.GrupBursa.Remove(entry);
                await _context.SaveChangesAsync();
            }

            return Ok();
        }

        [HttpGet("grupuri-burse")]
        public async Task<IActionResult> GetGrupuriBurse()
        {
            var entries = await _context.GrupBursa.ToListAsync();

            var grouped = entries
                .GroupBy(e => e.GrupBursa)
                .ToDictionary(g => g.Key, g => g.Select(x => x.Domeniu).ToList());

            return Ok(grouped);
        }



        [HttpGet("grupuri")]
        public async Task<IActionResult> GetGrupuri()
        {
            var entries = await _context.GrupDomeniu.ToListAsync();

            var grouped = entries
                .GroupBy(e => e.Grup)
                .ToDictionary(g => g.Key, g => g.Select(x => x.Domeniu).ToList());

            return Ok(grouped);
        }

        [HttpPost("grupuri/add")]
        public async Task<IActionResult> AddDomeniuToGrup([FromBody] GrupDomeniuEntry payload)
        {
            var exists = await _context.GrupDomeniu
                .AnyAsync(e => e.Grup == payload.Grup && e.Domeniu == payload.Domeniu);

            if (!exists)
            {
                _context.GrupDomeniu.Add(payload);
                await _context.SaveChangesAsync();
            }

            return Ok();
        }

        [HttpDelete("grupuri/remove")]
        public async Task<IActionResult> RemoveDomeniuFromGrup([FromQuery] string grup, [FromQuery] string domeniu)
        {
            var entry = await _context.GrupDomeniu
                .FirstOrDefaultAsync(e => e.Grup == grup && e.Domeniu == domeniu);

            if (entry != null)
            {
                _context.GrupDomeniu.Remove(entry);
                await _context.SaveChangesAsync();
            }

            return Ok();
        }


        [HttpPost("program-studii/add")]
        public async Task<IActionResult> AddDomeniuToGrupProgramStudii([FromBody] GrupProgramStudiiEntry payload)
        {
            // Verifici dacă există deja
            var exists = await _context.GrupBursa
                .AnyAsync(e => e.GrupBursa == payload.Grup && e.Domeniu == payload.Domeniu);

            if (!exists)
            {
                _context.GrupProgramStudii.Add(payload);
                await _context.SaveChangesAsync();
            }

            return Ok();
        }
        [HttpDelete("program-studii/remove")]
        public async Task<IActionResult> RemoveDomeniuFromGrupProgramStudii([FromQuery] string grup, [FromQuery] string domeniu)
        {
            var entry = await _context.GrupProgramStudii
                .FirstOrDefaultAsync(e => e.Grup == grup && e.Domeniu == domeniu);

            if (entry != null)
            {
                _context.GrupProgramStudii.Remove(entry);
                await _context.SaveChangesAsync();
            }

            return Ok();
        }

        [HttpGet("program-studii")]
        public async Task<IActionResult> GetGrupuriProgramStudii()
        {
            var entries = await _context.GrupProgramStudii.ToListAsync();

            var grouped = entries
                .GroupBy(e => e.Grup)
                .ToDictionary(g => g.Key, g => g.Select(x => x.Domeniu).ToList());

            return Ok(grouped);
        }

    }
}
