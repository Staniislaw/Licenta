using Burse.Data;
using Burse.Models.TemplatePDF;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace Burse.Controllers
{
    [ApiController]
    [Route("api/template")]
    public class TemplateController : ControllerBase
    {
        private readonly BurseDBContext _context;

        public TemplateController(BurseDBContext context)
        {
            _context = context;
        }

        [HttpPost("SaveTemplate")]
        public async Task<IActionResult> SaveTemplate([FromBody] TemplateEntity template)
        {
            if (string.IsNullOrWhiteSpace(template.Name) || string.IsNullOrWhiteSpace(template.ElementsJson))
                return BadRequest("Template Name și ElementsJson sunt obligatorii.");

            template.CreatedAt = DateTime.UtcNow;
            _context.TemplateEntity.Add(template);
            await _context.SaveChangesAsync();

            return Ok(template);
        }

        [HttpGet("GetTemplates")]
        public async Task<IActionResult> GetTemplates()
        {
            var templates = await _context.TemplateEntity
                .OrderByDescending(t => t.CreatedAt)
                .ToListAsync();

            return Ok(templates);
        }

        [HttpGet("GetTemplate")]
        public async Task<IActionResult> GetTemplate([FromQuery] int id)
        {
            var template = await _context.TemplateEntity.FindAsync(id);

            if (template == null)
                return NotFound();

            return Ok(template);
        }
        [HttpDelete("DeleteTemplate{id}")]
        public async Task<IActionResult> DeleteTemplate(int id)
        {
            var template = await _context.TemplateEntity.FindAsync(id);
            if (template == null)
                return NotFound();

            _context.TemplateEntity.Remove(template);
            await _context.SaveChangesAsync();

            return Ok(new { message = "Template șters cu succes!" });
        }
        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateTemplate(int id, [FromBody] TemplateEntity updatedTemplate)
        {
            var existingTemplate = await _context.TemplateEntity.FindAsync(id);
            if (existingTemplate == null)
                return NotFound();

            existingTemplate.Name = updatedTemplate.Name;
            existingTemplate.ElementsJson = updatedTemplate.ElementsJson;
            existingTemplate.CreatedAt = DateTime.UtcNow; // sau păstrezi data veche dacă vrei

            await _context.SaveChangesAsync();

            return Ok(existingTemplate);
        }
    }
}
