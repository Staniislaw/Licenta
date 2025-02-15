using Burse.Models;

using Microsoft.EntityFrameworkCore;

namespace Burse.Data
{
    public class BurseDBContext : DbContext
    {
        public BurseDBContext(DbContextOptions options) : base(options)
        {

        }
        public DbSet<FondBurse> FondBurse { get;set;}
        public DbSet<FormatiiStudii> FormatiiStudii { get;set;} 
    }
}
