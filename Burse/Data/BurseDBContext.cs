﻿using Burse.Models;
using Burse.Models.TemplatePDF;

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
        public DbSet<FondBurseMeritRepartizat> FondBurseMeritRepartizat { get;set;} 
        public DbSet<StudentRecord> StudentRecord { get;set;}
        public DbSet<GrupBursaEntry> GrupBursa { get; set; }
        public DbSet<GrupDomeniuEntry> GrupDomeniu { get; set; }
        public DbSet<GrupProgramStudiiEntry> GrupProgramStudii { get; set; }
        public DbSet<GrupPdfEntry> GrupPDF { get; set; }
        public DbSet<BursaIstoric> BursaIstoric { get; set; }
        public DbSet<TemplateEntity> TemplateEntity { get; set; }
        public DbSet<GrupAcronimEntry> GrupAcronim { get; set; }
    }
}
