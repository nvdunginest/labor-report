using LaborReport.Datas.Models;
using Microsoft.EntityFrameworkCore;
using System;

namespace LaborReport.Datas
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
        }

        public DbSet<InOut> InOuts { get; set; }
    }
}
