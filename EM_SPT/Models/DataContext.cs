using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EM_SPT.Models
{
    public class DataContext : DbContext
    {
        public DataContext()
        {

        }
        public DataContext(DbContextOptionsBuilder optionsBuilder)
        {
            OnConfiguring(optionsBuilder);
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json");
            var configuration = builder.Build();

            optionsBuilder.UseMySql(configuration["ConnectionStrings:DefaultConnection"]);

        }

        public DbSet<user> User { get; set; }
        public DbSet<mo> Mo { get; set; }
        public DbSet<oo> Oo { get; set; }
        public DbSet<answer> Answer { get; set; }
        public DbSet<klass> Klass { get; set; }
        public DbSet<param> Param { get; set; }
        public CompositeModel Composite { get; set; }

    }
}
