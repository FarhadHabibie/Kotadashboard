using KOTAdashboard.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace KOTAdashboard.Data

{
    public class KOTAdashboardContext : DbContext
    {
        public KOTAdashboardContext()
        {
        }

        public KOTAdashboardContext(DbContextOptions<KOTAdashboardContext> options) : base(options)
        {
            
        }    
        public DbSet<Corona> Coronas { get; set; }
    }
}
