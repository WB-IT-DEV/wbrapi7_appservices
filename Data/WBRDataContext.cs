using Microsoft.EntityFrameworkCore;
using wbrapi7_appservices.Entities;
namespace wbrapi7_appservices.Data
{
    public class WBRDataContext : DbContext
    {
        public WBRDataContext(DbContextOptions<WBRDataContext> options) : base(options) { }

        public DbSet<vciSafIncHeadStatus> vciSafIncHeadStatus { get; set; }
        public DbSet<vapJIBSharepoint> vapJIBSharepoint { get; set; }
        public DbSet<vapJIBHeader> vapJIBHeader { get; set; }


        

        public DbSet<tapEIBSubmitSupplierInv> tapEIBSubmitSupplierInv { get; set; }


    }
}
