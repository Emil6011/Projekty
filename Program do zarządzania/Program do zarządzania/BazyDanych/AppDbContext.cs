using Microsoft.EntityFrameworkCore;

using Zadanie_1_MVC_Dobre.Models;


namespace PostgreSQL.Data
{
    public class AppDbContext : DbContext
    {
		public AppDbContext(DbContextOptions options) : base(options)
		{
		}



		public DbSet<Klient> Klienci { get; set; }
	}
}