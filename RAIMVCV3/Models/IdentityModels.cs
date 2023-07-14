using System.Data.Entity;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;

namespace RAIMVCV3.Models
{
    // You can add profile data for the user by adding more properties to your ApplicationUser class, please visit https://go.microsoft.com/fwlink/?LinkID=317594 to learn more.
    public class ApplicationUser : IdentityUser
    {
        public async Task<ClaimsIdentity> GenerateUserIdentityAsync(UserManager<ApplicationUser> manager)
        {
            // Note the authenticationType must match the one defined in CookieAuthenticationOptions.AuthenticationType
            var userIdentity = await manager.CreateIdentityAsync(this, DefaultAuthenticationTypes.ApplicationCookie);
            // Add custom user claims here
            return userIdentity;
        }
    }

    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public ApplicationDbContext()
            : base("DefaultConnection", throwIfV1Schema: false)
        {
           // Database.SetInitializer(new DatabaseInitializer());
        }

        public static ApplicationDbContext Create()
        {
            return new ApplicationDbContext();
        }
        public DbSet<Loan> Loans { get; set; }
        public DbSet<LoanStatus> LoanStatus { get; set; }
        public DbSet<Client> Client { get; set; }
        public DbSet<DwellingType> DwellingType { get; set; }
        public DbSet<Entity> Entity { get; set; }
        public DbSet<State> State { get; set; }
        public DbSet<Investor> Investor { get; set; }
        public DbSet<LoanType> LoanType { get; set; }
       // public DbSet<Role> Role { get; set; }
        //public DbSet<User> User { get; set; }
    }
    
}