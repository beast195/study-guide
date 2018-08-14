using Microsoft.AspNet.Identity.EntityFramework;
using StudyGuide.Framework.Core.Models;
using System;
using System.Data.Entity;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace StudyGuide.Framework
{
    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public DbSet<Book> Books { get; set; }
        public DbSet<Chapter> Chapters { get; set; }
        public DbSet<SentenceLibrary> SentenceLibraries { get; set; }

        //public DbSet<ApplicationUserGroup> ApplicationUserGroups { get; set; }

        public ApplicationDbContext()
            : base("DefaultConnection", throwIfV1Schema: false)
        {
        }

        public static ApplicationDbContext Create()
        {
            return new ApplicationDbContext();
        }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<ApplicationUser>();
        }

        public override Task<int> SaveChangesAsync(CancellationToken cancellationToken)
        {
            AddTimestamps();
            return base.SaveChangesAsync(cancellationToken);
        }

        private void AddTimestamps()
        {
            var changes = ChangeTracker
                .Entries()
                .Where(e => e.State == EntityState.Added || e.State == EntityState.Modified);
            foreach (var entry in changes)
            {
                var type = entry.Entity.GetType();

                bool isTimestampable = type.GetInterfaces().Any(x => x == typeof(StudyGuide.Framework.Core.Interface.ITimestamp));

                if (isTimestampable)
                {
                    PropertyInfo property;
                    if (entry.State == EntityState.Added)
                    {
                        property = type.GetProperty("CreatedAt");
                        property?.SetValue(entry.Entity, DateTimeOffset.UtcNow, null);
                    }

                    property = type.GetProperty("UpdatedAt");
                    property?.SetValue(entry.Entity, DateTimeOffset.UtcNow, null);
                }
            }
        }
    }
}