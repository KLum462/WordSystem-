using Microsoft.EntityFrameworkCore;
using WordServer.Api.Data.Entities;

namespace WordServer.Api.Data
{
    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options)
        {
        }

        public DbSet<DocumentEntity> Documents => Set<DocumentEntity>();
        public DbSet<TemplateEntity> Templates => Set<TemplateEntity>();
        public DbSet<MetadataEntity> Metadata => Set<MetadataEntity>();
        public DbSet<SettingEntity> Settings => Set<SettingEntity>();
        public DbSet<LogEntryEntity> Logs => Set<LogEntryEntity>();

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<DocumentEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Title).IsRequired().HasMaxLength(255);
            });

            modelBuilder.Entity<TemplateEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Name).IsRequired().HasMaxLength(255);
            });

            modelBuilder.Entity<MetadataEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Author).HasMaxLength(255);
            });

            modelBuilder.Entity<SettingEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.HasIndex(e => e.Key).IsUnique();
            });

            modelBuilder.Entity<LogEntryEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
            });
        }
    }
}
