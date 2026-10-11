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

        // Additional enterprise tables
        public DbSet<UserEntity> Users => Set<UserEntity>();
        public DbSet<DocumentVersionEntity> DocumentVersions => Set<DocumentVersionEntity>();
        public DbSet<TemplateCategoryEntity> TemplateCategories => Set<TemplateCategoryEntity>();
        public DbSet<ProcessingHistoryEntity> ProcessingHistory => Set<ProcessingHistoryEntity>();
        public DbSet<DocumentCommentEntity> DocumentComments => Set<DocumentCommentEntity>();

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

            modelBuilder.Entity<UserEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.HasIndex(e => e.Username).IsUnique();
                entity.HasIndex(e => e.ApiKey).IsUnique();
            });

            modelBuilder.Entity<DocumentVersionEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.HasIndex(e => new { e.DocumentId, e.VersionNumber });
            });

            modelBuilder.Entity<TemplateCategoryEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.HasIndex(e => e.Name).IsUnique();
            });

            modelBuilder.Entity<ProcessingHistoryEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.HasIndex(e => e.Timestamp);
            });

            modelBuilder.Entity<DocumentCommentEntity>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.HasIndex(e => e.DocumentId);
            });
        }
    }
}
