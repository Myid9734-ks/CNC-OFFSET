using System.Data.Entity;
using System.Data.SQLite.EF6;

namespace FanucFocasTutorial
{
    // EF6 SQLite Configuration
    public class SQLiteConfiguration : DbConfiguration
    {
        public SQLiteConfiguration()
        {
            SetProviderFactory("System.Data.SQLite", System.Data.SQLite.SQLiteFactory.Instance);
            SetProviderFactory("System.Data.SQLite.EF6", SQLiteProviderFactory.Instance);
        }
    }

    [DbConfigurationType(typeof(SQLiteConfiguration))]
    public class MacroDbContext : DbContext
    {
        public MacroDbContext() : base("name=MacroConnection")
        {
            // 데이터베이스가 없으면 생성
            Database.SetInitializer(new CreateDatabaseIfNotExists<MacroDbContext>());
        }

        public DbSet<MacroData> MacroData { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            // 복합 키 설정
            modelBuilder.Entity<MacroData>()
                .HasKey(m => new { m.IpAddress, m.MacroNumber });

            base.OnModelCreating(modelBuilder);
        }
    }
}
