using System.Data.Entity;

namespace toolkit.excel.data
{
    public class ExcelDataContext : DbContext
    {
        public ExcelDataContext() : base("name=ExcelDataContextConnectionString")
        {
            Database.SetInitializer(new CreateDatabaseIfNotExists<ExcelDataContext>());
        }

        public DbSet<Log> Log { get; set; }
        public DbSet<ExcelDefinition> ExcelDefinition { get; set; }
        public DbSet<ColumnMapping> ColumnMapping { get; set; }
    }
}