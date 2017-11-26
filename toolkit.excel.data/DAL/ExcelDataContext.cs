using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Migrations;

namespace toolkit.excel.data
{
    internal sealed class Configuration : DbMigrationsConfiguration<ExcelDataContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = false;
        }
    }
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

    public class ExcelUnitTestDataContext : ExcelDataContext
    {
        public ExcelUnitTestDataContext()
        {
            Database.SetInitializer(new ExcelDataContextSeedInitializer());
        }
    }
    public class ExcelDataContextSeedInitializer : DropCreateDatabaseAlways<ExcelDataContext>
    {
        public override void InitializeDatabase(ExcelDataContext context)
        {
            if (context.Database.Exists())
            {
                context.Database.ExecuteSqlCommand(TransactionalBehavior.DoNotEnsureTransaction
                    , String.Format("ALTER DATABASE [{0}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE", context.Database.Connection.Database));
                context.Database.Delete();
            }
            context.Database.Create();
            Seed(context);        
        }
        protected override void Seed(ExcelDataContext context)
        {
            ExcelDefinition def = new ExcelDefinition
            {
                FileName = @"TestWorkbooks\UT1.xlsx",
                SheetName = "Sheet1",
                HasHeaderRow = true,
                Range = "A1:D5",
                TargetTable = "Test.UT1",
                ConnectionString = "Data Source=.;Initial Catalog=ExcelDBUNitTest;Integrated Security=true"
            };

            List<ColumnMapping> map = new List<ColumnMapping>
            {
                new ColumnMapping() {SourceColumn = "StringTest", TargetColumn = "StringTest"},
                new ColumnMapping() {SourceColumn = "DecimalTest", TargetColumn = "DecimalTest"},
                new ColumnMapping() {SourceColumn = "IntTest", TargetColumn = "IntTest"},
                new ColumnMapping() {SourceColumn = "GuidTest", TargetColumn = "GuidTest"}
            };

            def.ColumnMappings.AddRange(map);
            context.Entry(def).State = EntityState.Added;

            def = new ExcelDefinition
            {
                FileName = @"TestWorkbooks\UT2.xlsx",
                SheetName = "Sheet1",
                HasHeaderRow = true,
                Range = "A1:E5",
                TargetTable = "Test.UT2",
                ConnectionString = "Data Source=.;Initial Catalog=ExcelDBUNitTest;Integrated Security=true"
            };

            map = new List<ColumnMapping>
            {
                new ColumnMapping() {SourceColumn = "StringTest", TargetColumn = "StringTest1"},
                new ColumnMapping() {SourceColumn = "DecimalTest", TargetColumn = "DecimalTest1"},
                new ColumnMapping() {SourceColumn = "IntTest", TargetColumn = "IntTest1"},
                new ColumnMapping() {SourceColumn = "GuidTest", TargetColumn = "GuidTest1"},
                new ColumnMapping() {SourceColumn = "DateTest", TargetColumn = "DateTest"}    
            };
            def.ColumnMappings.AddRange(map);
            context.Entry(def).State = EntityState.Added;

            context.SaveChanges();
        }
    }
}