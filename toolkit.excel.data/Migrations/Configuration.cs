using System.Data.Entity.Migrations;

namespace toolkit.excel.data.Migrations
{
    internal sealed class Configuration : DbMigrationsConfiguration<ExcelDataContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = false;
        }

        protected override void Seed(ExcelDataContext context)
        {
            //  This method will be called after migrating to the latest version.

            //  You can use the DbSet<T>.AddOrUpdate() helper extension method 
            //  to avoid creating duplicate seed data.
        }
    }
}