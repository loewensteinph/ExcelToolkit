using System.Data.Entity.Migrations;

namespace toolkit.excel.data.Migrations
{
    public partial class initial : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                    "dbo.Logs",
                    c => new
                    {
                        LogId = c.Long(false, true),
                        Date = c.DateTime(false),
                        Thread = c.String(),
                        Level = c.String(),
                        Logger = c.String(),
                        Message = c.String(),
                        Exception = c.String()
                    })
                .PrimaryKey(t => t.LogId);
        }

        public override void Down()
        {
            DropTable("dbo.Logs");
        }
    }
}