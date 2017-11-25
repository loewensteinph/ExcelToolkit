using System.Data.Entity.Migrations;

namespace toolkit.excel.data.Migrations
{
    public partial class initial2 : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                    "dbo.ColumnMappings",
                    c => new
                    {
                        ColumnMappingId = c.Int(false, true),
                        SourceColumn = c.String(),
                        TargetColumn = c.String(),
                        ExcelDefinition_DefinitionId = c.Int()
                    })
                .PrimaryKey(t => t.ColumnMappingId)
                .ForeignKey("dbo.ExcelDefinitions", t => t.ExcelDefinition_DefinitionId)
                .Index(t => t.ExcelDefinition_DefinitionId);

            CreateTable(
                    "dbo.ExcelDefinitions",
                    c => new
                    {
                        DefinitionId = c.Int(false, true),
                        FileName = c.String(),
                        SheetName = c.String(),
                        Range = c.String(),
                        TargetTable = c.String(),
                        ConnectionString = c.String(),
                        HasHeaderRow = c.Boolean(false)
                    })
                .PrimaryKey(t => t.DefinitionId);
        }

        public override void Down()
        {
            DropForeignKey("dbo.ColumnMappings", "ExcelDefinition_DefinitionId", "dbo.ExcelDefinitions");
            DropIndex("dbo.ColumnMappings", new[] {"ExcelDefinition_DefinitionId"});
            DropTable("dbo.ExcelDefinitions");
            DropTable("dbo.ColumnMappings");
        }
    }
}