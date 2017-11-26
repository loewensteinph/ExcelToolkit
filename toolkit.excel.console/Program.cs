using toolkit.excel.data;

namespace toolkit.excel.console
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var reader = new ExcelReader("TestWB.xlsx", "Sheet1", "A1:A4", true);
            var result = reader.Read();

            /*
            ExcelDefinition def = new ExcelDefinition();
            def.FileName = "TestWB.xlsx";
            def.SheetName = "Sheet1";
            def.HasHeaderRow = true;
            def.Range = "A1:L108";
            def.TargetTable = "TestTable";
            def.ConnectionString = "Data Source=.;Initial Catalog=ExcelDB;Integrated Security=true";

            List<ColumnMapping> map = new List<ColumnMapping>();

            map.Add(new ColumnMapping() {SourceColumn = "name" ,TargetColumn = "name"});
            map.Add(new ColumnMapping() { SourceColumn = "object_id", TargetColumn = "object_id" });
            map.Add(new ColumnMapping() { SourceColumn = "principal_id", TargetColumn = "principal_id" });
            map.Add(new ColumnMapping() { SourceColumn = "create_date", TargetColumn = "create_date" });

            def.ColumnMappings.AddRange(map);

            dataContext.Entry(def).State = EntityState.Added;
            dataContext.SaveChanges();
            */
        }
    }
}