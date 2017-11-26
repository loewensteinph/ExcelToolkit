using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace toolkit.excel.data
{
    public class ExcelDefinition
    {
        public ExcelDefinition()
        {
            ColumnMappings = new List<ColumnMapping>();
        }
        [Key]
        public int DefinitionId { get; set; }
        public string FileName { get; set; }
        public string SheetName { get; set; }
        public string Range { get; set; }
        public string TargetTable { get; set; }
        public string ConnectionString { get; set; }
        public bool HasHeaderRow { get; set; }
        public List<ColumnMapping> ColumnMappings { get; set; }
    }

    public class ColumnMapping
    {
        [Key]
        public int ColumnMappingId { get; set; }
        public string SourceColumn { get; set; }
        public string TargetColumn { get; set; }
    }
}