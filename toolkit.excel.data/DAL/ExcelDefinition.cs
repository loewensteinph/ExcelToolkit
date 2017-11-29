using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace toolkit.excel.data
{
    /// <summary>
    /// Represents an Excel Import as Database Table
    /// </summary>  
    [Table("ExcelDefinition", Schema = "ctrl")]
    public class ExcelDefinition
    {
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
        public ExcelDefinition()
        {
            ColumnMappings = new List<ColumnMapping>();
            Imports = new List<ExcelImport>();
        }
        [Key]
        public int DefinitionId { get; set; }
        public bool IsActive { get; set; }
        public string FileName { get; set; }
        public string SheetName { get; set; }
        public string Range { get; set; }
        public string Annotation { get; set; }
        public bool RangeWidthAuto { get; set; }
        public bool RangeHeightAuto { get; set; }
        public string TargetTable { get; set; }
        public string ConnectionString { get; set; }
        public bool HasHeaderRow { get; set; }
        public bool DeleteBeforeImport { get; set; }
        public bool BulkInsert { get; set; }
        public bool ValidateDataTypes { get; set; }
        public bool RollbackOnError { get; set; }
        public List<ColumnMapping> ColumnMappings { get; set; }
        public List<ExcelImport> Imports { get; set; }
    }
    /// <summary>
    /// Represents Columns of an Excel Import as Database Table
    /// </summary>  
    [Table("ColumnMapping", Schema = "ctrl")]
    public class ColumnMapping
    {
        [Key]
        public int ColumnMappingId { get; set; }
        public string SourceColumn { get; set; }
        public string TargetColumn { get; set; }
    }
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member
}