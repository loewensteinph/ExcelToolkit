using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace toolkit.excel.data
{
    /// <summary>
    /// Represents an Excel Import as Database Table
    /// </summary>  
    [Table("ExcelImport", Schema = "log")]
    public class ExcelImport
    {
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ImportId { get; set; }
        ExcelDefinition Definition { get; set; }
        public DateTime ImportTimestamp { get; set; }
        public int RowsImported { get; set; }
        public int RowsWithErrors { get; set; }
        public string ResultStatus { get; set; }
    }
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member
}