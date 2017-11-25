namespace toolkit.excel.data
{
    public class ExcelDefinition
    {
        public string SheetName { get; set; }
        public string Range { get; set; }
        public string FileName { get; set; }
        public bool HasHeaderRow { get; set; }
    }
}