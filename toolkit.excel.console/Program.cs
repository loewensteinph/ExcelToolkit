using toolkit.excel.data;

namespace toolkit.excel.console
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var reader = new ExcelReader("TestWB.xlsx", "Sheet1", "A1:A4", true);
            var result = reader.Read();
            reader.Dispose();
        }
    }
}