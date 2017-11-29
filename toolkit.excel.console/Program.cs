using toolkit.excel.data;

namespace toolkit.excel.console
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            DataAccess da = new DataAccess(false);
            da.ProcessDefinitions();
        }
    }
}