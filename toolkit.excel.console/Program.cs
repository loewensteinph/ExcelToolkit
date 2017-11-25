using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using toolkit.excel.data;

namespace toolkit.excel.console
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelReader reader = new ExcelReader("","","",false);
            reader.Read();
        }
    }
}
