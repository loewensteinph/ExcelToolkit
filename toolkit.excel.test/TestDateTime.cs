using System;
using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using toolkit.excel.data;

namespace toolkit.excel.test
{
    [TestClass]
    public class TestDateTime
    {
        private readonly ExcelReader reader;
        private readonly DataTable result;
        public TestDateTime()
        {
           reader = new ExcelReader("TestWBDateTime.xlsx", "Sheet1", "A1:A5", false);
           result = reader.Read();
        }
        [TestMethod]
        public void ColumnCount()
        {
            Assert.AreEqual(1,result.Columns.Count);
        }
        [TestMethod]
        public void DataTypes()
        {
            Assert.AreEqual(typeof(DateTime), result.Columns[0].DataType);
        }
    }
}
