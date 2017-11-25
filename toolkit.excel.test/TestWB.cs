using System;
using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using toolkit.excel.data;

namespace toolkit.excel.test
{
    [TestClass]
    public class TestWB
    {
        private readonly ExcelReader reader;
        private readonly DataTable result;

        public TestWB()
        {
            reader = new ExcelReader("TestWB.xlsx", "Sheet1", "A1:L500", true);
            result = reader.Read();
        }

        [TestMethod]
        public void ColumnCount()
        {
            Assert.AreEqual(12, result.Columns.Count);
        }

        [TestMethod]
        public void DataTypes()
        {
            Assert.AreEqual(typeof(string), result.Columns[0].DataType);
            Assert.AreEqual(typeof(long), result.Columns[1].DataType);
            Assert.AreEqual(typeof(string), result.Columns[2].DataType);
            Assert.AreEqual(typeof(long), result.Columns[3].DataType);
            Assert.AreEqual(typeof(long), result.Columns[4].DataType);
            Assert.AreEqual(typeof(string), result.Columns[5].DataType);
            Assert.AreEqual(typeof(string), result.Columns[6].DataType);
            Assert.AreEqual(typeof(DateTime), result.Columns[7].DataType);
            Assert.AreEqual(typeof(DateTime), result.Columns[8].DataType);
            Assert.AreEqual(typeof(bool), result.Columns[9].DataType);
            Assert.AreEqual(typeof(bool), result.Columns[10].DataType);
            Assert.AreEqual(typeof(bool), result.Columns[11].DataType);
        }
    }
}