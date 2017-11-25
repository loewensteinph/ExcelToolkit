using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using toolkit.excel.data;

namespace toolkit.excel.test
{
    [TestClass]
    public class TestReader
    {
        private ExcelReader reader;
        private DataTable result;

        [TestMethod]
        public void InvailidFilename()
        {
            reader = new ExcelReader("X.xlsx", "Sheet1", "A1:A5", false);
            result = reader.Read();
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        public void InvalidRange()
        {
            reader = new ExcelReader("TestWB.xlsx", "Sheet1", "A1:A-1", false);
            result = reader.Read();
            Assert.AreEqual(null, result);
        }
        [TestMethod]
        public void SingleColumn()
        {
            reader = new ExcelReader("TestWB.xlsx", "Sheet1", "A2:A5", true);
            result = reader.Read();
            Assert.AreEqual(3, result.Rows.Count);
        }
        [TestMethod]
        public void SingleCell()
        {
            reader = new ExcelReader("TestWB.xlsx", "Sheet1", "A2:A2", false);
            result = reader.Read();
            Assert.AreEqual(1, result.Rows.Count);
        }
        [TestMethod]
        public void BigSheet()
        {
            reader = new ExcelReader("TestWBBig.xlsx", "Sheet1", "A1:L5500", true);
            result = reader.Read();
            Assert.AreEqual(5499, result.Rows.Count);
        }
        [TestMethod]
        public void BigSheet2()
        {
            reader = new ExcelReader("TestWBBig.xlsx", "Sheet1", "A1:L40000", true);
            result = reader.Read();
            Assert.AreEqual(39999, result.Rows.Count);
        }
    }
}