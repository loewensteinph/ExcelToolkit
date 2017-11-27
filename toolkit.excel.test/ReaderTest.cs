using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using toolkit.excel.data;

namespace toolkit.excel.test
{
    [TestClass]
    public class TestReader
    {
        private ExcelReader _reader;
        private DataTable _result;

        [TestMethod]
        public void InvailidFilename()
        {
            _reader = new ExcelReader(@"TestWorkbooks\X.xlsx", "Sheet1", "A1:A5", false);
            _result = _reader.Read();
            Assert.AreEqual(null, _result);
        }

        [TestMethod]
        public void InvalidRange()
        {
            _reader = new ExcelReader(@"TestWorkbooks\TestWB.xlsx", "Sheet1", "A1:A-1", false);
            _result = _reader.Read();
            Assert.AreEqual(null, _result);
        }
        [TestMethod]
        public void MissingSheetName()
        {
            _reader = new ExcelReader(@"TestWorkbooks\TestWB.xlsx", "", "A1:A5", false);
            _result = _reader.Read();
            Assert.AreEqual(null, _result);
        }
        [TestMethod]
        public void SingleColumn()
        {
            _reader = new ExcelReader(@"TestWorkbooks\TestWB.xlsx", "Sheet1", "A2:A5", true);
            _result = _reader.Read();
            Assert.AreEqual(3, _result.Rows.Count);
        }
        [TestMethod]
        public void SingleCell()
        {
            _reader = new ExcelReader(@"TestWorkbooks\TestWB.xlsx", "Sheet1", "A2:A2", false);
            _result = _reader.Read();
            Assert.AreEqual(1, _result.Rows.Count);
        }
        [TestMethod]
        public void BigSheet()
        {
            _reader = new ExcelReader(@"TestWorkbooks\TestWBBig.xlsx", "Sheet1", "A1:L5500", true);
            _reader.Exceldefinition.ValidateDataTypes = false;
            _result = _reader.Read();
            Assert.AreEqual(5499, _result.Rows.Count);
        }

        [TestMethod]
        public void BigSheet2()
        {
            _reader = new ExcelReader(@"TestWorkbooks\TestWBBig.xlsx", "Sheet1", "A1:L40000", true);
            _reader.Exceldefinition.ValidateDataTypes = false;
            _result = _reader.Read();        
            Assert.AreEqual(39999, _result.Rows.Count);
        }
    }
}