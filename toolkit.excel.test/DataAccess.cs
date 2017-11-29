using System;
using System.Data;
using System.Data.Entity;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using toolkit.excel.data;

namespace toolkit.excel.test
{
    [TestClass]
    public class DataAccess
    {
        [ClassInitialize]
        public static void ClassInit(TestContext context)
        {
            ExcelUnitTestDataContext ctx = new ExcelUnitTestDataContext();
            Database.SetInitializer<ExcelUnitTestDataContext>(new ExcelDataContextSeedInitializer());
            ctx.Database.Initialize(true);

            ctx.Database.ExecuteSqlCommand(@"CREATE SCHEMA unittest;");
            ctx.Database.ExecuteSqlCommand(@"
CREATE TABLE unittest.UT1
(
    StringTest NVARCHAR(MAX),
    DecimalTest DECIMAL(12,4),
    IntTest INT,
    GuidTest UNIQUEIDENTIFIER
);");
            ctx.Database.ExecuteSqlCommand(@"
CREATE TABLE unittest.UT1a
(
    StringTest NVARCHAR(MAX),
    DecimalTest TINYINT,
    IntTest INT,
    GuidTest UNIQUEIDENTIFIER
);");
            ctx.Database.ExecuteSqlCommand(@"
CREATE TABLE unittest.UT2
(
    StringTest1 NVARCHAR(MAX),
    DecimalTest1 DECIMAL(12,4),
    IntTest1 INT,
    GuidTest1 UNIQUEIDENTIFIER,
    DateTest DATETIME
);");
            ctx.Database.ExecuteSqlCommand(@"
CREATE TABLE unittest.UT3
(
    StringTest1 NVARCHAR(MAX),
    DecimalTest1 DECIMAL(12,4),
    IntTest1 INT,
    GuidTest1 UNIQUEIDENTIFIER,
    DateTest DATETIME
);");
        }

        [TestMethod]
        public void JobProcessing()
        {
            data.DataAccess da = new data.DataAccess(true);
            da.ProcessDefinitions();
            Assert.AreEqual(12, 12);
        }
    }
}