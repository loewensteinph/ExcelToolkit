using System;
using System.Data;
using System.Data.SqlClient;

namespace toolkit.excel.data
{
    /// <summary>
    /// Helper Class to Sync DB Tables with DataTables
    /// </summary>  
    public static class TableHelper
    {
        /// <summary>
        /// Returns a Create Table Statement for a given DataTable
        /// </summary>  
        public static string CreateTable(string connectionString, string tableName, DataTable table)
        {
            if (table != null)
            {
                string sqlsc;
                using (var connection = new SqlConnection(connectionString))
                {
                    //connection.Open();
                    sqlsc = @"IF OBJECT_ID(N'" + tableName + "', N'U') IS NULL " +
                            "BEGIN" +
                            "\n";
                    sqlsc = sqlsc + "CREATE TABLE " + tableName + "(";
                    for (var i = 0; i < table.Columns.Count; i++)
                    {
                        sqlsc += "\n" + table.Columns[i].ColumnName;
                        if (table.Columns[i].DataType == typeof(long))
                            sqlsc += " INT ";
                        else if (table.Columns[i].DataType == typeof(DateTime))
                            sqlsc += " DATETIME ";
                        else if (table.Columns[i].DataType == typeof(string))
                            sqlsc += " NVARCHAR(" + table.Columns[i].MaxLength.ToString().Replace("-1", "MAX") + ") ";
                        else if (table.Columns[i].DataType == typeof(float))
                            sqlsc += " DECIMAL ";
                        else if (table.Columns[i].DataType == typeof(double))
                            sqlsc += " DECIMAL ";
                        else
                            sqlsc += " NVARCHAR(" + table.Columns[i].MaxLength.ToString().Replace("-1", "MAX") + ") ";
                        if (table.Columns[i].AutoIncrement)
                            sqlsc += " IDENTITY(" + table.Columns[i].AutoIncrementSeed + "," +
                                     table.Columns[i].AutoIncrementStep + ") ";
                        if (!table.Columns[i].AllowDBNull)
                            sqlsc += " NOT NULL ";
                        sqlsc = sqlsc + ",";
                    }

                    //connection.Close();
                }
                sqlsc = sqlsc + ")";
                sqlsc = sqlsc + "\nEND";
                return sqlsc;
            }
            return null;
        }
    }
}