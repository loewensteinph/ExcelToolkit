using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using log4net;

namespace toolkit.excel.data
{
    /// <summary>Interacts with a given DataBase</summary>
    public class DataAccess
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(DataAccess));

        private List<ExcelDefinition> _excelDefinitions;

        /// <summary>Default Constructor</summary>
        public DataAccess()
        {
            GetDefinitions();
        }
        /// <summary>Unit Test Constructor</summary>
        public DataAccess(bool testMode)
        {
            GetTestDefinitions();
        }
        /// <summary>Unit Test Constructor</summary>
        public void ProcessDefinitions()
        {
            foreach (var definition in _excelDefinitions)
            {
                var reader = new ExcelReader(definition.FileName, definition.SheetName, definition.Range, true);
                var result = reader.Read();
                SaveData(result, definition);
            }
        }
        /// <summary>Persists Data in DataTable to Sql DB</summary>
        /// <param name="excelDefinition"></param>
        /// <param name="srcDataTable"></param>
        public static void SaveData(DataTable srcDataTable, ExcelDefinition excelDefinition)
        {
            Log.Info(string.Format("Saving {0} Row(s)", srcDataTable.Rows.Count));
            foreach (DataRow row in srcDataTable.Rows)
            InsertDataRow(row, excelDefinition);
        }
        /// <summary>Creates SqlParameter for Insert Command</summary>
        /// <param name="sqlCommand"></param>
        /// <param name="parameterName"></param>
        /// <param name="sourceColumn"></param>
        /// <param name="insertValue"></param>
        public static void InsertParameter(SqlCommand sqlCommand,
            string parameterName,
            string sourceColumn,
            object insertValue)
        {
            var parameter = new SqlParameter(parameterName, insertValue);

            parameter.Direction = ParameterDirection.Input;
            parameter.ParameterName = parameterName;
            parameter.SourceColumn = sourceColumn;
            parameter.SourceVersion = DataRowVersion.Current;

            sqlCommand.Parameters.Add(parameter);
        }
        /// <summary>Returns the Insert Statement</summary>
        /// <param name="srcDataTable"></param>
        /// <param name="excelDefinition"></param>
        public static string BuildInsertSql(DataTable srcDataTable, ExcelDefinition excelDefinition)
        {
            srcDataTable.TableName = excelDefinition.TargetTable;

            var sql = new StringBuilder("INSERT INTO " + srcDataTable.TableName + " (");
            var values = new StringBuilder("VALUES (");
            var bFirst = true;
            var bIdentity = false;
            string identityType = null;

            if (excelDefinition.ColumnMappings.Count > 0)
            {
                var sourceCols = string.Join(",", excelDefinition.ColumnMappings.Select(x => "@" + x.SourceColumn));
                var targetCols = string.Join(",", excelDefinition.ColumnMappings.Select(x => x.TargetColumn));
                sql.Append(targetCols);
                values.Append(sourceCols);
                sql.Append(") ");
                sql.Append(values);
                sql.Append(")");
                return sql.ToString();
                ;
            }
            foreach (DataColumn column in srcDataTable.Columns)
                if (column.AutoIncrement)
                {
                    bIdentity = true;

                    switch (column.DataType.Name)
                    {
                        case "Int16":
                            identityType = "smallint";
                            break;
                        case "SByte":
                            identityType = "tinyint";
                            break;
                        case "Int64":
                            identityType = "bigint";
                            break;
                        case "Decimal":
                            identityType = "decimal";
                            break;
                        default:
                            identityType = "int";
                            break;
                    }
                }
                else
                {
                    if (bFirst)
                    {
                        bFirst = false;
                    }
                    else
                    {
                        sql.Append(", ");
                        values.Append(", ");
                    }

                    sql.Append(column.ColumnName);
                    values.Append("@");
                    values.Append(column.ColumnName);
                }
            sql.Append(") ");
            sql.Append(values);
            sql.Append(")");

            if (bIdentity)
            {
                sql.Append("; SELECT CAST(scope_identity() AS ");
                sql.Append(identityType);
                sql.Append(")");
            }
            return sql.ToString();
            ;
        }
        /// <summary>Builds the final SqlCommand</summary>
        /// <param name="dataRow"></param>
        /// <param name="excelDefinition"></param>
        public static SqlCommand CreateInsertCommand(DataRow dataRow, ExcelDefinition excelDefinition)
        {
            var table = dataRow.Table;
            var sql = BuildInsertSql(table, excelDefinition);
            var command = new SqlCommand(sql);
            command.CommandType = CommandType.Text;

            if (excelDefinition.ColumnMappings.Count > 0)
            {
                foreach (var mapping in excelDefinition.ColumnMappings)
                {
                    var parameterName = "@" + mapping.SourceColumn;
                    InsertParameter(command, parameterName,
                        mapping.SourceColumn,
                        dataRow[mapping.SourceColumn]);
                }
                return command;
            }

            foreach (DataColumn column in table.Columns)
                if (!column.AutoIncrement)
                {
                    var parameterName = "@" + column.ColumnName;
                    InsertParameter(command, parameterName,
                        column.ColumnName,
                        dataRow[column.ColumnName]);
                }
            return command;
        }
        /// <summary>Executes Insert Command for each row</summary>
        /// <param name="dataRow"></param>
        /// <param name="excelDefinition"></param>
        public static void InsertDataRow(DataRow dataRow, ExcelDefinition excelDefinition)
        {
            var command = CreateInsertCommand(dataRow, excelDefinition);

            using (var connection = new SqlConnection(excelDefinition.ConnectionString))
            {
                command.Connection = connection;
                command.CommandType = CommandType.Text;
                connection.Open();
                try
                {
                    command.ExecuteScalar();
                    connection.Close();
                }
                catch (Exception ex)
                {
                    Log.Error(String.Format("Error Inserting Row {0} Table {1}", dataRow.Table.Rows.IndexOf(dataRow),dataRow.Table.TableName),ex);
                    connection.Close();
                }
            }
        }
        /// <summary>Returns all defined Excel Definitions from DB Context</summary>
        public void GetDefinitions()
        {
            using (var context = new ExcelDataContext())
            {
                _excelDefinitions = context.ExcelDefinition.Include(u => u.ColumnMappings).ToList();
                Log.Info(string.Format("Found {0} Definition(s)", _excelDefinitions.Count));
            }
        }
        /// <summary>Returns all defined Excel Definitions from Unittest DB Context</summary>
        public void GetTestDefinitions()
        {
            using (var context = new ExcelUnitTestDataContext())
            {
                _excelDefinitions = context.ExcelDefinition.Include(u => u.ColumnMappings).ToList();
                Log.Info(string.Format("Found {0} Definition(s)", _excelDefinitions.Count));
            }
        }
    }
}