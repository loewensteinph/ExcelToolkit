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
    public class DataAccess
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(DataAccess));

        private List<ExcelDefinition> _excelDefinitions;

        public DataAccess()
        {
            GetDefinitions();
        }

        public DataAccess(bool testMode)
        {
            GetTestDefinitions();
        }

        public void ProcessDefinitions()
        {
            foreach (var definition in _excelDefinitions)
            {
                var reader = new ExcelReader(definition.FileName, definition.SheetName, definition.Range, true);
                var result = reader.Read();
                SaveData(result, definition);
            }
        }

        public static void SaveData(DataTable srcTable, ExcelDefinition def)
        {
            Log.Info(string.Format("Saving {0} Row(s)", srcTable.Rows.Count));
            foreach (DataRow row in srcTable.Rows)
            InsertDataRow(row, def);
        }

        public static void InsertParameter(SqlCommand command,
            string parameterName,
            string sourceColumn,
            object value)
        {
            var parameter = new SqlParameter(parameterName, value);

            parameter.Direction = ParameterDirection.Input;
            parameter.ParameterName = parameterName;
            parameter.SourceColumn = sourceColumn;
            parameter.SourceVersion = DataRowVersion.Current;

            command.Parameters.Add(parameter);
        }

        public static string BuildInsertSql(DataTable table, ExcelDefinition def)
        {
            table.TableName = def.TargetTable;

            var sql = new StringBuilder("INSERT INTO " + table.TableName + " (");
            var values = new StringBuilder("VALUES (");
            var bFirst = true;
            var bIdentity = false;
            string identityType = null;

            if (def.ColumnMappings.Count > 0)
            {
                var sourceCols = string.Join(",", def.ColumnMappings.Select(x => "@" + x.SourceColumn));
                var targetCols = string.Join(",", def.ColumnMappings.Select(x => x.TargetColumn));
                sql.Append(targetCols);
                values.Append(sourceCols);
                sql.Append(") ");
                sql.Append(values);
                sql.Append(")");
                return sql.ToString();
                ;
            }

            foreach (DataColumn column in table.Columns)
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

        public static SqlCommand CreateInsertCommand(DataRow row, ExcelDefinition def)
        {
            var table = row.Table;
            var sql = BuildInsertSql(table, def);
            var command = new SqlCommand(sql);
            command.CommandType = CommandType.Text;

            if (def.ColumnMappings.Count > 0)
            {
                foreach (var mapping in def.ColumnMappings)
                {
                    var parameterName = "@" + mapping.SourceColumn;
                    InsertParameter(command, parameterName,
                        mapping.SourceColumn,
                        row[mapping.SourceColumn]);
                }
                return command;
            }

            foreach (DataColumn column in table.Columns)
                if (!column.AutoIncrement)
                {
                    var parameterName = "@" + column.ColumnName;
                    InsertParameter(command, parameterName,
                        column.ColumnName,
                        row[column.ColumnName]);
                }
            return command;
        }

        public static void InsertDataRow(DataRow row, ExcelDefinition def)
        {
            var command = CreateInsertCommand(row, def);

            using (var connection = new SqlConnection(def.ConnectionString))
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
                    Log.Error(String.Format("Error Inserting Row {0} Table {1}", row.Table.Rows.IndexOf(row),row.Table.TableName),ex);
                    connection.Close();
                }

            }
        }

        public void GetDefinitions()
        {
            using (var context = new ExcelDataContext())
            {
                _excelDefinitions = context.ExcelDefinition.Include(u => u.ColumnMappings).ToList();
                Log.Info(string.Format("Found {0} Definition(s)", _excelDefinitions.Count));
            }
        }
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