using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace toolkit.excel.data
{
    public class ExcelReader
    {
        private static CultureInfo[] _cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);
        private static HashSet<string> _patterns;
        public ExcelDefinition Definition;

        public ExcelReader(string fileName, string sheetName, string range, bool hasHeaderRow)
        {
            Definition = new ExcelDefinition
            {
                Range = range,
                SheetName = sheetName,
                FileName = fileName,
                HasHeaderRow = hasHeaderRow
            };
            foreach (var culture in _cultures)
            {
                _patterns = new HashSet<string>();
                _cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);
                _patterns.UnionWith(culture.DateTimeFormat.GetAllDateTimePatterns());
            }
        }

        public DataTable Read()
        {
            var datatypes = new Dictionary<string, object>();

            using (var pck = new ExcelPackage())
            {
                using (var stream = File.OpenRead(Definition.FileName))
                {
                    pck.Load(stream);
                }

                var ws = pck.Workbook.Worksheets[Definition.SheetName];

                if (ws == null)
                    ws = pck.Workbook.Worksheets[0];

                if (ws == null)
                    throw new NullReferenceException();

                var rawTbl = new DataTable();
                var finalTbl = new DataTable();

                var wsCol = ws.Cells[Definition.Range];          
                var startCol = wsCol.Start.Column;
                var endCol = wsCol.End.Column;
                var startRow = wsCol.Start.Row;
                var endRow = wsCol.End.Row;

                if (Definition.HasHeaderRow)
                {
                    startRow = startRow + 1;
                    endRow = endRow + 1;
                }

                if (Definition.HasHeaderRow)
                {
                    foreach (var firstRowCell in ws.Cells[startRow -1, startCol, startRow -1, endCol])
                    {
                        var col = new DataColumn(firstRowCell.Text);
                        rawTbl.Columns.Add(col);
                    }
                }
                else
                {
                    foreach (var firstRowCell in ws.Cells[startRow, startCol, startRow, endCol])
                    {
                        var col = new DataColumn("Col" + firstRowCell.Start.Column);
                        rawTbl.Columns.Add(col);
                    }
                }

                // Check Datatypes

                foreach (var col in rawTbl.Columns)
                {
                    object currentDataType = typeof(string);
                    var first = true;

                    DataColumn curColumn = (DataColumn)col;
                    int curColumnIndex = rawTbl.Columns.IndexOf(curColumn) + 1;
                    var wsRow = ws.Cells[startRow, curColumnIndex, endRow, curColumnIndex];

                    foreach (var wsCell in wsRow)
                    {
                        if (!wsCell.Text.Equals(string.Empty))
                        {
                            if (currentDataType != ParseString(wsCell.Text) && !first)
                            {
                                if (currentDataType.Equals(typeof(decimal)) &&
                                    ParseString(wsCell.Text).Equals(typeof(long)))
                                    break;
                                currentDataType = typeof(string);
                                break;
                            }
                            currentDataType = ParseString(wsCell.Text);
                        }
                        first = false;
                    }
                    datatypes.Add(curColumn.ColumnName, currentDataType);

                }

                for (var rowNum = startRow; rowNum <= endRow; rowNum++)
                {
                    var row = rawTbl.NewRow();

                    var wsRow = ws.Cells[rowNum, startCol, rowNum, rawTbl.Columns.Count + 1];
                    var hascontent = false;
                    foreach (var cell in wsRow)
                        if (!cell.Text.Equals(string.Empty))
                            hascontent = true;
                    if (hascontent)
                        foreach (var cell in wsRow)
                        {
                            object colType;
                            DataColumn dataColumn = rawTbl.Columns[cell.Start.Column - 1];


                            datatypes.TryGetValue(dataColumn.ColumnName,
                                out colType);
                            if (!finalTbl.Columns.Contains(dataColumn.ColumnName))
                                finalTbl.Columns.Add(dataColumn.ColumnName);
                            finalTbl.Columns[cell.Start.Column - 1].DataType = colType as Type;
                            row[dataColumn.ColumnName] = cell.Value;
                            if (cell.Text.Equals("NULL"))
                                row[cell.Start.Column - 1] = DBNull.Value;
                            else
                                row[cell.Start.Column - 1] = cell.Text;
                        }
                    rawTbl.Rows.Add(row);
                }
                foreach (DataRow dr in rawTbl.Rows)
                    finalTbl.ImportRow(dr);
                return finalTbl;
            }
        }

        public static bool IsGuid(string value)
        {
            Guid x;
            return Guid.TryParse(value, out x);
        }

        public static object ParseString(string str)
        {
            long intValue;
            decimal decimalValue;
            Guid guidValue;
            bool boolValue;
            DateTime datetimeValue;
            // Place checks higher if if-else statement to give higher priority to type.
            if (long.TryParse(str, out intValue))
                return intValue.GetType();
            if (DateTime.TryParseExact(str, _patterns.ToArray(), CultureInfo.InvariantCulture,
                DateTimeStyles.None, out datetimeValue))
                return datetimeValue.GetType();
            if (DateTime.TryParse(str, out datetimeValue))
                return datetimeValue.GetType();
            if (decimal.TryParse(str, out decimalValue))
                return decimalValue.GetType();
            if (Guid.TryParse(str, out guidValue))
                return guidValue.GetType();
            if (bool.TryParse(str, out boolValue))
                return boolValue.GetType();
            return typeof(string);
        }
    }
}