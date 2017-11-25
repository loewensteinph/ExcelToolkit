using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using log4net;
using OfficeOpenXml;

namespace toolkit.excel.data
{
    public class ExcelReader
    {
        private ILog Log = LogManager.GetLogger(typeof(ExcelReader));

        private readonly CultureInfo[] _cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);
        private readonly HashSet<string> _patterns;
        public ExcelDefinition Definition;
        private readonly DataTable _rawTbl = new DataTable();
        private readonly DataTable _finalTbl = new DataTable();

        public ExcelReader(string fileName, string sheetName, string range, bool hasHeaderRow)
        {
            Log = LogManager.GetLogger(typeof(ExcelReader));
            Log.Info(string.Format("Starting Import: {0}", fileName));

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

        private bool CheckFilePath(string fileName)
        {
            try
            {
                if (!File.Exists(fileName))
                {
                    File.Open(fileName, FileMode.Open);
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(string.Format("File {0} not found!", fileName), ex);
            }
            return false;
        }

        private bool CheckRange(ExcelWorksheet sheet)
        {
            try
            {
                var wsCol = sheet.Cells[Definition.Range];
            }
            catch (Exception ex)
            {
                Log.Error(
                    string.Format("Error accessing Range: '{0}' in Workbook: '{1}' Sheet: '{2}'", Definition.Range,
                        Definition.FileName, sheet.Name), ex);
                return false;
            }
            return true;
        }

        public DataTable Read()
        {
            if (!CheckFilePath(Definition.FileName))
                return null;

            using (var pck = new ExcelPackage())
            {
                ExcelWorksheet ws = GetExcelWorksheet(pck);

                if (ws == null)
                    ws = pck.Workbook.Worksheets[0];

                if (ws == null)
                    throw new NullReferenceException();

                if (!CheckRange(ws))
                    return null;
                ExcelRange wsCol = ws.Cells[Definition.Range];

                Int32 startCol = wsCol.Start.Column;
                Int32 endCol = wsCol.End.Column;
                Int32 startRow = wsCol.Start.Row;
                Int32 endRow = wsCol.End.Row;

                if (Definition.HasHeaderRow)
                {
                    startRow = startRow + 1;
                }

                if (Definition.HasHeaderRow)
                    foreach (var firstRowCell in ws.Cells[startRow - 1, startCol, startRow - 1, endCol])
                    {
                        DataColumn col = new DataColumn(firstRowCell.Text);
                        _rawTbl.Columns.Add(col);
                    }
                else
                    foreach (var firstRowCell in ws.Cells[startRow, startCol, startRow, endCol])
                    {
                        DataColumn col = new DataColumn("Col" + firstRowCell.Start.Column);
                        _rawTbl.Columns.Add(col);
                    }

                // Check Datatypes
                AddValuesToRawTable(startRow, endRow, _rawTbl, ws, startCol);

                GetDataTypes(_rawTbl);

                foreach (DataRow dr in _rawTbl.Rows)
                    _finalTbl.ImportRow(dr);
                return _finalTbl;
            }
        }

        private ExcelWorksheet GetExcelWorksheet(ExcelPackage pck)
        {
            using (FileStream stream = File.OpenRead(Definition.FileName))
            {
                pck.Load(stream);
            }

            var ws = pck.Workbook.Worksheets[Definition.SheetName];
            return ws;
        }

        private static void AddValuesToRawTable(Int32 startRow, Int32 endRow, DataTable rawTbl, ExcelWorksheet ws, Int32 startCol)
        {
            for (Int32 rowNum = startRow; rowNum <= endRow; rowNum++)
            {
                DataRow row = rawTbl.NewRow();

                Int32 endColumn;

                if (rawTbl.Columns.Count == 1)
                    endColumn = 1;
                else
                {
                    endColumn = rawTbl.Columns.Count + 1;
                }

                ExcelRange wsRow = ws.Cells[rowNum, startCol, rowNum, endColumn];

                Int32 cellsWithoutContent = 0;
                foreach (ExcelRangeBase cell in wsRow)
                {
                    if (String.IsNullOrEmpty(cell.Text) || String.IsNullOrWhiteSpace(cell.Text))
                        cellsWithoutContent++;
                }
                if (rawTbl.Columns.Count - cellsWithoutContent > 0)
                {
                    foreach (var cell in wsRow)
                    {
                        DataColumn dataColumn = rawTbl.Columns[cell.Start.Column - 1];
                        row[dataColumn.ColumnName] = cell.Value;
                        if (cell.Text.Equals("NULL"))
                            row[cell.Start.Column - 1] = DBNull.Value;
                        else
                            row[cell.Start.Column - 1] = cell.Text;
                    }
                    rawTbl.Rows.Add(row);
                }
            }
            rawTbl.Rows.Cast<DataRow>().ToList().FindAll(row => String.IsNullOrEmpty(String.Join("", row.ItemArray))).ForEach(Row =>
                { rawTbl.Rows.Remove(Row); });
        }

        private void GetDataTypes(DataTable rawTbl)
        {
            foreach (DataColumn col in rawTbl.Columns)
            {
                object currentDataType = typeof(string);
                bool first = true;

                foreach (DataRow row in rawTbl.Rows)
                {
                    if (!String.IsNullOrEmpty(row.ToString()))
                    {
                        if (currentDataType != ParseString(row[col.ColumnName].ToString()) && !first)
                        {
                            if (currentDataType.Equals(typeof(decimal)) &&
                                ParseString(row[col.ColumnName].ToString()).Equals(typeof(long)))
                                break;
                            currentDataType = typeof(string);
                            break;
                        }
                        currentDataType = ParseString(row[col.ColumnName].ToString());
                    }
                    first = false;
                }
                DataColumn finalColumn = new DataColumn
                {
                    DataType = currentDataType as Type,
                    ColumnName = col.ColumnName
                };
                if (!_finalTbl.Columns.Contains(finalColumn.ColumnName))
                    _finalTbl.Columns.Add(finalColumn);
            }
        }

        public static bool IsGuid(string value)
        {
            Guid x;
            return Guid.TryParse(value, out x);
        }

        public object ParseString(string str)
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