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
    /// <summary>
    /// Reads a given Excel Workbook and converts it into a Datatable
    /// </summary>  
    public class ExcelReader
    {
        private ILog Log = LogManager.GetLogger(typeof(ExcelReader));

        private readonly CultureInfo[] _cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);
        private readonly HashSet<string> _patterns;
        private ExcelDefinition Definition;
        private readonly DataTable rawDataTable = new DataTable();
        private readonly DataTable _finalDataTable = new DataTable();

        /// <summary>Constructor Method</summary>
        /// <param name="fileName">Location of Excel File</param>
        /// <param name="sheetName">Name of Worksheet</param>
        /// <param name="range">Excel Range i.e. A1:C5</param>
        /// <param name="hasHeaderRow">Determines existance of Header Row</param>
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

            _patterns = new HashSet<string>();

            foreach (var culture in _cultures.Where(c=>c.Name == "en-US" || c.Name == "de-DE"))
            {
                _patterns.UnionWith(culture.DateTimeFormat.GetAllDateTimePatterns('d'));
                _patterns.UnionWith(culture.DateTimeFormat.GetAllDateTimePatterns('D'));
                _patterns.UnionWith(culture.DateTimeFormat.GetAllDateTimePatterns('f'));
                _patterns.UnionWith(culture.DateTimeFormat.GetAllDateTimePatterns('U'));
                _patterns.UnionWith(culture.DateTimeFormat.GetAllDateTimePatterns('g'));
                _patterns.UnionWith(culture.DateTimeFormat.GetAllDateTimePatterns('G'));
            }
        }
        /// <summary>Checks existance of an Excel File</summary>
        /// <param name="fileName">Location of Excel File</param>
        private bool CheckFilePath(string fileName)
        {
            try
            {
                if (!File.Exists(fileName))
                {
                    File.Open(fileName, FileMode.Open);
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(string.Format("File {0} not found!", fileName), ex);
            }
            return false;
        }
        /// <summary>Checks existance of a defined range</summary>
        /// <param name="workSheet">Excel Worksheet to check</param>
        private bool CheckDefinedRange(ExcelWorksheet workSheet)
        {
            try
            {
                var wsCol = workSheet.Cells[Definition.Range];
            }
            catch (Exception ex)
            {
                Log.Error(
                    string.Format("Error accessing Range: '{0}' in Workbook: '{1}' Sheet: '{2}'", Definition.Range,
                        Definition.FileName, workSheet.Name), ex);
                return false;
            }
            return true;
        }
        /// <summary>Creates a DataTable based on a given <see cref="ExcelDefinition"/></summary>
        public DataTable Read()
        {
            if (!CheckFilePath(Definition.FileName))
                return null;

            using (var pck = new ExcelPackage())
            {
                ExcelWorksheet ws = GetExcelWorksheet(pck);

                if (ws == null)
                    return null;

                if (!CheckDefinedRange(ws))
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
                        rawDataTable.Columns.Add(col);
                    }
                else
                    foreach (var firstRowCell in ws.Cells[startRow, startCol, startRow, endCol])
                    {
                        DataColumn col = new DataColumn("Col" + firstRowCell.Start.Column);
                        rawDataTable.Columns.Add(col);
                    }
                AddValuesToRawTable(startRow, endRow, rawDataTable, ws, startCol);
                GetDataTypes(rawDataTable);

                foreach (DataRow dr in rawDataTable.Rows)
                    _finalDataTable.ImportRow(dr);
                return _finalDataTable;
            }
        }
        /// <summary>Resturns a certain Worksheet from a given Excel File</summary>
        /// <param name="excelPackage">Excelpackage</param>
        private ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage)
        {
            ExcelWorksheet ws;
            using (FileStream stream = File.OpenRead(Definition.FileName))
            {
                excelPackage.Load(stream);
            }
            try
            {
                ws = excelPackage.Workbook.Worksheets[Definition.SheetName];
            }
            catch (Exception ex)
            {
                Log.Error("Worksheet not found!",ex);
                return null;
            }

            return ws;
        }
        /// <summary>Adds raw values from <paramref name="worksheet"/> to <paramref name="rawDataTable"/></summary>
        /// <param name="worksheetStartRow">Data Range Start Row</param>
        /// <param name="worksheetEndRow">Data Range End Row</param>
        /// <param name="rawDataTable">Raw DataTable</param>
        /// <param name="worksheet">Excel Worksheet to process</param>
        /// <param name="worksheetStartCol">First Column to process</param> 
        private static void AddValuesToRawTable(Int32 worksheetStartRow, Int32 worksheetEndRow, DataTable rawDataTable, ExcelWorksheet worksheet, Int32 worksheetStartCol)
        {
            for (Int32 rowNum = worksheetStartRow; rowNum <= worksheetEndRow; rowNum++)
            {
                DataRow row = rawDataTable.NewRow();

                Int32 endColumn;

                if (rawDataTable.Columns.Count == 1)
                    endColumn = 1;
                else
                {
                    endColumn = rawDataTable.Columns.Count + 1;
                }

                ExcelRange wsRow = worksheet.Cells[rowNum, worksheetStartCol, rowNum, endColumn];

                Int32 cellsWithoutContent = 0;
                foreach (ExcelRangeBase cell in wsRow)
                {
                    if (String.IsNullOrEmpty(cell.Text) || String.IsNullOrWhiteSpace(cell.Text))
                        cellsWithoutContent++;
                }
                if (rawDataTable.Columns.Count - cellsWithoutContent > 0)
                {
                    foreach (var cell in wsRow)
                    {
                        DataColumn dataColumn = rawDataTable.Columns[cell.Start.Column - 1];
                        row[dataColumn.ColumnName] = cell.Value;
                        if (cell.Text.Equals("NULL"))
                            row[cell.Start.Column - 1] = DBNull.Value;
                        else
                            row[cell.Start.Column - 1] = cell.Text;
                    }
                    rawDataTable.Rows.Add(row);
                }
            }
            rawDataTable.Rows.Cast<DataRow>().ToList().FindAll(row => String.IsNullOrEmpty(String.Join("", row.ItemArray))).ForEach(Row =>
                { rawDataTable.Rows.Remove(Row); });
        }
        /// <summary>Sets DataTypes for Columns of <see name="_finalDataTable"/></summary>
        /// <param name="rawDataTable">Raw DataTable</param>
        private void GetDataTypes(DataTable rawDataTable)
        {
            foreach (DataColumn col in rawDataTable.Columns)
            {
                object currentDataType = typeof(string);
                bool first = true;

                foreach (DataRow row in rawDataTable.Rows)
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
                    Log.Info(String.Format("Column {0} Row {1} DataType {2} identified Value {3}", col.ColumnName,row.Table.Rows.IndexOf(row), currentDataType, row[col.ColumnName].ToString()));
                }
                DataColumn finalColumn = new DataColumn
                {
                    DataType = currentDataType as Type,
                    ColumnName = col.ColumnName
                };
                if (!_finalDataTable.Columns.Contains(finalColumn.ColumnName))
                    _finalDataTable.Columns.Add(finalColumn);
            }
        }
        /// <summary>Identifies if a given String <paramref name="stringToParse"/> is of Type <see cref="System.Guid"/>></summary>
        /// <param name="stringToParse">String to parse</param>
        public static bool IsGuid(string stringToParse)
        {
            Guid x;
            return Guid.TryParse(stringToParse, out x);
        }
        /// <summary>Identifies the DataType of a given String <paramref name="stringToParse"/></summary>
        /// <param name="stringToParse">String to parse</param>
        public object ParseString(string stringToParse)
        {
            long intValue;
            decimal decimalValue;
            Guid guidValue;
            bool boolValue;
            DateTime datetimeValue;

            CultureInfo cultureinfo = new CultureInfo("de-DE");
            CultureInfo cultureinfoUs = new CultureInfo("en-US");

            if (long.TryParse(stringToParse, out intValue))
                return intValue.GetType();
            if (DateTime.TryParseExact(stringToParse, _patterns.ToArray(), cultureinfo,
                DateTimeStyles.None, out datetimeValue))
                return datetimeValue.GetType();
            if (DateTime.TryParseExact(stringToParse, _patterns.ToArray(), cultureinfoUs,
                DateTimeStyles.None, out datetimeValue))
                return datetimeValue.GetType();
            if (decimal.TryParse(stringToParse, out decimalValue))
                return decimalValue.GetType();
            if (DateTime.TryParse(stringToParse, out datetimeValue))
                return datetimeValue.GetType();
            if (Guid.TryParse(stringToParse, out guidValue))
                return guidValue.GetType();
            if (bool.TryParse(stringToParse, out boolValue))
                return boolValue.GetType();
            return typeof(string);
        }
    }
}