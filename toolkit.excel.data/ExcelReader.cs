using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using log4net;
using OfficeOpenXml;
using System.Threading;

namespace toolkit.excel.data
{
    /// <summary>
    /// Reads a given Excel Workbook and converts it into a Datatable
    /// </summary>  
    public class ExcelReader : IDisposable
    {
        private ILog Log = LogManager.GetLogger(typeof(ExcelReader));

        private readonly CultureInfo[] _cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);
        private HashSet<string> _patterns;
        public ExcelDefinition Exceldefinition;
        private readonly DataTable _rawDataTable = new DataTable();
        private readonly DataTable _finalDataTable = new DataTable();

        private ExcelWorksheet excelWorksheet;

        private Int32 worksheetStartColumn;
        private Int32 worksheetEndColumn;
        private Int32 worksheetStartRow;
        private Int32 worksheetEndRow;

        /// <summary>Constructor Method</summary>
        /// <param name="definition">Excel Definition</param>
        public ExcelReader(ExcelDefinition definition)
        {
            initializeCulture();
            Exceldefinition = definition;
            Log = LogManager.GetLogger(typeof(ExcelReader));
            Log.Info(string.Format("Starting Import: {0}", Exceldefinition.FileName));
        }

        /// <summary>Constructor Method</summary>
        /// <param name="fileName">Location of Excel File</param>
        /// <param name="sheetName">Name of Worksheet</param>
        /// <param name="range">Excel Range i.e. A1:C5</param>
        /// <param name="hasHeaderRow">Determines existance of Header Row</param>
        public ExcelReader(string fileName, string sheetName, string range, bool hasHeaderRow)
        {
            initializeCulture();
            Log = LogManager.GetLogger(typeof(ExcelReader));
            Log.Info(string.Format("Starting Import: {0}", fileName));

            Exceldefinition = new ExcelDefinition
            {
                Range = range,
                SheetName = sheetName,
                FileName = fileName,
                HasHeaderRow = hasHeaderRow
            };
        }
        private void initializeCulture()
        {
            _patterns = new HashSet<string>();

            foreach (var culture in _cultures.Where(c => c.Name == "en-US" || c.Name == "de-DE"))
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
                var wsCol = workSheet.Cells[Exceldefinition.Range];
            }
            catch (Exception ex)
            {
                Log.Error(
                    string.Format("Error accessing Range: '{0}' in Workbook: '{1}' Sheet: '{2}'", Exceldefinition.Range,
                        Exceldefinition.FileName, workSheet.Name), ex);
                return false;
            }
            return true;
        }
        /// <summary>Creates a DataTable based on a given <see cref="ExcelDefinition"/></summary>
        public DataTable Read()
        {
            if (!CheckFilePath(Exceldefinition.FileName))
                return null;

            using (var pck = new ExcelPackage())
            {
                excelWorksheet = GetExcelWorksheet(pck);

                if (excelWorksheet == null)
                    return null;

                if (!CheckDefinedRange(excelWorksheet))
                    return null;
                ExcelRange wsCol = excelWorksheet.Cells[Exceldefinition.Range];

                worksheetStartColumn = wsCol.Start.Column;
                worksheetEndColumn = wsCol.End.Column;
                worksheetStartRow = wsCol.Start.Row;
                worksheetEndRow = wsCol.End.Row;

                if (Exceldefinition.RangeWidthAuto)
                {
                    worksheetEndColumn = excelWorksheet.Dimension.End.Column;
                }
                if (Exceldefinition.RangeHeightAuto)
                {
                    worksheetEndRow = excelWorksheet.Dimension.End.Row;
                }

                if (Exceldefinition.HasHeaderRow)
                {
                    worksheetStartRow = worksheetStartRow + 1;
                }

                if (Exceldefinition.HasHeaderRow)
                    foreach (var firstRowCell in excelWorksheet.Cells[worksheetStartRow - 1, worksheetStartColumn, worksheetStartRow - 1, worksheetEndColumn])
                    {
                        DataColumn col = new DataColumn(firstRowCell.Text);
                        _rawDataTable.Columns.Add(col);
                    }
                else
                    foreach (var firstRowCell in excelWorksheet.Cells[worksheetStartRow, worksheetStartColumn, worksheetStartRow, worksheetEndColumn])
                    {
                        DataColumn col = new DataColumn("Col" + firstRowCell.Start.Column);
                        _rawDataTable.Columns.Add(col);
                    }

                AddValuesToRawTable();

                if (Exceldefinition.ValidateDataTypes)
                {
                    AddTypedColumns();
                }
                else
                {
                    AddRawColumns();
                }

                foreach (DataRow dr in _rawDataTable.Rows)
                    _finalDataTable.ImportRow(dr);
                return _finalDataTable;
            }
        }

        /// <summary>Resturns a certain Worksheet from a given Excel File</summary>
        /// <param name="excelPackage">Excelpackage</param>
        private ExcelWorksheet GetExcelWorksheet(ExcelPackage excelPackage)
        {
            ExcelWorksheet ws;
            using (FileStream stream = File.OpenRead(Exceldefinition.FileName))
            {
                excelPackage.Load(stream);
            }
            try
            {
                ws = excelPackage.Workbook.Worksheets[Exceldefinition.SheetName];
            }
            catch (Exception ex)
            {
                Log.Error("Worksheet not found!", ex);
                return null;
            }

            return ws;
        }
        /// <summary>Adds raw values from <see name="excelWorksheet"/> to <see name="_rawDataTable"/></summary>
        public void AddValuesToRawTable()
        {
            for (Int32 rowNum = worksheetStartRow; rowNum <= worksheetEndRow; rowNum++)
            {
                DataRow row = _rawDataTable.NewRow();

                Int32 endColumn;

                if (_rawDataTable.Columns.Count == 1)
                    endColumn = 1;
                else
                {
                    endColumn = _rawDataTable.Columns.Count + 1;
                }

                ExcelRange wsRow = excelWorksheet.Cells[rowNum, worksheetStartColumn, rowNum, worksheetEndColumn];

                foreach (var cell in wsRow)
                {
                    int colIndex =  Math.Abs(wsRow.Start.Column - cell.Start.Column);

                    DataColumn dataColumn = _rawDataTable.Columns[colIndex];
                    row[dataColumn.ColumnName] = cell.Value;
                    if (cell.Text.Equals("NULL"))
                        row[colIndex] = DBNull.Value;
                    else
                        row[colIndex] = cell.Text;
                }
                _rawDataTable.Rows.Add(row);
            }
            _rawDataTable.Rows.Cast<DataRow>().ToList().FindAll(row => String.IsNullOrEmpty(String.Join("", row.ItemArray))).ForEach(Row =>
                { _rawDataTable.Rows.Remove(Row); });
        }
        /// <summary>Sets DataTypes for Columns of <see name="_finalDataTable"/></summary>
        private void AddRawColumns()
        {
            foreach (DataColumn col in _rawDataTable.Columns)
            {
                DataColumn finalColumn = new DataColumn
                {
                    DataType = typeof(string),
                    ColumnName = col.ColumnName
                };
                if (!_finalDataTable.Columns.Contains(finalColumn.ColumnName))
                    _finalDataTable.Columns.Add(finalColumn);
            }
        }
        /// <summary>Sets DataTypes for Columns of <see name="_finalDataTable"/></summary>
        private void AddTypedColumns()
        {
            foreach (DataColumn col in _rawDataTable.Columns)
            {
                object currentDataType = typeof(string);
                bool first = true;

                foreach (DataRow row in _rawDataTable.Rows)
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
                    Log.Debug(String.Format("Column {0} Row {1} DataType {2} identified Value {3}", col.ColumnName, row.Table.Rows.IndexOf(row), currentDataType, row[col.ColumnName].ToString()));
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

            CultureInfo cultureinfoDe = new CultureInfo("de-DE");
            CultureInfo cultureinfoUs = new CultureInfo("en-US");

            if (long.TryParse(stringToParse, out intValue))
                return intValue.GetType();

            if (stringToParse.Contains("-") || stringToParse.Contains("/"))
            {
                if (DateTime.TryParseExact(stringToParse, _patterns.ToArray(), cultureinfoUs,
                    DateTimeStyles.None, out datetimeValue))
                    return datetimeValue.GetType();
            }
            if (stringToParse.Contains("."))
            {
                if (DateTime.TryParseExact(stringToParse, _patterns.ToArray(), cultureinfoDe,
                    DateTimeStyles.None, out datetimeValue))
                    return datetimeValue.GetType();
            }
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

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    this._finalDataTable.Dispose();
                    this._rawDataTable.Dispose();
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }
}