using ClosedXML.Excel;
using ExcelWriteReader.Workbook.Constants;
using ExcelWriteReader.Workbook.Helpers.Interfaces;
using ExcelWriteReader.Workbook.Model.Interfaces;
using ExcelWriteReader.Workbook.StaticFunctions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Interop = Microsoft.Office.Interop.Excel;

namespace ExcelWriteReader.Workbook.Model
{
    /// <summary>
    /// This class writes data to CloseXML and Interop Interop sheets
    /// Removing the need for duplicating methods for each package.
    /// </summary>
    public class ExcelWorkbook : IExcelWorkbook
    {
        private XLWorkbook _closedXmlWorkbook;
        private Interop.Workbook _interopWorkbook;
        private Interop.Application _excelApplication;
        private readonly string _excelApplicationName = "Excel.Application";
        private IXLWorksheet _closedXMLWorksheet;
        private Interop.Worksheet _interopWorksheet;
        private object[,] _data;
        private readonly int _startRow = 1, _startColumn = 1;
        private readonly IArrayHelper _arrayHelper;
        private readonly IClosedXMLHelper _closedXMLHelper;

        internal ExcelWorkbook(ExcelPackage excelPackage, string filePath, IArrayHelper arrayHelper,
            IClosedXMLHelper closedXmlHelper)
        {
            Package = excelPackage;
            _arrayHelper = arrayHelper;
            _closedXMLHelper = closedXmlHelper;

            switch (excelPackage)
            {
                case ExcelPackage.ClosedXml:
                    _closedXmlWorkbook = new XLWorkbook(filePath);
                    break;
                default:
                    _excelApplication = (Interop.Application)Marshal.GetActiveObject(_excelApplicationName);
                    string workbookName = GetWorkbookName(filePath);
                    _interopWorkbook = _excelApplication.Workbooks[workbookName];
                    break;
            }
        }

        ///// <summary>
        ///// Returns a closed xml workbook
        ///// </summary>
        internal ExcelWorkbook(IArrayHelper arrayHelper, IClosedXMLHelper closedXmlHelper)
        {
            _arrayHelper = arrayHelper;
            _closedXMLHelper = closedXmlHelper;
            Package = ExcelPackage.ClosedXml;
            _closedXmlWorkbook = new XLWorkbook();
        }

        public ExcelPackage Package { get; }

        private string GetWorkbookName(string filePath)
        {
            return Path.GetFileName(filePath);
        }

        public void SetActiveWorkSheet(string workSheetName)
        {
            switch (Package)
            {
                case ExcelPackage.ClosedXml:
                    {
                        if (_closedXmlWorkbook.Worksheets.All(x => x.Name != workSheetName))
                            _closedXMLWorksheet = _closedXmlWorkbook.AddWorksheet(workSheetName);
                        else
                            _closedXMLWorksheet = _closedXmlWorkbook.Worksheets.First(x => x.Name == workSheetName);
                        break;
                    }
                default:
                    {
                        foreach (Interop.Worksheet sheet in _interopWorkbook.Sheets)
                        {
                            if (sheet.Name == workSheetName)
                                _interopWorksheet = sheet;
                        }

                        if (_interopWorkbook == null)
                        {
                            _interopWorksheet = _interopWorkbook.Worksheets.Add(); //Changed to add tab if doesn't exist.
                            _interopWorksheet.Name = workSheetName;
                        }
                        break;
                    }
            }
        }


        public string[] GetStringVectorNamedRangeOrTable(string sheetName, string namedRange)
        {
            var data = ReadNamedRangeOrTable(sheetName, namedRange);
            string[,] stringData = data.GetTextArray();
            string[] output = _arrayHelper.ConvertArrayToVector(stringData);
            return output;
        }

        public string[,] GetStringArrayNamedRangeOrTable(string sheetName, string namedRange)
        {
            var data = ReadNamedRangeOrTable(sheetName, namedRange);
            return data.GetTextArray();
        }

        /// <summary>
        /// Interop views tables as named ranges, whereas interop views them separately
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="namedRangeOrTableName"></param>
        /// <returns></returns>
        public IRangeData ReadNamedRangeOrTable(string sheetName, string namedRangeOrTableName)
        {
            switch (Package)
            {
                case ExcelPackage.ClosedXml:
                    IDictionary<ExcelDataType, object> dic = _closedXMLHelper
                        .ReadNamedRangeOrTable(_closedXmlWorkbook, namedRangeOrTableName);
                    return new RangeData(dic);
                default:
                    return ReadInteropExcel
                        .ReadExcelSheetNamedRange(_interopWorkbook, sheetName, namedRangeOrTableName);
            }
        }


        public IEnumerable<string> GetNamedRangesOfActiveWorksheet()
        {
            switch (Package)
            {
                case ExcelPackage.ClosedXml:
                    return _closedXMLWorksheet.NamedRanges.Select(x => x.Name);
                default:
                    {
                        var output = new List<string>();
                        foreach (Interop.Name name in _interopWorkbook.Names)
                            output.Add(name.Name);
                        return output;
                    }
            }
        }

        public void SetArrayDimensions(int numRows, int numColumns)
        {
            _data = new object[numRows, numColumns];
        }

        public void ClearWorksheet()
        {
            switch (Package)
            {
                case ExcelPackage.ClosedXml:
                    _closedXMLWorksheet.Clear();
                    break;
                default:
                    _interopWorksheet.Cells.Clear();
                    break;
            }
        }

        public void SetWorkSheetCellValue(int rowIndex, int columnIndex, object value)
        {
            switch (Package)
            {
                case ExcelPackage.ClosedXml:
                    _closedXMLWorksheet.Cell(rowIndex, columnIndex).Value = value;
                    break;
                default:
                    {
                        _data[rowIndex - 1, columnIndex - 1] = value;
                        break;
                    }
            }
        }

        /// <summary>
        /// Only needed for Interop
        /// </summary>
        public void WriteArrayToSheet()
        {
            if (Package.Equals(ExcelPackage.Interop))
            {
                var noRows = _data.GetLength(0);
                var noCols = _data.GetLength(1);
                Interop.Range excelRange = _interopWorksheet
                    .Range[_interopWorksheet.Cells[_startRow, _startColumn],
                        _interopWorksheet.Cells[noRows, noCols]];
                excelRange.Value2 = _data;
            }
        }

        public void SaveAs(string filePath)
        {
            switch (Package)
            {
                case ExcelPackage.ClosedXml:
                    _closedXmlWorkbook.SaveAs(filePath);
                    break;
                default:
                    //Only use interop if open
                    _interopWorkbook.Save();
                    break;
            }
        }

        public void ReleaseFromMemory()
        {
            switch (Package)
            {
                case ExcelPackage.ClosedXml:
                    _closedXmlWorkbook.Dispose();
                    break;
                default:
                    {
                        if (_interopWorksheet != null) Marshal.ReleaseComObject(_interopWorksheet);
                        if (_interopWorkbook != null) Marshal.ReleaseComObject(_interopWorkbook);
                        if (_excelApplication != null) Marshal.ReleaseComObject(_excelApplication);
                        _data = null;
                        break;
                    }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
