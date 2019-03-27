using ExcelWriteReader.Workbook.Constants;
using ExcelWriteReader.Workbook.Factory;
using ExcelWriteReader.Workbook.Factory.Interfaces;
using ExcelWriteReader.Workbook.Model.Interfaces;
using System;
using System.IO;

namespace ExcelWriteReader.Console
{
    /// <summary>
    /// This script demonstrates reading data from an excel spreadsheet
    /// using the ExcelWorkbook class
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            string tableName = "Table", namedRangeName = "NamedRange", sheetName = "Data";      
            string testExcelSheetName = "Data.xlsx";
            string consoleProjectLocalPath = Directory.GetParent(Directory
                .GetParent(Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).FullName).FullName).FullName;

            //ClosedXML requires .xlsx and .xlsm files. Older version of excel are no compatible.
            string pathXls = consoleProjectLocalPath + "\\" + testExcelSheetName;
            IExcelWorkbookFactory factory = new ExcelWorkbookFactory();
            IExcelWorkbook workbook;

            //First assume excel worksheet is closed and try to read with ClosedXML. IExcelWorkbook is opened in 
            //ClosedXML mode.
            try
            {
                workbook = factory.GetExcelWorkbook(ExcelPackage.ClosedXml, pathXls);
            }
            //If excel workbook is open closed XML will throw an exception. The IExcelWorkbook will then be 
            //opened in Interop mode
            catch(IOException e)
            {
                workbook = factory.GetExcelWorkbook(ExcelPackage.Interop, pathXls);
            }

            IRangeData importedDataTable = workbook.ReadNamedRangeOrTable(sheetName, tableName);
            string[,] tableTextData = importedDataTable.GetTextArray();
            double?[,] tableNumericData = importedDataTable.GetNumericArray();

            IRangeData importedDataNamedRange = workbook.ReadNamedRangeOrTable(sheetName, namedRangeName);
            string[,] namedRangeTextData = importedDataNamedRange.GetTextArray();
            double?[,] namedRangeNumericData = importedDataNamedRange.GetNumericArray();

            //table includes headers in ClosedXML model but not in Interop
            //so height is one less than the same named range in Interop
            bool lengthTest = tableTextData.GetLength(0) == (namedRangeTextData.GetLength(0) - 1);
            bool isInterop = lengthTest;
        }
    }
}
