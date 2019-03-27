using ExcelWriteReader.Workbook.Constants;
using ExcelWriteReader.Workbook.Model;
using ExcelWriteReader.Workbook.Model.Interfaces;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Interop = Microsoft.Office.Interop.Excel;

namespace ExcelWriteReader.Workbook.StaticFunctions
{
    internal static class ReadInteropExcel
    {
        private static IRangeData Xlsread(Interop.Range excelRange)
        {
            try
            {
                Array readValues;
                var value = excelRange.Cells.Value2;
                if (value is string)
                {
                    readValues = Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 });
                    readValues.SetValue(value, 1, 1);
                }
                else if (excelRange.Cells.Count.Equals(1))
                {
                    var dic = new Dictionary<ExcelDataType, object>
                    {
                        {ExcelDataType.Numeric, excelRange.Cells.Value2}
                    };
                    return new RangeData(dic);
                }
                else
                {
                    readValues = (Array)excelRange.Cells.Value2;
                }

                string[,] stringArray = null;
                double?[,] numericArray = null;
                Array readFormulae;

                if (readValues != null)
                {
                    stringArray = GetStringArray(readValues);
                    numericArray = GetNumericArray(readValues);
                }

                var formula = excelRange.Cells.Formula as string;
                if (formula != null && formula == "")
                    readFormulae = null;
                else
                {
                    if (excelRange.Cells.Formula is string)
                    {
                        readFormulae = Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 });
                        readFormulae.SetValue(excelRange.Cells.Formula, 1, 1);
                    }
                    else
                    {
                        readFormulae = (Array)excelRange.Cells.Formula;
                    }

                }

                var output = new Dictionary<ExcelDataType, object>();

                output.Add(ExcelDataType.Text, stringArray);
                output.Add(ExcelDataType.Numeric, numericArray);
                output.Add(ExcelDataType.Formulae, readFormulae);
                return new RangeData(output);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        internal static IRangeData ReadExcelSheetUsedRange(Interop.Workbook workbook, string sheetName)
        {
            Interop.Worksheet excelWorksheet = null;
            Interop.Range excelRange = null;
            try
            {
                excelWorksheet = (Interop.Worksheet)workbook.Worksheets.get_Item(sheetName);
                excelRange = excelWorksheet.UsedRange;
                return Xlsread(excelRange);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //release objects
                if (excelWorksheet != null) Marshal.ReleaseComObject(excelWorksheet);
                if (excelRange != null) Marshal.ReleaseComObject(excelRange);
            }
        }
        internal static IRangeData ReadExcelSheetNamedRange(Interop.Workbook workbook, string sheetName, string namedRange)
        {
            Interop.Worksheet excelWorksheet = null;
            Interop.Range excelRange = null;
            try
            {
                excelWorksheet = (Interop.Worksheet)workbook.Worksheets.get_Item(sheetName);
                excelRange = excelWorksheet.Range[namedRange];
                return Xlsread(excelRange);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //release objects
                if (excelWorksheet != null) Marshal.ReleaseComObject(excelWorksheet);
                if (excelRange != null) Marshal.ReleaseComObject(excelRange);
            }
        }
      
        private static string[,] GetStringArray(Array values)
        {
            double number;
            // create a new string array
            string[,] stringArray = new string[values.GetLength(0), values.GetLength(1)];
            // loop through the 2-D System.Array and populate the output array
            for (int iRow = 1; iRow <= values.GetLength(0); iRow++)
            {
                for (int iCol = 1; iCol <= values.GetLength(1); iCol++)
                {
                    if (values.GetValue(iRow, iCol) == null || double.TryParse(values.GetValue(iRow, iCol).ToString(), out number))
                        stringArray[iRow - 1, iCol - 1] = "";
                    else
                        stringArray[iRow - 1, iCol - 1] = values.GetValue(iRow, iCol).ToString();
                }
            }
            return stringArray;
        }

        private static double?[,] GetNumericArray(Array values)
        {
            double number;
            // create a new double array
            double?[,] doubleArray = new double?[values.GetLength(0), values.GetLength(1)];
            // loop through the 2-D System.Array and populate the output array
            for (int iRow = 1; iRow <= values.GetLength(0); iRow++)
            {
                for (int iCol = 1; iCol <= values.GetLength(1); iCol++)
                {
                    if (values.GetValue(iRow, iCol) == null || !double.TryParse(values.GetValue(iRow, iCol).ToString(), out number))
                        doubleArray[iRow - 1, iCol - 1] = null;
                    else
                        doubleArray[iRow - 1, iCol - 1] = number;
                }
            }
            return doubleArray;
        }
    }
}