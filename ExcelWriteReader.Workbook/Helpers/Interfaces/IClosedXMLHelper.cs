using ClosedXML.Excel;
using ExcelWriteReader.Workbook.Constants;
using System.Collections.Generic;

namespace ExcelWriteReader.Workbook.Helpers.Interfaces
{
    internal interface IClosedXMLHelper
    {
        IDictionary<ExcelDataType, object> ReadExcelNamedRange(IXLWorkbook workbook, string namedRange);
        IDictionary<ExcelDataType, object> ReadTable(IXLWorkbook workbook, string tableName);
        IDictionary<ExcelDataType, object> ReadNamedRangeOrTable(IXLWorkbook workbook,
            string namedRangeOrTableName);
        IDictionary<ExcelDataType, object> ReadTable(IXLTable table);
    }
}