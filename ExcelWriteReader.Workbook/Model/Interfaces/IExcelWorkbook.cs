using ExcelWriteReader.Workbook.Constants;
using System.Collections.Generic;

namespace ExcelWriteReader.Workbook.Model.Interfaces
{
    public interface IExcelWorkbook
    {
        void SetActiveWorkSheet(string workSheetName);
        void ClearWorksheet();
        void SetWorkSheetCellValue(int rowValue, int columnValue, object value);
        void SaveAs(string filePath);
        void ReleaseFromMemory();
        /// <summary>
        /// Only needed for Interop
        /// </summary>
        void WriteArrayToSheet();
        //void SetArrayDimensionsForStats(IList<ISeriesStats> statsList);
        void SetArrayDimensions(int numRows, int numColumns);
        IEnumerable<string> GetNamedRangesOfActiveWorksheet();
        //List<KeyValuePair<string, object>> ReadNamedRange(string sheetName, string namedRange);
        IRangeData ReadNamedRangeOrTable(string sheetName, string namedRange);
        string[] GetStringVectorNamedRangeOrTable(string sheetName, string namedRange);
        string[,] GetStringArrayNamedRangeOrTable(string sheetName, string namedRange);
        ExcelPackage Package { get; }
    }
}