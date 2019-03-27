using ExcelWriteReader.Workbook.Constants;
using ExcelWriteReader.Workbook.Model.Interfaces;

namespace ExcelWriteReader.Workbook.Factory.Interfaces
{
    public interface IExcelWorkbookFactory
    {
        IExcelWorkbook GetExcelWorkbook(ExcelPackage excelPackage, string filePath);

        ///// <summary>
        ///// Gives you a closed xml workbook
        ///// </summary>
        IExcelWorkbook GetExcelWorkbook();
    }
}