using ExcelWriteReader.Workbook.Constants;
using ExcelWriteReader.Workbook.Factory.Interfaces;
using ExcelWriteReader.Workbook.Helpers;
using ExcelWriteReader.Workbook.Helpers.Interfaces;
using ExcelWriteReader.Workbook.Model;
using ExcelWriteReader.Workbook.Model.Interfaces;

namespace ExcelWriteReader.Workbook.Factory
{
    public class ExcelWorkbookFactory : IExcelWorkbookFactory
    {
        private readonly IArrayHelper _arrayHelper;
        private readonly IClosedXMLHelper _closedXmlHelper;

        public ExcelWorkbookFactory()
        {
            _arrayHelper = new ArrayHelper();
            _closedXmlHelper = new ClosedXMLHelper();
        }

        public IExcelWorkbook GetExcelWorkbook(ExcelPackage excelPackage, string filePath)
        {
            return new ExcelWorkbook(excelPackage, filePath, _arrayHelper, _closedXmlHelper);
        }
        ///// <summary>
        ///// Returns a closed xml workbook
        ///// </summary>
        public IExcelWorkbook GetExcelWorkbook()
        {
            return new ExcelWorkbook(_arrayHelper, _closedXmlHelper);
        }
    }
}
