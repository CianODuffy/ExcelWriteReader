using ClosedXML.Excel;
using ExcelWriteReader.Workbook.Constants;
using ExcelWriteReader.Workbook.Exceptions;
using ExcelWriteReader.Workbook.Helpers.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelWriteReader.Workbook.Helpers
{
    internal class ClosedXMLHelper : IClosedXMLHelper
    {
        public IDictionary<ExcelDataType, object> ReadExcelNamedRange(IXLWorkbook workbook, string namedRange)
        {
            IXLNamedRange range = workbook.NamedRange(namedRange);
            if (range == null)
                throw new ClosedXMLReadException($"There is no named range called {namedRange} in this workbook");

            int count = range.Ranges.First().Cells().Count();

            //if solo then don't return collection.
            if (count.Equals(1))
            {
                IXLCell cell = range.Ranges.First().Cells().First();
                switch (cell.DataType)
                {
                    case XLDataType.DateTime:
                        return new Dictionary<ExcelDataType, object>
                        {
                            {ExcelDataType.Numeric, cell.GetDateTime().ToOADate()}
                        };
                    case XLDataType.Number:
                        return new Dictionary<ExcelDataType, object>
                        {
                            {ExcelDataType.Numeric, cell.GetDouble()}
                        };
                    case XLDataType.Text:
                        string textValue = GetSingleCellTextValue(cell);
                        KeyValuePair<ExcelDataType, object> parsed = ParseString(textValue);

                        switch (parsed.Key)
                        {
                            case ExcelDataType.Numeric:
                                return new Dictionary<ExcelDataType, object>
                                {
                                    {ExcelDataType.Numeric, (double) parsed.Value}
                                };
                            case ExcelDataType.Text:
                                string[,] array = new string[1, 1];
                                array[0, 0] = textValue;
                                return new Dictionary<ExcelDataType, object>
                                {
                                    {ExcelDataType.Text, array}
                                };
                            default:
                                throw new NotImplementedException("I haven't implemented formulas yet");
                        }
                }
            }
            IXLTable table = range.Ranges.First().AsTable();
            return ReadTable(table);
        }

        public IDictionary<ExcelDataType, object> ReadTable(IXLWorkbook workbook, string tableName)
        {
            IXLTable table = workbook.Table(tableName);
            return ReadTable(table);
        }

        public IDictionary<ExcelDataType, object> ReadNamedRangeOrTable(IXLWorkbook workbook,
            string namedRangeOrTableName)
        {
            try
            {
                return ReadExcelNamedRange(workbook, namedRangeOrTableName);
            }
            catch (ClosedXMLReadException e)
            {
                return ReadTable(workbook, namedRangeOrTableName);
            }
        }

        //This bit is needed because excel often thinks numbers are text
        //ClosedXML seems to have trouble with some references to include this to get around exceptions
        private string GetSingleCellTextValue(IXLCell cell)
        {
            return (string)cell.CachedValue;
        }

        private KeyValuePair<ExcelDataType, object> ParseString(string textValue)
        {
            bool pass = double.TryParse(textValue, out var doubleValue);
            if (pass)
                return new KeyValuePair<ExcelDataType, object>(ExcelDataType.Numeric, doubleValue);
            return new KeyValuePair<ExcelDataType, object>(ExcelDataType.Text, textValue);
        }

        public IDictionary<ExcelDataType, object> ReadTable(IXLTable table)
        {
            int rows = table.RowCount();
            int columns = table.ColumnCount();
            var numericData = new double?[rows, columns];
            var textData = new string[rows, columns];
            var formula = new string[rows, columns];

            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                    IXLCell cell = table.Cell(i, j);
                    switch (cell.DataType)
                    {
                        case XLDataType.Number:
                            numericData[i - 1, j - 1] = cell.GetDouble();
                            break;
                        case XLDataType.Text:
                            {
                                string textValue = GetSingleCellTextValue(cell);
                                KeyValuePair<ExcelDataType, object> parsed = ParseString(textValue);

                                switch (parsed.Key)
                                {
                                    case ExcelDataType.Numeric:
                                        numericData[i - 1, j - 1] = (double)parsed.Value;
                                        break;
                                    case ExcelDataType.Text:
                                        textData[i - 1, j - 1] = textValue;
                                        break;
                                    default:
                                        throw new NotImplementedException("I haven't implemented formulas yet");
                                }
                                break;
                            }
                    }

                    if (cell.HasFormula)
                        formula[i - 1, j - 1] = cell.FormulaA1;
                }
            }
            return new Dictionary<ExcelDataType, object>
            {
                {ExcelDataType.Numeric, numericData},
                {ExcelDataType.Formulae, formula},
                {ExcelDataType.Text, textData}
            };
        }
    }
}