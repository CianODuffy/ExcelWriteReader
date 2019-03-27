using ExcelWriteReader.Workbook.Constants;
using System.Collections.Generic;

namespace ExcelWriteReader.Workbook.Extensions
{
    internal static class DataTypeExtensions
    {
        internal static T CastData<T>(this IDictionary<ExcelDataType, object> data, ExcelDataType dataType)
        {
            return (T)data[dataType];
        }
    }
}
