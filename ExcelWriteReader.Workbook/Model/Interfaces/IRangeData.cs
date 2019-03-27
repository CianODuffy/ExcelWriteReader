using System;

namespace ExcelWriteReader.Workbook.Model.Interfaces
{
    public interface IRangeData
    {
        string[,] GetTextArray();
        double?[,] GetNumericArray();
        /// <summary>
        /// Range is single cell
        /// </summary>
        /// <returns></returns>
        double GetNumber();
        /// <summary>
        /// Range is single cell
        /// </summary>
        /// <returns></returns>
        string GetString();
        /// <summary>
        /// Range is single cell
        /// </summary>
        /// <returns></returns>
        DateTime GetSingleDatetime();
    }
}