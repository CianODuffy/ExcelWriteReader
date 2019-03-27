namespace ExcelWriteReader.Workbook.Helpers.Interfaces
{
    internal interface IArrayHelper
    {
        T[] ConvertArrayToVector<T>(T[,] twoDimensionalArray);
        int[,] CastNulableToIntArray(double?[,] array);
        /// <summary>
        /// Inclusive of cells
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="twoDimensionalArray"></param>
        /// <param name="subMatrixStartRow"></param>
        /// <param name="subMatrixEndRow"></param>
        /// <param name="subMatrixColumn"></param>
        /// <returns></returns>
        T[] ConvertArrayToVector<T>(T[,] twoDimensionalArray, int subMatrixStartRow, int subMatrixEndRow,
            int subMatrixColumn);
        double[] ConvertToValues(double?[] vector);
        double[,] ConvertToValues(double?[,] nullableArray);
        /// <summary>
        /// Swaps columns and rows
        /// </summary>
        /// <param name="array"></param>
        /// <returns></returns>
        double[][] ConvertToJaggedArray(double[,] array);

        int GetFirstNullRowIndexInColumn(double?[,] array, int startIndex, int columnIndex);
    }
}