using ExcelWriteReader.Workbook.Helpers.Interfaces;

namespace ExcelWriteReader.Workbook.Helpers
{
    internal class ArrayHelper : IArrayHelper
    {
        public T[] ConvertArrayToVector<T>(T[,] twoDimensionalArray)
        {
            int length = twoDimensionalArray.GetLength(0);
            T[] output = new T[length];
            for (int i = 0; i < length; i++)
            {
                output[i] = twoDimensionalArray[i, 0];
            }
            return output;
        }
        /// <summary>
        /// Inclusive of cells
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="twoDimensionalArray"></param>
        /// <param name="subMatrixStartRow"></param>
        /// <param name="subMatrixEndRow"></param>
        /// <param name="subMatrixColumn"></param>
        /// <returns></returns>
        public T[] ConvertArrayToVector<T>(T[,] twoDimensionalArray, int subMatrixStartRow, int subMatrixEndRow,
            int subMatrixColumn)
        {
            T[] output = new T[subMatrixEndRow - subMatrixStartRow + 1];
            for (int i = subMatrixStartRow; i <= subMatrixEndRow; i++)
            {
                output[i - subMatrixStartRow] = twoDimensionalArray[i, subMatrixColumn];
            }
            return output;
        }

        public int GetFirstNullRowIndexInColumn(double?[,] array, int startIndex, int columnIndex)
        {
            int length = array.GetLength(0);
            for (int i = startIndex; i < length; i++)
                if (!array[i, columnIndex].HasValue)
                    return i;
            return length;
        }


        public double[] ConvertToValues(double?[] vector)
        {
            int rows = vector.Length;
            var output = new double[rows];
            for (int i = 0; i < rows; i++)
            {
                if (vector[i].HasValue)
                    output[i] = vector[i].Value;
            }
            return output;
        }

        public double[,] ConvertToValues(double?[,] nullableArray)
        {
            int rows = nullableArray.GetLength(0);
            int columns = nullableArray.GetLength(1);
            var output = new double[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    if (nullableArray[i, j].HasValue)
                        output[i, j] = nullableArray[i, j].Value;
                }

            }
            return output;
        }
        /// <summary>
        /// Swaps columns and rows
        /// </summary>
        /// <param name="array"></param>
        /// <returns></returns>
        public double[][] ConvertToJaggedArray(double[,] array)
        {
            int rows = array.GetLength(0);
            int columns = array.GetLength(1);
            //[Time][Vector]
            var output = new double[columns][];
            for (int i = 0; i < columns; i++)
            {
                output[i] = new double[rows];
                for (int j = 0; j < rows; j++)
                {
                    output[i][j] = array[j, i];
                }
            }

            return output;
        }

        public int[,] CastNulableToIntArray(double?[,] array)
        {
            int rows = array.GetLength(0);
            int columns = array.GetLength(1);
            var output = new int[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    if (array[i, j].HasValue)
                        output[i, j] = (int)array[i, j].Value;
                }
            }
            return output;
        }
    }
}
