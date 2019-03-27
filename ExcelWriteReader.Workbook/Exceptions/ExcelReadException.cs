using System;

namespace ExcelWriteReader.Workbook.Exceptions
{
    public class ExcelReadException : Exception
    {
        public ExcelReadException(string message) : base(message)
        {

        }

        public ExcelReadException(string message, Exception innerException) : base(message, innerException)
        {

        }
    }
}
