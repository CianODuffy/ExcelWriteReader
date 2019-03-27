using System;

namespace ExcelWriteReader.Workbook.Exceptions
{
    public class InteropReadException : ExcelReadException
    {
        public InteropReadException(string message) : base(message)
        {

        }

        public InteropReadException(string message, Exception innerException) : base(message, innerException)
        {

        }
    }
}

