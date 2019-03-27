using System;

namespace ExcelWriteReader.Workbook.Exceptions
{
    public class ClosedXMLReadException : ExcelReadException
    {
        public ClosedXMLReadException(string message) : base(message)
        {

        }

        public ClosedXMLReadException(string message, Exception innerException) : base(message, innerException)
        {

        }
    }
}
