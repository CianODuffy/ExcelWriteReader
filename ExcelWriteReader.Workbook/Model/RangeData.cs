using ExcelWriteReader.Workbook.Constants;
using ExcelWriteReader.Workbook.Exceptions;
using ExcelWriteReader.Workbook.Extensions;
using ExcelWriteReader.Workbook.Model.Interfaces;
using ExcelWriteReader.Workbook.StaticFunctions;
using System;
using System.Collections.Generic;

namespace ExcelWriteReader.Workbook.Model
{
    /// <summary>
    /// You need to match up the data in the imported sheet to know which method to call
    /// </summary>
    public class RangeData : IRangeData
    {
        private readonly IDictionary<ExcelDataType, object> _data;
        private readonly string _prefix = "This range does not contain data of type";

        internal RangeData(IDictionary<ExcelDataType, object> data)
        {
            _data = data;
        }

        public DateTime GetSingleDatetime()
        {
            try
            {
                int serialDate = (int)GetNumber();
                return DateFunctions.FromExcelSerialDate(serialDate);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new ExcelReadException(_prefix + " DateTime", e);
            }
        }

        public string[,] GetTextArray()
        {
            try
            {
                return _data.CastData<string[,]>(ExcelDataType.Text);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new ExcelReadException(_prefix + " string[,]", e);
            }
        }

        public double?[,] GetNumericArray()
        {
            try
            {
                return _data.CastData<double?[,]>(ExcelDataType.Numeric);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new ExcelReadException(_prefix + " double?[,]", e);
            }
        }

        public double GetNumber()
        {
            try
            {
                return (double)_data[ExcelDataType.Numeric];
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new ExcelReadException(_prefix + " double", e);
            }
        }

        public string GetString()
        {
            try
            {
                string[,] textArray = GetTextArray();
                return textArray[0, 0];
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new ExcelReadException(_prefix + " string", e);
            }
        }
    }
}
