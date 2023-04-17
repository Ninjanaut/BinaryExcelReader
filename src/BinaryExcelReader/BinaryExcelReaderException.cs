using System;

namespace Ninjanaut.IO
{
    public class BinaryExcelReaderException : Exception
    {
        public BinaryExcelReaderException()
        {
        }

        public BinaryExcelReaderException(string message) : base(message)
        {
        }

        public BinaryExcelReaderException(string message, Exception inner) : base(message, inner)
        {
        }
    }
}
