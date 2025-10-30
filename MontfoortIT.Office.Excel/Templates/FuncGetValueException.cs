using System;
using System.Runtime.Serialization;

namespace MontfoortIT.Office.Excel.Templates
{
    [Serializable]
    internal class FuncGetValueException : Exception
    {
        private Exception _e;

        public FuncGetValueException()
        {
        }

        public FuncGetValueException(Exception e)
        {
            _e = e;
        }

        public FuncGetValueException(string message) : base(message)
        {
        }

        public FuncGetValueException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected FuncGetValueException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}