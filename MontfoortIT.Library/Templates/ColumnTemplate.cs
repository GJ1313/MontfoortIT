using System;
using System.Collections.Generic;
using System.Text;

namespace MontfoortIT.Library.Templates
{
    public abstract class ColumnTemplate
    {
        public string Header { get; private set; }
        public string FieldName { get; private set; }
        public string Format { get; private set; }

        public ColumnTemplate(string fieldName)
            : this(fieldName, fieldName)
        {
        }

        public ColumnTemplate(string header, string fieldName)
        {
            if (string.IsNullOrEmpty(header))
                throw new ArgumentException("header");

            if (string.IsNullOrEmpty(fieldName))
                throw new ArgumentException("fieldName");

            Header = header;
            FieldName = fieldName;
        }

        public ColumnTemplate(string header, string fieldName, string format)
        {
            if (string.IsNullOrEmpty(header))
                throw new ArgumentException("header");

            if (string.IsNullOrEmpty(fieldName))
                throw new ArgumentException("fieldName");

            Header = header;
            FieldName = fieldName;
            Format = format;
        }

        public abstract object GetValue(object o);
    }
}
