using System;

namespace MontfoortIT.Office.Excel
{
    public class ColumnTemplate
    {
        public string Header { get; private set; }
        public string FieldName { get; private set; }
        public string Format { get; private set; }
        public NumberFormat NumberFormat { get; set; }

        public ColumnTemplate(string fieldName)
            :this(fieldName,fieldName)
        {
            NumberFormat = NumberFormat.Default;
        }

        public ColumnTemplate(string header, string fieldName)
        {
            if(string.IsNullOrEmpty(header))
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

        public virtual object GetValue(object o)
        {
            if (o == null)
                return null;

            if (!string.IsNullOrEmpty(FieldName))
            {
                var propType = o.GetType().GetProperty(FieldName);
                if(propType != null)
                    return propType.GetValue(o, null);
            }

            throw new NotImplementedException("Not implemented yet in Standard");
            //if (string.IsNullOrEmpty(Format))
            //{
            //    return DataBinder.Eval(o, FieldName);    
            //}
            //return DataBinder.Eval(o, FieldName, Format);
        }
    }
}
