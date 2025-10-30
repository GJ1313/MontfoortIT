using System;
using System.Collections.Generic;

namespace MontfoortIT.Library.Streams.FileConvertors
{
    internal class InMemoryStructureWriter : IStructureWriter
    {
        List<string> _columns = new List<string>();

        string currentText = string.Empty;
        
        public void WriteEndColumn()
        {
            _columns.Add(currentText);
            currentText = string.Empty;
            
        }

        public void WriteEndDocument()
        {
        }

        public void WriteEndRow()
        {
            _columns = new List<string>();
        }

        public void WriteEndTable()
        {
            throw new System.NotImplementedException();
        }

        public void WriteStartColumn()
        {
            
        }

        public void WriteStartDocument()
        {
            
        }

        public void WriteStartRow()
        {
            
        }

        public void WriteStartTable()
        {
            
        }

        public void WriteString(string text)
        {
            currentText += text;
        }

        internal string[] GetColumns()
        {
            return _columns.ToArray();
        }
    }
}