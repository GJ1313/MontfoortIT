using System.Collections.Generic;
using MontfoortIT.Library.Templates;
using System;

namespace MontfoortIT.Library.Streams.FileConvertors
{
    internal class ObjectWriter<T>: IStructureWriter
    {
        private List<FuncColumnTemplate<T>> _templates;
        private Func<T> _constructor;
        private int _columnIndex;
        private T _currentObject;

        public ObjectWriter(List<FuncColumnTemplate<T>> templates, Func<T> constructor)
        {
            _templates = templates;
            _constructor = constructor;
        }

        public List<T> Objects { get; internal set; }

        public void WriteStartColumn()
        {
            
        }

        public void WriteStartDocument()
        {
            Objects = new List<T>();
        }

        public void WriteStartRow()
        {
            _currentObject = _constructor();
            _columnIndex = 0;
        }

        public void WriteStartTable()
        {
            
        }

        public void WriteString(string text)
        {
            if (_columnIndex >= _templates.Count)
                return;

            _templates[_columnIndex].SetValue(_currentObject, text);
        }


        public void WriteEndColumn()
        {
            _columnIndex++;
        }

        public void WriteEndDocument()
        {
            
        }

        public void WriteEndRow()
        {
            Objects.Add(_currentObject);
        }

        public void WriteEndTable()
        {
        }

    }
}