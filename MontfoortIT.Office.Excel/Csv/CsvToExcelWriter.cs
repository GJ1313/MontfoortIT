using MontfoortIT.Library.Streams.FileConvertors;
using System;
using System.Collections.Generic;
using System.Text;

namespace MontfoortIT.Office.Excel.Csv
{
    public class CsvToExcelWriter : IStructureWriter, IDisposable
    {
        private Application _application;

        int _currentColumn = 0;
        int _currentRow = 0;
        private Sheet _currentSheet;

        public CsvToExcelWriter()
        {
            _application = new Application();
            _currentSheet = _application.Workbook.Sheets.Create("start");
        }

        public CsvToExcelWriter(Application application, Sheet sheet, int startRow = 0)
        {
            _application = application;
            _currentSheet = sheet;
            _currentRow = startRow;
        }

        public void WriteEndColumn()
        {
            _currentColumn++;
        }

        public void WriteEndDocument()
        {            
        }

        public void WriteEndRow()
        {
            _currentRow++;
            _currentColumn = 0;
        }

        public void WriteEndTable()
        {
            
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
            _currentSheet.Cells[_currentRow, _currentColumn].Text = text;
        }

        public void Dispose()
        {
            if (_application != null)
            {
                _application.Dispose();
                _application = null;
            }
        }
    }
}
