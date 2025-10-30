using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MontfoortIT.Office.Excel.Csv
{
    public class CsvSheet : ISheet, IAsyncDisposable
    {
        private Options _options;
        private StreamWriter _writer;
        private int _rowCount = 0;

        public int RowCount { get { return _rowCount; } }

        public List<ColumnTemplate> ColumnTemplate { get; set; }

        public async Task WriteFromObjectsAsync<D>(IAsyncEnumerable<D> objects, Stream fileStream, Options options, Encoding encoding, bool writeHeaderRow = true)
        {
            Initialize(fileStream, options, encoding);

            Line line;
            if (writeHeaderRow)
            {
                line = GenerateHeader();
                await WriteLineToCsvAsync(line);
            }

            await foreach (object o in objects)
            {
                line = CreateLineFromObject(o);
                await WriteLineToCsvAsync(line);
            }

            await WriteObjectsFinishedAsync();
        }

        private void Initialize(Stream fileStream, Options options, Encoding encoding)
        {
            _options = options;
            // Do not close stream, caller is responsible for closing filestream
            _writer = new StreamWriter(fileStream, encoding);

            if (ColumnTemplate == null)
                throw new ArgumentNullException("The property ColumnTemplate should be set before calling 'FillFromObjects'");
        }

        public async Task WriteFromObjectAsync<D>(D obj, bool writeHeader = false)
        {
            if (_writer == null)
                throw new NotSupportedException("First call Initialize");

            if (ColumnTemplate == null)
                throw new ArgumentNullException("The property ColumnTemplate should be set before calling 'FillFromObjects'");

            Line line;
            if (writeHeader)
            {
                line = GenerateHeader();
                await WriteLineToCsvAsync(line);
            }

            line = CreateLineFromObject(obj);
            await WriteLineToCsvAsync(line);
        }

        public async Task WriteFromObjectsAsync<D>(IEnumerable<D> objects, Stream fileStream, Options options, Encoding encoding, bool writeHeaderRow = true)
        {
            _options = options;
            // Do not close stream, caller is responsible for closing filestream
            _writer = new StreamWriter(fileStream, encoding);

            if (ColumnTemplate == null)
                throw new ArgumentNullException("The property ColumnTemplate should be set before calling 'FillFromObjects'");

            Line line;
            if (writeHeaderRow)
            {
                line = GenerateHeader();
                await WriteLineToCsvAsync(line);
            }

            foreach (object o in objects)
            {
                line = CreateLineFromObject(o);
                await WriteLineToCsvAsync(line);
            }

            await WriteObjectsFinishedAsync();
        }

        private Line CreateLineFromObject(object obj)
        {
            Line line = new Line();

            int column = 0;
            foreach (ColumnTemplate columnTemplate in ColumnTemplate)
            {
                object value = columnTemplate.GetValue(obj);

                Column cell = line[column];
                cell.NumberFormat = columnTemplate.NumberFormat;
                try
                {
                    if (value == null)
                    { } // ignore
                    else if (value is double doubleValue)
                        cell.Number = (decimal)doubleValue; // First cast is nescesarry for unboxing
                    else if (value is float floatValue)
                        cell.Number = (decimal)floatValue;
                    else if (value is decimal decimalValue)
                        cell.Number = decimalValue;
                    else
                    {
                        string text = value.ToString();
                        if (text != null)
                            cell.Text = text;
                    }
                }
                catch (InvalidCastException)
                { }


                column++;
            }

            return line;
        }


        private Line GenerateHeader()
        {
            Line line = new Line();
            int column = 0;
            foreach (ColumnTemplate columnTemplate in ColumnTemplate)
            {
                if (columnTemplate.Header != null)
                    line[column].Text = columnTemplate.Header;
                column++;
            }

            return line;
        }


        private async Task WriteLineToCsvAsync(Line line)
        {
            bool firstCell = true;
            for (int c = 0; c < line.ColumnCount; c++)
            {
                if (!firstCell)
                    await _writer.WriteAsync(_options.Seperator);
                else
                    firstCell = false;

                var cell = line[c];
                if (cell.Number.HasValue)
                {
                    await _writer.WriteAsync(cell.Number.Value.ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    if (!string.IsNullOrEmpty(cell.Text))
                    {
                        string text = cell.Text;
                        if (text.Contains('"') || text.Contains(_options.Seperator) || _options.QuotesAroundText)
                            text = $"\"{text.Replace("\"", "\"\"")}\"";

                        await _writer.WriteAsync(text);
                    }
                }
            }
            if (_options.AddSeperatorOnLineEnd)
                await _writer.WriteAsync(_options.Seperator);
            await _writer.WriteLineAsync();

            _rowCount++;
        }

        public async Task WriteObjectsFinishedAsync()
        {
            if (_writer != null)
            {
                await _writer.FlushAsync();
                _writer = null;
            }
        }

        public async ValueTask DisposeAsync()
        {
            await WriteObjectsFinishedAsync();
        }
    }
}
