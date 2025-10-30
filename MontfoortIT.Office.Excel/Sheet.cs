using MontfoortIT.Office.Excel.Csv;
using MontfoortIT.Office.Excel.Templates;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace MontfoortIT.Office.Excel
{
    public class Sheet:IDisposable, ISheet
    {
        private readonly Application _application;

        public string Name { get; set; }
        public int SheetId { get; set; }

        private CellCollection _cells;
        private PackagePart _originalPackageSheet;
        private Stream _stream;        
        private Options _options;
        private StreamWriter _writer;

        public CellCollection Cells
        {
            get
            {
                if (_cells == null)
                {
                    _cells = new CellCollection(_application.SharedStrings);
                }
                return _cells;
            }
        }


        public string Id
        {
            get
            {
                return "rIdNew" + SheetId;
            }
        }
        
        public List<ColumnTemplate> ColumnTemplate { get; set; }


        internal Sheet(Application application, string name, int sheetID)
        {
            _application = application;
            Name = name;
            SheetId = sheetID;            

            _options = new Options();
            _options.Seperator = ',';
        }

        public XDocument CreateDoc()
        {
            XDocument worksheetDoc;
            XElement sheetData;
            CreateBaseDoc(out worksheetDoc, out sheetData);

            Cells.AddToSheetData(sheetData, _application.ConvertNumberFormatToSAttribute);

            return worksheetDoc;
        }

        private static void CreateBaseDoc(out XDocument worksheetDoc, out XElement sheetData)
        {
            //Create a new XML document for the worksheet.
            worksheetDoc = new XDocument();

            //Obtain a reference to the root node, and then add
            //the XML declaration.

            worksheetDoc.Declaration = new XDeclaration("1.0", "UTF-8", "yes");

            //Create and append the worksheet node
            //to the document.
            XElement workSheet = new XElement(Namespaces.OpenFormat + "worksheet");
            worksheetDoc.Add(workSheet);

            //Create and add the sheetData node.
            sheetData = new XElement(Namespaces.OpenFormat + "sheetData");
            workSheet.Add(sheetData);
        }

        internal async Task<XDocument> CreateDocAndFillAsync<D>(IAsyncEnumerable<D> objects, int startRow = 0, bool writeHeaderRow = true)
        {
            int rowNumber;
            var row = StartFillHeader(startRow, writeHeaderRow, out rowNumber);

            CreateBaseDoc(out XDocument worksheetDoc, out XElement sheetData);
            Cells.AddToSheetData(sheetData, _application.ConvertNumberFormatToSAttribute);

            Dictionary<int, XElement> rowsDictionary = BuildUpRowsDictionary(sheetData);

            if (row != null)
            {
                Cells.WriteRowToExcell(sheetData, row, rowNumber, _application.ConvertNumberFormatToSAttribute, rowsDictionary, true);
                rowNumber++;
            }

            await Cells.FillAndAddToSheetDataAsync(this, sheetData, _application.ConvertNumberFormatToSAttribute, objects, rowNumber, rowsDictionary);

            return worksheetDoc;
        }

        internal async Task<XDocument> FillObjectsInTemplateAsync<D>(IAsyncEnumerable<D> objects, int startRow = 0, bool writeHeaderRow = true)
        {
            int rowNumber;
            var row = StartFillHeader(startRow, writeHeaderRow, out rowNumber);

            XDocument xSheet;
            using (Stream sharedStream = _originalPackageSheet.GetStream())
            using (XmlTextReader xmlReader = new XmlTextReader(sharedStream))
            {
                xSheet = XDocument.Load(xmlReader);
            }

            XElement sheetData = xSheet.Descendants(Namespaces.OpenFormat + "sheetData").First();
            Dictionary<int, XElement> rowsDictionary = BuildUpRowsDictionary(sheetData);

            for (int i = 0; i < startRow; i++) // fill first rows with current data
            {
                Row rowMain = Cells.GetRow(i);

                // write row...
                Cells.WriteRowToExcell(sheetData, rowMain, i, _application.ConvertNumberFormatToSAttribute, rowsDictionary, true);
            }

            if (row != null)
            {
                Cells.WriteRowToExcell(sheetData, row, rowNumber, _application.ConvertNumberFormatToSAttribute, rowsDictionary, true);
                rowNumber++;
            }

            await Cells.FillAndAddToSheetDataAsync(this, sheetData, _application.ConvertNumberFormatToSAttribute, objects, rowNumber, rowsDictionary);

            return xSheet;
        }

        private static Dictionary<int, XElement> BuildUpRowsDictionary(XElement sheetData)
        {
            Dictionary<int, XElement> rowsDictionary = new Dictionary<int, XElement>(); // for faster lookup of a row
            var rowElements = sheetData.Elements(Namespaces.OpenFormat + "row");
            foreach (var rowElement in rowElements)
            {
                int rowIndex;
                if (int.TryParse(rowElement.Attribute("r").Value, out rowIndex))
                    rowsDictionary.Add(rowIndex, rowElement);
            }

            return rowsDictionary;
        }

        internal void Read(PackagePart packageSheet, int emptyRowsToIgnore = -1)
        {
            SheetReader reader = new SheetReader(this, _application.SharedStrings);
            reader.Read(packageSheet, emptyRowsToIgnore);

            _originalPackageSheet = packageSheet;
        }

        public void FillFromObjects<D>(IEnumerable<D> objects, int startRow = 0, bool writeHeaderRow = true)
        {
            if (ColumnTemplate == null)
                throw new ArgumentNullException("The property ColumnTemplate should be set before calling 'FillFromObjects'");

            int rowNumber = startRow;

            Row row;
            if (writeHeaderRow)
            {
                row = GenerateHeader(rowNumber);
                Cells.AddRow(row);
                rowNumber++;
            }            

            foreach (object o in objects)
            {
                row = CreateRowFromObject(rowNumber, o);
                Cells.AddRow(row);
                rowNumber++;
            }
        }

        public async Task FillFromObjectsAsync<D>(IAsyncEnumerable<D> objects, int startRow = 0, bool writeHeaderRow = true)
        {
            int rowNumber;            
            var headerRow = StartFillHeader(startRow, writeHeaderRow, out rowNumber);
            if (headerRow != null)
            {
                Cells.AddRow(headerRow);
                rowNumber++;
            }

            await foreach (object o in objects)
            {
                Row row = CreateRowFromObject(rowNumber, o);
                Cells.AddRow(row);
                rowNumber++;
            }
        }



        private Row? StartFillHeader(int startRow, bool writeHeaderRow, out int rowNumber)
        {
            if (ColumnTemplate == null)
                throw new ArgumentNullException("The property ColumnTemplate should be set before calling 'FillFromObjects'");

            rowNumber = startRow;
            if (writeHeaderRow)
                return GenerateHeader(rowNumber);                                                
            return null;
        }

        public async Task WriteFromObjectsAsync<D>(IAsyncEnumerable<D> objects, Stream fileStream, Options options, Encoding encoding, bool writeHeaderRow = true)
        {
            await FillFromObjectsAsync(objects, writeHeaderRow: writeHeaderRow);

            _application.WriteTo(fileStream);
        }


        public Task WriteFromObjectsAsync<D>(IEnumerable<D> objects, Stream fileStream, Options options, Encoding encoding, bool writeHeaderRow = true)
        {
            FillFromObjects(objects, writeHeaderRow: writeHeaderRow);

            _application.WriteTo(fileStream);

            return Task.CompletedTask;
        }




        public IEnumerable<D> SetObjectsFromSheet<D>()
            where D: new()
        {
            if (ColumnTemplate == null)
                throw new ArgumentException("ColumnTemplate not set");

            List<ColumnTemplate> template = GetTemplatePerColumn().ToList();
            
            for (int rowIndex = 1; rowIndex < Cells.RowCount; rowIndex++)
            {
                D dataObj = new D();

                for (int cellIndex = 0; cellIndex < Cells.ColumnCount; cellIndex++)
                {
                    FuncColumnTemplate<D> funcColumn = template[cellIndex] as FuncColumnTemplate<D>;
                    if (funcColumn != null)
                    {
                        var cell = Cells[rowIndex, cellIndex];
                        //cell.NumberFormat = funcColumn.NumberFormat;
                        if (cell.Number.HasValue)
                            funcColumn.SetValue(dataObj, cell.Number);
                        else
                            funcColumn.SetValue(dataObj, cell.Text);
                    }
                }
                yield return dataObj;
            }
        }

        private IEnumerable<ColumnTemplate> GetTemplatePerColumn()
        {
            if (ColumnTemplate.Count >= Cells.ColumnCount)
            {
                foreach (var template in ColumnTemplate)
                {
                    yield return template;
                }
            }
            else
            {
                // read header
                for (int cellIndex = 0; cellIndex < Cells.ColumnCount; cellIndex++)
                {
                    var cell = Cells[0, cellIndex];
                    string header = cell.Text.Trim();
                    var template = ColumnTemplate.FirstOrDefault(t => t.Header == header);
                    yield return template;
                }
            }

        }

        internal Row CreateRowFromObject(int rowNumber, object obj)
        {
            if(ColumnTemplate == null)
                throw new NotSupportedException("ColumnTemplate on sheet not set");

            Row row = new Row(rowNumber,_application.SharedStrings);
            
            int column = 0;
            foreach (ColumnTemplate columnTemplate in ColumnTemplate)
            {
                object value = columnTemplate.GetValue(obj);
                if (value != null)
                {
                    try
                    {
                        Cell cell = row[column];
                        cell.NumberFormat = columnTemplate.NumberFormat;

                        if (value is double doubleValue)
                            cell.Number = (decimal)doubleValue; // First cast is nescesarry for unboxing
                        else if (value is double?)
                        {
                            double? valDb = (double?)value;
                            if(valDb.HasValue)
                                cell.Number = (decimal)valDb.Value; // First cast is nescesarry for unboxing
                        }
                        else if (value is float floatValue)
                            cell.Number = (decimal)floatValue;
                        else if (value is float?)
                        {
                            float? valDb = (float?)value;
                            if (valDb.HasValue)
                                cell.Number = (decimal)valDb.Value; // First cast is nescesarry for unboxing
                        }
                        else if (value is decimal valueDecimal)
                            cell.Number = valueDecimal;
                        else if (value is decimal?)
                        {
                            decimal? valDb = (decimal?)value;
                            if (valDb.HasValue)
                                cell.Number = valDb.Value; // First cast is nescesarry for unboxing
                        }
                        else if(value is DateTime valueDateTime)
                        {
                            cell.Date = valueDateTime;
                        }
                        else if (value is DateTime?)
                        {
                            DateTime? valDb = (DateTime?)value;
                            if (valDb.HasValue)
                                cell.Date = valDb.Value; // First cast is nescesarry for unboxing
                        }
                        else
                        {
                            string text = value.ToString();
                            if (text != null)
                                cell.Text = text;
                        }
                    }
                    catch (InvalidCastException)
                    { }
                    catch(OverflowException)
                    { }
                }

                column++;
            }

            return row;
        }

        private Row GenerateHeader(int rowNumber)
        {
            Row row = new Row(rowNumber,_application.SharedStrings);
            int column = 0;
            foreach (ColumnTemplate columnTemplate in ColumnTemplate)
            {
                if (columnTemplate.Header != null)
                    row[column].Text = columnTemplate.Header;
                column++;
            }

            return row;
        }

        


        internal XDocument WriteInTemplate()
        {
            XDocument xSheet;
            using (Stream sharedStream = _originalPackageSheet.GetStream())
            using (XmlTextReader xmlReader = new XmlTextReader(sharedStream))
            {
                xSheet = XDocument.Load(xmlReader);
            }

            XElement sheetData = xSheet.Descendants(Namespaces.OpenFormat + "sheetData").First();
            Cells.AddToSheetData(sheetData,_application.ConvertNumberFormatToSAttribute);

            return xSheet;
        }


        public void SetWriteStream(Stream stream, Encoding encoding, Options options = null)
        {
            if (options != null)
                _options = options;

            _stream = stream;
            _writer = new StreamWriter(stream, encoding);
        }

        public Task WriteObjectsFinishedAsync()
        {
            return _writer.FlushAsync();
        }

        //[Obsolete("Use CsvSheet")]
        /// <summary>
        /// Also see CsvSheet for an alternative
        /// </summary>
        /// <param name="fileStream"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public Task WriteAsCsvAsync(Stream fileStream, Encoding encoding)
        {
            Options options = new Options();
            options.Seperator = ',';

            return WriteAsCsvAsync(fileStream, options, encoding);
        }

        //[Obsolete("Use CsvSheet")]
        /// <summary>
        /// Also see CsvSheet for an alternative
        /// </summary>
        /// <param name="fileStream"></param>
        /// <param name="options"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public async Task WriteAsCsvAsync(Stream fileStream, Options options, Encoding encoding)
        {
            _options = options;
            // Do not close stream, caller is responsible for closing filestream
            _writer = new StreamWriter(fileStream, encoding);

            for (int r = options.SkipHeader ? 1 : 0; r < Cells.RowCount; r++)
            {
                Row row = Cells.GetRow(r);

                await WriteRowToCsvAsync(row);
            }

            await WriteObjectsFinishedAsync();
        }

        private async Task WriteRowToCsvAsync(Row row)
        {
            bool firstCell = true;
            for (int c = 0; c < row.ColumnCount; c++)
            {
                if (!firstCell)
                    await _writer.WriteAsync(_options.Seperator);
                else
                    firstCell = false;

                var cell = row[c];
                if (cell.Number.HasValue)
                {
                    await _writer.WriteAsync(cell.Number.Value.ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    string text = cell.ToString();
                    if (text.Contains('"') || text.Contains(_options.Seperator) || _options.QuotesAroundText)
                        text = $"\"{text}\"";

                    await _writer.WriteAsync(text);
                }
            }
            if(_options.AddSeperatorOnLineEnd)
                await _writer.WriteAsync(_options.Seperator);
            await _writer.WriteLineAsync();            
        }

        public Task FlushAsync()
        {
            return _writer.FlushAsync();
        }
        

        public void Dispose()
        {
            if (_writer != null)
                _writer.Dispose();

            if(_stream!=null)
                _stream.Dispose();            
        }

    }
}
