using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace MontfoortIT.Office.Excel
{
    public class CellCollection
    {
        private readonly SharedStrings _sharedStrings;
        private Dictionary<int, Dictionary<int, Cell>> _cellsPerRow = new Dictionary<int, Dictionary<int, Cell>>();
        private int? _columnCount;
        private Regex _regExpr = new Regex(@"([A-Z]+)(\d+)");

        public CellCollection(SharedStrings sharedStrings)
        {
            if (sharedStrings == null) throw new ArgumentNullException("sharedStrings");

            _sharedStrings = sharedStrings;
        }

        public Cell this[string column]
        {
            get
            {
                int columIndex;
                int row;
                ExcelColumnToColumnRowIndex(column, out row, out columIndex);

                return this[row, columIndex];
            }
        }

        public Cell this[int row, string column]
        {
            get
            {
                int columIndex = ConvertColumnToInt(column.ToUpper());
                return this[row, columIndex];
            }
        }

        public Cell this[int row, int column]
        {
            get
            {
                Dictionary<int, Cell> result;
                Cell cell;
                if (_cellsPerRow.TryGetValue(row, out result) && result.TryGetValue(column, out cell))
                {
                    return cell;
                }

                cell = new Cell(_sharedStrings);
                cell.Row = row;
                cell.Column = column;

                if (column > _columnCount)
                    _columnCount = column; // reload columns

                if (result == null)
                {
                    Dictionary<int, Cell> subCells = new Dictionary<int, Cell>();
                    subCells.Add(column, cell);
                    _cellsPerRow.Add(row, subCells);
                }
                else
                {
                    result.Add(column, cell);
                }
                return cell;
            }
        }

        public void ExcelColumnToColumnRowIndex(string column, out int row, out int columIndex)
        {            
            Match match = _regExpr.Match(column);
            if (match.Success)
            {
                columIndex = ConvertColumnToInt(match.Groups[1].Value);
                row = int.Parse(match.Groups[2].Value) - 1;

            }
            else
                throw new NotSupportedException(string.Format("Column value '{0}' not reckognized", column));
        }

        private static int ConvertColumnToInt(string column)
        {
            int columnCount = 0;
                        

            for (int i = 0; i < column.Length; i++)
            {
                int currentLength = column.Length-i;                

                int firstChar = column[i] - 'A';
                if(currentLength>1)
                    firstChar = (currentLength-1) * 26 * (firstChar+1);
                
                columnCount += firstChar;
            }

            return columnCount;
            
        }

        public int RowCount
        {
            get
            {
                return _cellsPerRow.Count;
            }
        }

        public int ColumnCount
        {
            get
            {
                if(!_columnCount.HasValue)
                {
                    _columnCount = _cellsPerRow.Max(c => c.Value.Keys.Max()) + 1;
                }
                return _columnCount.Value;
            }

        }


        internal void AddToSheetData(XElement sheetData, Func<NumberFormat,bool,string> convertNumberFormatToSAttribute)
        {
            // TODO: This will be very slow in big documents...
            int maxCell = 0;
            if(_cellsPerRow.Count>0)
                maxCell = _cellsPerRow.Select(c=>c.Value.Max(k => k.Key)).Max();

            Dictionary<int, XElement> rowsDictionary = new Dictionary<int, XElement>(); // for faster lookup of a row
            var rowElements = sheetData.Elements(Namespaces.OpenFormat + "row");
            foreach (var rowElement in rowElements)
            {
                int rowIndex;
                if (int.TryParse(rowElement.Attribute("r").Value, out rowIndex))
                    rowsDictionary.Add(rowIndex, rowElement);
            }


            foreach (KeyValuePair<int, Dictionary<int, Cell>> rowPairs in _cellsPerRow.OrderBy(r=>r.Key))
            {
                WriteRowToExcell(sheetData, rowPairs.Key, convertNumberFormatToSAttribute, maxCell, rowsDictionary, rowPairs.Value.Select(s=>s.Value), false);
            }
        }

        internal void WriteRowToExcell(XElement sheetData, Row row, int rowNumber, Func<NumberFormat, bool, string> convertNumberFormatToSAttribute, Dictionary<int, XElement> rowsDictionary, bool isHeader)
        {
            WriteRowToExcell(sheetData, rowNumber, convertNumberFormatToSAttribute, row.ColumnCount, rowsDictionary, row.Cells, isHeader);
        }

        private void WriteRowToExcell(XElement sheetData, int rowNumber, Func<NumberFormat, bool, string> convertNumberFormatToSAttribute, int maxCell, Dictionary<int, XElement> rowsDictionary, IEnumerable<Cell> cells, bool isHeader)
        {
            //Create and add the row node. 
            int row = rowNumber+1;

            XElement rNode;
            if (!rowsDictionary.TryGetValue(row, out rNode))
            {
                rNode = new XElement(Namespaces.OpenFormat + "row",
                    new XAttribute("r", row),
                    new XAttribute("spans", "1:" + Math.Max(maxCell, 1))
                    );
                sheetData.Add(rNode);

                rowsDictionary.Add(row, rNode);
            }

            foreach (Cell cell in cells)
            {
                XElement cNode = GetCnode(rNode, cell.Column, row, cell);

                if (cell.Number.HasValue)
                {
                    cNode.SetAttributeValue("s", convertNumberFormatToSAttribute(cell.NumberFormat, isHeader)); //((int)cell.NumberFormat).ToString()

                    SetCellNodeValue(cNode, cell.Number.Value.ToString(CultureInfo.InvariantCulture));
                }
                else if (cell.Date.HasValue)
                {
                    cNode.SetAttributeValue("s", convertNumberFormatToSAttribute(cell.NumberFormat, isHeader)); // ((int)cell.NumberFormat).ToString()

                    SetCellNodeValue(cNode, ((int)cell.Date.Value.ToOADate()).ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    cNode.SetAttributeValue("t", "s");

                    int sharedIndex = cell.SharedIndex;
                    if (sharedIndex >= 0)
                        SetCellNodeValue(cNode, sharedIndex.ToString());
                }
            }
        }

        internal async Task FillAndAddToSheetDataAsync<D>(Sheet sheet, XElement sheetData, Func<NumberFormat, bool, string> convertNumberFormatToSAttribute, IAsyncEnumerable<D> objects, int rowNumber, Dictionary<int, XElement> rowsDictionary)
        {
            int maxCell = 0;
            await foreach (var obj in objects)
            {
                Row row = sheet.CreateRowFromObject(rowNumber, obj);

                if (maxCell == 0)
                    maxCell = row.ColumnCount;

                // write row...
                WriteRowToExcell(sheetData, rowNumber, convertNumberFormatToSAttribute, maxCell, rowsDictionary, row.Cells, false);

                rowNumber++;

            }
        }

        internal void AddRow(Row row)
        {
            foreach (var cell in row.Cells)
            {
                this[cell.Row, cell.Column].MergeValue(cell);
            }            
        }

        private static void SetCellNodeValue(XElement cNode, string value)
        {
             //Add the "Hello World" text to the worksheet.
            XElement vNode = cNode.Element(Namespaces.OpenFormat + "v");
            if(vNode==null)
            {
                vNode = new XElement(Namespaces.OpenFormat + "v");
                cNode.Add(vNode);
            }
            
            vNode.Value = value;
        }

        private XElement GetCnode(XElement rNode, int column, int row, Cell cell)
        {
            string cellCode = Cell.ToTextRowIndeX(column + 1) + row;

            XElement cNode = rNode.Elements(Namespaces.OpenFormat + "c").Where(c => c.Attribute("r").Value == cellCode).FirstOrDefault();
            if (cNode == null)
            {
                cNode = new XElement(Namespaces.OpenFormat + "c",
                            new XAttribute("r", Cell.ToTextRowIndeX(column + 1) + row),
                            new XAttribute("s", ((int)cell.NumberFormat).ToString())
                            );
                rNode.Add(cNode);
            }

            return cNode;
        }

        internal Row GetRow(int rowNumber)
        {
            Row row = new Row(rowNumber,_sharedStrings);
            row.ColumnCount = ColumnCount;

            Dictionary<int, Cell> cells;
            if(_cellsPerRow.TryGetValue(rowNumber, out cells))
            {
                foreach (var cellKv in cells)
                {
                    row[cellKv.Key].MergeValue(cellKv.Value);
                }
            }
            return row;
        }

    }
}
