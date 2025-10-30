using System.Linq;
using System.Collections.Generic;
using System;

namespace MontfoortIT.Office.Excel
{
    internal class Row
    {
        private Dictionary<int, Cell> _cells = new Dictionary<int, Cell>();
        private int? _columnCount;
        private int _rowNumber;
        private SharedStrings _sharedStrings;

        internal IEnumerable<Cell> Cells
        {
            get
            {
                return _cells.Select(c => c.Value);
            }
        }

        public int ColumnCount
        {
            get
            {
                if (!_columnCount.HasValue)
                    _columnCount = _cells.Keys.Max()+1;
                return _columnCount.Value;
            }
            internal set { _columnCount = value; }
        }
        

        public Cell this[int column]
        {
            get
            {
                Cell cell;
                if (_cells.TryGetValue(column, out cell))
                {
                    return cell;
                }

                cell = new Cell(_sharedStrings);
                cell.Row = _rowNumber;
                cell.Column = column;

                if (!_columnCount.HasValue || column >= _columnCount)
                    _columnCount = column+1; // reload columns
                
                _cells.Add(column, cell);
                
                return cell;
            }
        }

        internal Row(int rowNumber, SharedStrings sharedStrings)
        {            
            _sharedStrings = sharedStrings;
            _rowNumber = rowNumber;
        }

        internal bool IsEmpty()
        {
            return _cells.Values.All(c => c.IsEmpty());
        }
    }
}