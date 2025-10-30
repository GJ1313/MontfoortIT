using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MontfoortIT.Office.Excel.Csv
{
    public class Line
    {
        private Dictionary<int, Column> _columns = new Dictionary<int, Column>();

        public Column this[int column]
        {
            get
            {
                if (_columns.TryGetValue(column, out Column cell))
                    return cell;

                cell = new Column
                {
                    Index = column
                };
                _columns.Add(column, cell);
                
                if (column >= ColumnCount)
                    ColumnCount = column+1;

                return cell;
            }
        }

        public int ColumnCount { get; private set; }
                
    }
}
