using System;
using System.Collections.Generic;
using System.Text;

namespace MontfoortIT.Office.Excel
{
    public class Cell
    {
        private readonly SharedStrings _sharedStrings;

        private string _text;
        private int _sharedIndex = -1;
        private DateTime? _date;
        private bool _dateSet;

        public int SharedIndex
        {
            get
            {
                return _sharedIndex;
            }
        }

        public string Text
        {
            get
            {
                if (_sharedStrings.ForCsv)
                    return _text;

                if (_sharedIndex == -1)
                    return string.Empty;
                return _sharedStrings[_sharedIndex];
            }
            set
            {
                if (_sharedStrings.ForCsv)
                    _text = value;
                else
                    _sharedIndex = _sharedStrings.Add(value);
            }
        }

        public DateTime? Date 
        {
            get
            {
                if (_dateSet)
                    return _date;

                int days;
                if (int.TryParse(Text, out days))
                {
                    try
                    {
                        _date = new DateTime(1900, 1, 1).AddDays(days - 2); // -2 is leap time bug  http://polymathprogrammer.com/2009/10/26/the-leap-year-1900-bug-in-excel/
                        return _date;
                    }
                    catch(ArgumentOutOfRangeException) // sometimes the days is to big
                    {
                        return null;
                    }
                }
                return null;
            }
            set
            {
                _date = value;
                if (value != null)
                {
                    _dateSet = true;
                    NumberFormat = NumberFormat.Date;
                }
            }
        }

        public decimal? Number { get; set; }
        public int Row { get; internal set; }
        public int Column { get; internal set; }
        public NumberFormat NumberFormat { get; internal set; }
        
        internal Cell(SharedStrings sharedStrings)
        {
            if (sharedStrings == null) throw new ArgumentNullException("sharedStrings");

            _sharedStrings = sharedStrings;
            NumberFormat = NumberFormat.Default;
        }

        public static string ToTextRowIndeX(int column)
        {
            List<byte> chars = new List<byte>();

            ToTextCharacters(column, chars);

            return Encoding.ASCII.GetString(chars.ToArray());
        }

        private static int ToTextCharacters(int column, List<byte> chars)
        {
            int iAlpha = column / 27;
            if (iAlpha > 26)
                throw new NotSupportedException("Columns greater than ZZ are not supported");                

            int iRemainder = column - (iAlpha * 27);

            if (iAlpha > 0)
            {
                if (iRemainder >= 26)
                {
                    chars.Add((byte)(iAlpha + 65));
                    chars.Add(65);

                }
                else
                {
                    chars.Add((byte)(iAlpha + 64));
                    chars.Add((byte)(iRemainder + 65));
                }
            }
            else if (iRemainder > 0)
                chars.Add((byte)(iRemainder + 64));

            return iAlpha;
        }

        public override string ToString()
        {
            if (Number.HasValue)
                return Number.ToString();
            if (!string.IsNullOrEmpty(Text))
                return Text;
            if (Date.HasValue)
                return Date.ToString();            
            return "";
        }
        public bool IsEmpty()
        {
            return string.IsNullOrEmpty(ToString());
        }

        internal void MergeValue(Cell cell)
        {
            NumberFormat = cell.NumberFormat;
            if (cell.Number.HasValue)
                Number = cell.Number;
            if (!string.IsNullOrEmpty(cell.Text))
                Text = cell.Text;
            if(cell._dateSet)
                Date = cell.Date;
        }
    }
}
