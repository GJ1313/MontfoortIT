namespace MontfoortIT.Office.Excel.Csv
{
    public class Column
    {
        public int Index { get; internal set; }
        public string Text { get; internal set; }

        public decimal? Number { get; set; }
        public NumberFormat NumberFormat { get; internal set; }
    }
}