namespace MontfoortIT.Office.Excel.Csv
{
    public class Options
    {
        public bool QuotesAroundText { get; set; }
        public bool SkipHeader { get; set; } = false;
        public char Seperator { get; set; } = ',';
        public bool AddSeperatorOnLineEnd { get; set; } = false;
    }
}