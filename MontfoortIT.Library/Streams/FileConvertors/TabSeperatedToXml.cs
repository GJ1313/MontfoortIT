namespace MontfoortIT.Library.Streams.FileConvertors
{
    /// <summary>
    /// Converts a tab seperated file to an xml file in the format table/row/column
    /// </summary>
    public class TabSeperatedToXml : SeperatedFileToXmlBase
    {
        protected override char Seperator
        {
            get { return '\t'; }
        }
    }
}