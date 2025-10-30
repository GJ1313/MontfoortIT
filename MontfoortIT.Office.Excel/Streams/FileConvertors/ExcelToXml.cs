using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace MontfoortIT.Office.Excel.Streams.FileConvertors
{
    public class ExcelToXml
    {
        public XDocument ConvertToDocument(Stream fileToImport, int sheetIndex)
        {
            Application excelApp = new Application();
            excelApp.ReadFile(fileToImport, 1000);

            Sheet sheet = excelApp.Workbook.Sheets[sheetIndex];

            using (MemoryStream stream = new MemoryStream())
            {
                XmlWriter writer = XmlWriter.Create(stream);

                Convert(sheet, writer);
                writer.Flush();

                stream.Position = 0;
                return XDocument.Load(XmlReader.Create(stream));
            }
        }

        private void Convert(Sheet sheet, XmlWriter to)
        {
            Encoding utf = Encoding.GetEncoding("ISO-8859-15");

            to.WriteStartDocument();
            to.WriteStartElement("table");

            for (int row = 0; row < sheet.Cells.RowCount; row++)
            {
                to.WriteStartElement("row");
                
                for (int cell = 0; cell < sheet.Cells.ColumnCount; cell++)
                {
                    to.WriteStartElement("column");
                    var columnText = sheet.Cells[row, cell].ToString();

                    if(IsHeader(row))
                    {
                        if (string.IsNullOrEmpty(columnText))
                            columnText = "EmptyHeader" + cell;
                    }

                    to.WriteString(columnText);
                    to.WriteEndElement();                    
                }

                to.WriteEndElement();                
            }
            
            to.WriteEndElement();
            to.WriteEndDocument();
        }

        private bool IsHeader(int row)
        {
            return row == 0;
        }
    }
}
