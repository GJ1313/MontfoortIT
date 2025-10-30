using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace MontfoortIT.Office.Excel
{
    internal class SheetReader
    {
        private Sheet _sheet;
        private SharedStrings _sharedStrings;
                
        private int _row;

        public SheetReader(Sheet sheet, SharedStrings sharedStrings)
        {
            _sheet = sheet;
            _sharedStrings = sharedStrings;
        }
        internal void Read(System.IO.Packaging.PackagePart packageSheet, int emptyRowsToIgnore=-1)
        {
            _row = 0;

            int emptyCount = 0;

            using (Stream sharedStream = packageSheet.GetStream())
            using (XmlTextReader xmlReader = new XmlTextReader(sharedStream))
            {
                while (xmlReader.Read())
                {
                    if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "row")
                    {
                        ReadRow(xmlReader);

                        if(emptyRowsToIgnore>-1)
                        {
                            if (RowIsEmpty(_row))
                            {
                                emptyCount++;
                                if (emptyCount >= emptyRowsToIgnore)
                                    return;
                            }
                            else
                                emptyCount = 0;
                        }

                        _row++;
                    }
                }
            }
        }

        private bool RowIsEmpty(int row)
        {
            return _sheet.Cells.GetRow(row).IsEmpty();            
        }

        private void ReadRow(XmlTextReader xmlReader)
        {
            xmlReader.Read();
            while (xmlReader.NodeType != XmlNodeType.EndElement && xmlReader.Name != "row")
            {
                if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "c")
                {
                    string cellPostion = xmlReader.GetAttribute("r");
                    string cellType = xmlReader.GetAttribute("t");
                    string cellContent;
                    if (cellType == "s")
                    {
                        if (xmlReader.IsEmptyElement)
                        {
                            cellContent = "";
                        }
                        else
                        {
                            xmlReader.Read();
                            if (xmlReader.Name == "c")
                                xmlReader.Read();

                            int sharedStringIndex = xmlReader.ReadElementContentAsInt();
                            cellContent = _sharedStrings[sharedStringIndex];
                        }
                    }
                    else
                    { 
                        if (xmlReader.IsEmptyElement)
                            cellContent = "";
                        else
                        {
                            xmlReader.Read();

                            while ((xmlReader.NodeType != XmlNodeType.Element && xmlReader.Name != "v") || xmlReader.Name == "f") // ignore formula
                                xmlReader.Read();

                            string content = "";
                            if (xmlReader.Name == "v")
                                content = xmlReader.ReadElementContentAsString();
                            // Ignore, empty content
                            cellContent = content;
                        }
                    }
                    
                    _sheet.Cells[cellPostion].Text = cellContent;

                }
                
                xmlReader.Read();
            }
        }

    }
}
