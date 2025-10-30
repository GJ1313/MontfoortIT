using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace MontfoortIT.Office.Excel
{
    public class Workbook
    {
        private SheetCollection _sheets;

        /// <summary>
        ///  
        /// </summary>
        public SheetCollection Sheets
        {
            get
            {
                return _sheets;
            }
        }

        internal Workbook(Application application)
        {
            _sheets = new SheetCollection(application);
        }

        public XDocument CreateDoc()
        {
            XDocument workbookDoc = new XDocument();

            //Obtain a reference to the root node, and then add
            //the XML declaration.

            workbookDoc.Declaration = new XDeclaration("1.0", "UTF-8", "yes");



            //Create and append the workbook node
            //to the document.
            XElement workBook =
                new XElement(Namespaces.OpenFormat + "workbook");

            workbookDoc.Add(workBook);


            //Create and append the sheets node to the 
            //workBook node.
            XElement sheets = new XElement(Namespaces.OpenFormat + "sheets");
            workBook.Add(sheets);


            //Create and append the sheet node to the 
            //sheets node.

            XElement xSheet;
            foreach (Sheet sheet in Sheets)
            {
                xSheet = AddSheetToXml(sheets, sheet);
            }

            return workbookDoc;
        }

        internal static XElement GetSheetsXml(XDocument workbook)
        {
            return workbook.Element(Namespaces.OpenFormat + "workbook").Element(Namespaces.OpenFormat + "sheets");
        }

        internal static XElement AddSheetToXml(XElement sheets, Sheet sheet)
        {
            XElement xSheet = new XElement(Namespaces.OpenFormat + "sheet",
                            new XAttribute("name", sheet.Name),
                            new XAttribute("sheetId", sheet.SheetId),
                            new XAttribute(Namespaces.RFormat + "id", sheet.Id));
            sheets.Add(xSheet);
            return xSheet;
        }
    }
}
