using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace MontfoortIT.Office.Excel
{
    class Styles
    {
        public XDocument CreateDoc()
        {
            XDocument docr = XDocument.Parse("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"><numFmts count=\"1\"><numFmt numFmtId=\"165\" formatCode=\"#,##0.00\"/></numFmts><fonts count=\"1\" x14ac:knownFonts=\"1\"><font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/><scheme val=\"minor\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"2\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf borderId=\"0\" fillId=\"0\" fontId=\"0\" numFmtId=\"14\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtind=\"0\"/></cellStyles><dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/><extLst><ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/></ext></extLst></styleSheet>");
            return docr;

            


            ////Create a new XML document for the sharedStrings.
            //XDocument doc = new XDocument();

            ////Obtain a reference to the root node, and then add
            ////the XML declaration.

            //doc.Declaration = new XDeclaration("1.0", "UTF-8", "yes");

            //// <numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;€&quot;\ #,##0.00"/></numFmts>

            ////Create and append the sst node.
            //XElement sstNode = new XElement(Namespaces.OpenFormat + "styleSheet");
            //doc.Add(sstNode);

            //XElement numFormats = new XElement(Namespaces.OpenFormat + "numFmts",
            //    new XAttribute("count", 1));
            //sstNode.Add(numFormats);

            //XElement numFormat = new XElement(Namespaces.OpenFormat + "numFmt"
            //    ,new XAttribute("numFmtId", 164)
            //    , new XAttribute("formatCode", "\"€\"\\ #,##0.00")
            //    );
            //numFormats.Add(numFormat);

            //XElement cellXfs = new XElement(Namespaces.OpenFormat + "cellXfs", 
            //    new XAttribute("count","2"),
            //    new XElement(Namespaces.OpenFormat + "xf",
            //        new XAttribute("numFmtId", "0"),
            //        new XAttribute("fontId", "0"),
            //        new XAttribute("fillId", "0"),
            //        new XAttribute("borderId", "0"),
            //        new XAttribute("xfId", "0")
            //        ),
            //    new XElement(Namespaces.OpenFormat + "xf",
            //        new XAttribute("numFmtId", "165"),
            //        new XAttribute("fontId", "0"),
            //        new XAttribute("fillId", "0"),
            //        new XAttribute("borderId", "0"),
            //        new XAttribute("xfId", "0"),
            //        new XAttribute("applyNumberFormat", "1")
            //        )
            //    );
            //sstNode.Add(cellXfs);
            ////XElement.Parse("<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>"));

            //return doc;
        }
    }
}
