using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace MontfoortIT.Office.Excel
{
    public static class Namespaces
    {
        internal static readonly XNamespace OpenFormat = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal static readonly XNamespace RFormat = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        

        internal const string NsWorkbook = "application/vnd.openxmlformats-" +
                            "officedocument.spreadsheetml.sheet.main+xml";
        internal const string WorkbookRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        
        internal const string NsWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

        internal const string WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        internal const string StylesRelationshipNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        internal const string StylesType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

        internal static readonly XNamespace WordProcessingXml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    }
}
