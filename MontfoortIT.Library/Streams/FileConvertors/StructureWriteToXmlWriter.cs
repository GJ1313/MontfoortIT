using System;
using System.Xml;

namespace MontfoortIT.Library.Streams.FileConvertors
{
    internal class StructureWriteToXmlWriter : IStructureWriter
    {
        private XmlWriter _to;

        public StructureWriteToXmlWriter(XmlWriter to)
        {
            _to = to;
        }

        public void WriteEndColumn()
        {
            _to.WriteEndElement();
        }

        public void WriteEndDocument()
        {
            _to.WriteEndDocument();
        }

        public void WriteEndRow()
        {
            _to.WriteEndElement();
        }

        public void WriteEndTable()
        {
            _to.WriteEndElement();
        }

        public void WriteStartColumn()
        {
            _to.WriteStartElement("column");
        }

        public void WriteStartDocument()
        {
            _to.WriteStartDocument();
        }

        public void WriteStartRow()
        {
            _to.WriteStartElement("row");
        }

        public void WriteStartTable()
        {
            _to.WriteStartElement("table");
        }

        public void WriteString(string text)
        {
            if(!string.IsNullOrEmpty(text))
                text = text.Replace((char)29, ' ');

            _to.WriteString(text);
        }



        //StreamReader streamReader = new StreamReader(from);

        //Encoding utf = Encoding.GetEncoding("ISO-8859-15");

        //to.WriteStartDocument();
        //to.WriteStartElement("table");

        //List<int> excludeChars = ExcludeChars.Select(excludeChar => (int) excludeChar).ToList();

        //int c = streamReader.Read();
        //while (c>0)
        //{
        //    to.WriteStartElement("row");

        //    while (c>=0 && c!= '\n')
        //    {
        //        to.WriteStartElement("column");

        //        List<byte> bytes = new List<byte>();
        //        //StringBuilder strBuilder = new StringBuilder();
        //        while (c >= 0 && (c != Seperator || ContinueRead) && c != '\n')
        //        {
        //            if (excludeChars.Contains(c))
        //                c = streamReader.Read();

        //            char ch = (char)c;

        //            ProcessChar(ch);

        //            if (c != '\r' && c!='\n')
        //            {
        //                if (c == 65533)
        //                    bytes.Add(0xA4); // TODO: Should find out how unicode works here
        //                else
        //                    bytes.Add((byte) c);
        //            }

        //            // strBuilder.Append(ch);

        //            c = streamReader.Read();
        //        }

        //        string text = utf.GetString(bytes.ToArray());
        //        text = CleanText(text);
        //        to.WriteString(text);

        //        to.WriteEndElement();

        //        if(c==Seperator)
        //            c = streamReader.Read();
        //    }

        //    to.WriteEndElement();

        //    if (c == '\n')
        //        c = streamReader.Read();
        //}

        //to.WriteEndElement();
        //to.WriteEndDocument();
    }
}