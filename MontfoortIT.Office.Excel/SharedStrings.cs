using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using MontfoortIT.Library.Extensions;

namespace MontfoortIT.Office.Excel
{
    public class SharedStrings
    {
        private static char[] _INVALID_CHARS = new[] { (char)0x0003, (char)0x00, (char)0x01, (char)0x02, (char)0x06, (char)0x08, (char)0x1A, (char)0x0B,
            (char)0x1F, (char)0x1E }; 
        
        private int _count;
        private readonly List<string> _list = new List<string>();
        private readonly System.Collections.Concurrent.ConcurrentDictionary<string, int> _fastGenerateIndex = new System.Collections.Concurrent.ConcurrentDictionary<string, int>();
        
        internal bool ForCsv { get; set; }

        private object _writeLock = new object();

        public int Add(string text)
        {
            lock (_writeLock)
            {
                _count++;

                string cleanText;
                if (string.IsNullOrEmpty(text))
                    cleanText = text;
                else
                    cleanText = text.MultiReplace(' ', _INVALID_CHARS); // Gives errors in Excel

                if (_fastGenerateIndex.TryGetValue(cleanText, out int index))
                    return index;

                index = AddInternal(cleanText);
                bool success = _fastGenerateIndex.TryAdd(cleanText, index);
                if (success)
                    return index;

                return _fastGenerateIndex[cleanText];
            }
        }

        private int AddInternal(string text)
        {
            lock (_writeLock)
            {
                int index = _list.Count;
                _list.Add(text);

                return index;
            }
        }

        public string this[int sharedIndex]
        {
            get { return _list[sharedIndex]; }
        }

        public XDocument CreateDoc()
        {
            //Create a new XML document for the sharedStrings.
            XDocument sharedStringsDoc = new XDocument();

            //Obtain a reference to the root node, and then add
            //the XML declaration.

            sharedStringsDoc.Declaration = new XDeclaration("1.0", "UTF-8", "yes");

            //Create and append the sst node.
            XElement sstNode = new XElement(Namespaces.OpenFormat + "sst",
                new XAttribute("count", _count),
                new XAttribute("uniqueCount", _list.Count)
                );
            sharedStringsDoc.Add(sstNode);

            foreach (string s in _list)
            {
                //Create and append the si node.
                XElement siNode = new XElement(Namespaces.OpenFormat + "si");
                sstNode.Add(siNode);

                //Create and append the t node.
                XElement tNode = new XElement(Namespaces.OpenFormat + "t");
                tNode.Value = s;
                siNode.Add(tNode);
            }

            return sharedStringsDoc;
        }

        internal void Read(System.IO.Packaging.PackagePart packagePart)
        {
            if (packagePart == null)
                return;

            using (Stream sharedStream = packagePart.GetStream())
            using (XmlTextReader xmlReader = new XmlTextReader(sharedStream)) 
            {
                while (xmlReader.Read())
                {
                    if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "si")
                    {
                        StringBuilder builder = new StringBuilder();
                        xmlReader.Read();
                        while (xmlReader.NodeType != XmlNodeType.EndElement || xmlReader.Name != "si")
                        {
                            if (xmlReader.Name == "t")
                                builder.Append(xmlReader.ReadElementContentAsString());
                            else
                                xmlReader.Read();
                        }

                        AddInternal(builder.ToString());
                    }
                }
            }
        }

        internal void Clean()
        {
            _list.Clear();
            _fastGenerateIndex.Clear();
            _count = 0;
        }
    }
}
