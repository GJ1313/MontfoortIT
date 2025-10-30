using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace MontfoortIT.Library.Streams.FileConvertors
{
    public abstract class SeperatedFileToXmlBase: SeperatedFileToBase
    {
        


        /// <summary>
        /// Converts a tab file to the xml structure
        /// </summary>
        /// <param name="e"></param>
        public Task ConvertAsync(Stream from, XmlWriter to, Encoding fromEndcoding = null)
        {
            return base.ConvertAsync(from, new StructureWriteToXmlWriter(to), fromEndcoding);            
        }


        
        /// <summary>
        /// Checks if the current row is empty
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static bool IsEmpty(XElement row)
        {
            return row.Elements().All(e => string.IsNullOrEmpty(e.Value));
        }

        /// <summary>
        /// Converts a document to an xdocument
        /// </summary>
        /// <param name="from"></param>
        /// <returns></returns>
        public async Task<XDocument> ConvertToDocumentAsync(Stream from, Encoding fromEncoding = null)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                XmlWriter writer = XmlWriter.Create(stream);

                await ConvertAsync(from, writer, fromEncoding);
                writer.Flush();

                stream.Position = 0;
                return XDocument.Load(XmlReader.Create(stream));
            }
        }
    }
}