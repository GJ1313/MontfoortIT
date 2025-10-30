using System;
using System.Xml.Linq;

namespace MontfoortIT.Library.Streams.FileConvertors
{
    public class SemiColonSeperatorToXml: SeperatedFileToXmlBase
    {
        /// <summary>
        /// The seperator used for the file
        /// </summary>
        protected override char Seperator
        {
            get { return ';'; }
        }

        
    }
}
