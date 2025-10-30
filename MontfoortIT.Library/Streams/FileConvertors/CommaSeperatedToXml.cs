using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MontfoortIT.Library.Streams.FileConvertors
{
    ///<summary>
    /// Transfers a comma seperated file to xml, csv
    ///</summary>
    public class CommaSeperatedToXml: SeperatedFileToXmlBase
    {
        private bool _continueReadBlockStarted = false;

        protected override char Seperator
        {
            get { return ','; }
        }

        protected override bool ContinueRead
        {
            get
            {
                return _continueReadBlockStarted;
            }
        }

        protected override char[] ExcludeChars
        {
            get
            {
                return new[]{'\\'};
            }
        }


        protected override bool ProcessChar(char ch)
        {
            if(ch == '"') // Do not split in " blocks
            {
                _continueReadBlockStarted = !_continueReadBlockStarted;
                return false;                    
            }

            return base.ProcessChar(ch);
        }
    }
}