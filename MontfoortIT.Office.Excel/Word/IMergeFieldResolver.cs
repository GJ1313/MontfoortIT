using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MontfoortIT.Office.Excel.Word
{
    public interface IMergeFieldResolver
    {
        string GetFieldValue(string field);
    }
}
