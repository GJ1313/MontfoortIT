
using MontfoortIT.Office.Excel.Templates;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace MontfoortIT.Office.Excel.Standard.Templates
{
    public class FuncTemplateList<T>:List<ColumnTemplate>
    {
        public void Add(string fieldName, Expression<Func<T,object>> funcExpr)
        {
            var columnTemplate = new FuncColumnTemplate<T>(fieldName, funcExpr);
            base.Add(columnTemplate);
        }
    }
}
