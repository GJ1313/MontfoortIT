using System;
using System.Linq.Expressions;
using System.Reflection;
using MontfoortIT.Library.Expressions;

namespace MontfoortIT.Library.Merge
{
    public static class MergeExtensions
    {
        public static void Merge<T, V>(this T destination, T source, Expression<Func<T, V>> valueFuncExpr)
        //where T: TableEntityBase
        {
            Func<T, V> valueFunc = valueFuncExpr.Compile();

            V original = valueFunc(destination);
            V def = default(V);

            bool setValue = false;
            if (original is string)
                setValue = string.IsNullOrEmpty((string)(object)original);
            else if (original == null || original.Equals(def))
                setValue = true;
            
            if (setValue)
            {
                V newValue = valueFunc(source);

                if (newValue != null && !newValue.Equals(original))
                {
                    PropertyInfo p = ExpressionFunctions.GetPropertyFromExpression(valueFuncExpr);
                    p.SetValue(destination, newValue);
                }
            }
        }
    }
}
