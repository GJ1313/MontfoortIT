using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MontfoortIT.Library.Expressions
{
    public static class ExpressionFunctions
    {
        //http://stackoverflow.com/questions/17115634/get-propertyinfo-of-a-parameter-passed-as-lambda-expression
        public static PropertyInfo GetPropertyFromExpression<T, V>(Expression<Func<T, V>> getPropertyLambda)
        {
            MemberExpression Exp = null;

            //this line is necessary, because sometimes the expression comes in as Convert(originalexpression)
            if (getPropertyLambda.Body is UnaryExpression)
            {
                var UnExp = (UnaryExpression)getPropertyLambda.Body;
                if (UnExp.Operand is MemberExpression)
                {
                    Exp = (MemberExpression)UnExp.Operand;
                }
                else
                    throw new ArgumentException();
            }
            else if (getPropertyLambda.Body is MemberExpression)
            {
                Exp = (MemberExpression)getPropertyLambda.Body;
            }
            else if (getPropertyLambda.Body is InvocationExpression) //((FieldExpression)(((InvocationExpression)(getPropertyLambda.Body)).Expression)).
            {
                var invocation = (InvocationExpression)getPropertyLambda.Body;
                Exp = (MemberExpression)invocation.Expression;
            }
            else
            {
                throw new ArgumentException();
            }

            string name = Exp.Member.Name;
            return typeof(T).GetProperty(name);

        }
    }
}
