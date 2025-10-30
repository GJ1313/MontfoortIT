using System;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;

namespace MontfoortIT.Library.Templates
{
    public class FuncColumnTemplate<T> : ColumnTemplate
    {
        private readonly Func<T, object> _getFunc;
        private PropertyInfo _propertyType;
        private Expression<Func<T, object>> _getFuncExpr;

        public FuncColumnTemplate(string fieldName, Expression<Func<T, object>> funcExpr)
            : base(fieldName)
        {
            _getFuncExpr = funcExpr;
            _getFunc = funcExpr.Compile();
        }

        public override object GetValue(object o)
        {
            try
            {
                return _getFunc.Invoke((T)o);
            }
            catch (NullReferenceException) // Sometimes a nullreference occurs because a specified optional property is not set. Then just return a null value
            {
                return null;
            }
        }

        public bool SetValue(T o, object value)
        {
            if (_propertyType == null)
            {
                if (_getFuncExpr.Body is UnaryExpression bodyExpression)
                {
                    var propertyExpression = bodyExpression.Operand as MemberExpression;
                    ProcessMemberExpression(propertyExpression);
                }
                if (_getFuncExpr.Body is MemberExpression memberExpression)
                    ProcessMemberExpression(memberExpression);
                
                if (_propertyType == null)
                    return false;
            }

            var propType = _propertyType.PropertyType;
            if (_propertyType.PropertyType.FullName.StartsWith("System.Nullable"))
            {
                if (value == null || value as string == "")
                {
                    _propertyType.SetValue(o, null);
                    return true;
                }

                propType = propType.GetGenericArguments()[0];
            }

            if (propType == typeof(string))
                _propertyType.SetValue(o, Convert.ToString(value));
            else if (propType == typeof(int))
            {
                object clean = CleanValForType(value, propType);
                _propertyType.SetValue(o, Convert.ToInt32(clean));
            }
            else if (propType == typeof(Int16))
            {
                object clean = CleanValForType(value, propType);
                _propertyType.SetValue(o, Convert.ToInt16(clean));
            }
            else if (propType == typeof(Boolean))
            {
                object clean = CleanValForType(value, propType);
                _propertyType.SetValue(o, Convert.ToBoolean(clean));
            }
            else if (propType == typeof(decimal))
            {
                object clean = CleanValForType(value, propType);
                _propertyType.SetValue(o, Convert.ToDecimal(clean, CultureInfo.InvariantCulture));
            }
            else if (propType == typeof(DateTime))
            {
                object clean = CleanValForType(value, propType);
                _propertyType.SetValue(o, Convert.ToDateTime(clean, CultureInfo.InvariantCulture));
            }
            else
                _propertyType.SetValue(o, value);

            return true;
        }

        protected virtual object CleanValForType(object value, Type propType)
        {
            if(value!=null && value is string valueString)
                return valueString.Trim('"');

            return value;
        }

        private void ProcessMemberExpression(MemberExpression propertyExpression)
        {
            if (propertyExpression != null)
            {
                string propertyName = propertyExpression.Member.Name;
                _propertyType = typeof(T).GetProperty(propertyName);
            }
        }
    }
}
