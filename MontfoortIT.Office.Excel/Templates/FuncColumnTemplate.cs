using System;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;

namespace MontfoortIT.Office.Excel.Templates
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

        public FuncColumnTemplate(string fieldName, string format, Expression<Func<T, object>> funcExpr)
            : base(fieldName,format)
        {
            _getFuncExpr = funcExpr;
            _getFunc = funcExpr.Compile();
        }

        public override object GetValue(object o)
        {
            if (o == null)
                return null;

            try
            {
                return _getFunc.Invoke((T)o);
            }
            catch (NullReferenceException) // Sometimes a nullreference occurs because a specified optional property is not set. Then just return a null value
            {
                return null;
            }
            catch(Exception e)
            {
                throw new FuncGetValueException($"{_getFunc} returned an error on object {o}",  e);
            }
        }

        public bool SetValue(T o, object value)
        {
            if (_propertyType == null)
            {
                if(_getFuncExpr.Body is MemberExpression member)
                {
                    SetPropertyType(member);
                }
                                
                UnaryExpression bodyExpression = _getFuncExpr.Body as UnaryExpression;
                if (bodyExpression != null)
                {
                    var propertyExpression = bodyExpression.Operand as MemberExpression;
                    SetPropertyType(propertyExpression);
                }

                if (_propertyType == null)
                {
                    if(this is PropertyNameColumnTemplate<T> pr)
                    {
                        return pr.SetProperty(o, value);
                    }

                    return false;
                }
            }

            var propType = _propertyType.PropertyType;
            if (_propertyType.PropertyType.FullName.StartsWith("System.Nullable"))
            {
                if(value == null ||  value as string == "")
                {
                    _propertyType.SetValue(o, null);
                    return true;
                }

                propType = propType.GetGenericArguments()[0];
            }

            if (propType == typeof(string))
            {
                _propertyType.SetValue(o, Convert.ToString(value));
                return true;
            }

            if (value is string valStr)
            {
                bool ignore = valStr switch
                {
                    "." => true,
                    _ => false
                };
                if (ignore) //
                    return true;
            }

            if (propType == typeof(int))
                _propertyType.SetValue(o, Convert.ToInt32(value));
            else if (propType == typeof(Int16))
                _propertyType.SetValue(o, Convert.ToInt16(value));
            else if (propType == typeof(Boolean))
                _propertyType.SetValue(o, Convert.ToBoolean(value));
            else if (propType == typeof(decimal))
                _propertyType.SetValue(o, Convert.ToDecimal(value, CultureInfo.InvariantCulture));
            else if (propType == typeof(Uri))
            {
                if (value == null)
                {
                    _propertyType.SetValue(o, null);
                }
                else
                {
                    Uri realUri;
                    if (Uri.TryCreate(value.ToString(), UriKind.Absolute, out realUri))
                        _propertyType.SetValue(o, realUri);
                }
            }
            else
                _propertyType.SetValue(o, value);

            return true;
        }

        private void SetPropertyType(MemberExpression propertyExpression)
        {
            if (propertyExpression != null)
            {
                string propertyName = propertyExpression.Member.Name;
                _propertyType = typeof(T).GetProperty(propertyName);
            }
        }
    }
}
