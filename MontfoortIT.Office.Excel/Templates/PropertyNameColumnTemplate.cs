using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

namespace MontfoortIT.Office.Excel.Templates
{
    public class PropertyNameColumnTemplate<T> : FuncColumnTemplate<T>
    {
        private PropertyInfo _property;
        private string _name;

        public PropertyNameColumnTemplate(string name)
            : base(name, PropertyByNameFunc(name))
        {

            _name = name;
        }

        private static Expression<Func<T, object>> PropertyByNameFunc(string name)
        {
            var property = typeof(T).GetProperty(name);
            return t => property.GetValue(t);
        }

        public static IEnumerable<ColumnTemplate> GenerateListFromType()
        {
            var props = typeof(T).GetProperties();

            foreach (var prop in props)
            {
                yield return new PropertyNameColumnTemplate<T>(prop.Name);
            }
        }

        internal bool SetProperty(T o, object value)
        {
            if (_property == null)
                _property = typeof(T).GetProperty(_name);

            _property.SetValue(o, value);
            return true;
        }
    }
}
