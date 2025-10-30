using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace MontfoortIT.Office.Excel.Templates
{
    public class DataTableConverter
    {
        public static System.Data.DataTable Convert<T>(IEnumerable<T> items, IEnumerable<FuncColumnTemplate<T>> columnTemplates)
        {
            List<FuncColumnTemplate<T>> columns = columnTemplates.ToList();

            DataTable table = new DataTable();
            foreach (var columnTemplate in columns)
            {
                table.Columns.Add(columnTemplate.Header);
            }

            foreach (var item in items)
            {
                object[] values = new object[columns.Count];
                for (int i = 0; i < columns.Count; i++)
                {
                    values[i] = columns[i].GetValue(item);
                }

                table.Rows.Add(values);
            }

            return table;
        }
    }
}
