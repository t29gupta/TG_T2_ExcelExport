using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TG_Reporting
{
    public static class ListExtention
    {
        public static DataTable ToDataTable<T>(this List<T> list)
        {
            var resTable = new DataTable(typeof(T).Name);

            // Create columns in the DataTable
            PropertyInfo[] listProps = typeof(T).GetProperties();
            foreach (var prop in listProps)
            {
                resTable.Columns.Add(new DataColumn(prop.Name, GetColumnDataType(prop)));
            }

            foreach (T t in list)
            {
                DataRow row = resTable.NewRow();
                foreach (PropertyInfo info in listProps)
                {
                    row[info.Name] = info.GetValue(t, null);
                }
                resTable.Rows.Add(row);
            }

            return resTable;
        }

        private static Type GetColumnDataType(PropertyInfo prop)
        {
            return (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ?
                                Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
        }
    }
}
