
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public static class MyExtension
{
    public static void RemoveEmptyRows(this DataTable dt)
    {
        for (int i = dt.Rows.Count; i >= 1; i--)
        {
            int index = i - 1;
            DataRow row = dt.Rows[index];
            bool isRowEmpty = row.ItemArray.Select(k => (k + "").Trim()).All(k => string.IsNullOrEmpty(k));
            if (isRowEmpty)
                dt.Rows[index].Delete();
        }
        dt.AcceptChanges();
    }
    public static void RemoveEmptyCols(this DataTable dt)
    {
        for (int i = dt.Columns.Count; i >= 1; i--)
        {
            int index = i - 1;
            bool isEmptyColumn = dt.AsEnumerable().All(row => string.IsNullOrEmpty((row[index] + "")));
            if (isEmptyColumn)
            {
                if (dt.Columns[index].ColumnName.StartsWith("F") && dt.Columns[index].ColumnName.Length < 4) // remove empty col
                    dt.Columns.RemoveAt(index);
            }
        }
        dt.AcceptChanges();
    }

}

