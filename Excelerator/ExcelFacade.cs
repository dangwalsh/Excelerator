namespace Gensler.Revit.Excelerator
{
    using Gensler.Revit.Excelerator.Models;
    using Microsoft.Office.Interop.Excel;
    using System.Collections.Generic;

    static class ExcelFacade
    {
        public static Column GetSelection(Application application, string name)
        {
            var selection = application.Selection as Range;

            return EnumerateSelection(name, selection.Cells);
        }

        static Column EnumerateSelection(string name, Range range)
        {
            object[,] values = (object[,])range.Value2;
            var rows = values.GetLength(0);
            var cols = values.GetLength(1);
            var items = new List<object>();
            var column = new Column()
            {
                Name = name
            };

            for (int i = 1; i <= rows; ++i)
                for (int j = 1; j <= cols; ++j)
                    items.Add(values[i, j]);

            column.Items = items;
            column.Count = items.Count;

            return column;
        }
    }
}
