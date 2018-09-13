namespace Gensler.Revit.Excelerator
{
    using Gensler.Revit.Excelerator.Interfaces;
    using Microsoft.Office.Interop.Excel;
    using System.Collections.Generic;

    class ExcelReader : IExcelReader
    {
        Application m_Application;

        public Dictionary<string, double> Data { get; set; }

        public ExcelReader(Application application)
        {
            m_Application = application;
        }

        public void Select()
        {
            var selection = m_Application.Selection as Range;

            ParseSelection(selection.Cells);
        }

        void ParseSelection(Range range)
        {
            Data = new Dictionary<string, double>();

            object[,] values = (object[,])range.Value2;
            var row = values.GetLength(0);
            var col = values.GetLength(1);

            for (int i = 1; i <= row; ++i)
                for (int j = 1; j <= col; ++j)
                {
                    string name = (string)values[i, j];
                    double value = (double)values[i, ++j];
                    Data.Add(name, value);
                }
        }
    }
}
