namespace Gensler.Revit.Excelerator.Models
{
    using Microsoft.Office.Interop.Excel;
    using System.Collections.Generic;
    using System.ComponentModel;
    using Revit = Autodesk.Revit.DB;

    class ExcelItem : INotifyPropertyChanged
    {
        ParamField _revitParam;
        Range _excelRange;
        List<object> _values;
        int _count;

        public event PropertyChangedEventHandler PropertyChanged;

        public List<object> Values
        {
            get { return _values; }
            set
            {
                _values = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(Values)));
            }
        }

        public int Count
        {
            get { return _count; }
            set
            {
                _count = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(Count)));
            }
        }

        public ParamField RevitParam
        {
            get { return _revitParam; }
            set
            {
                _revitParam = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(RevitParam)));
            }
        }

        public Range ExcelRange
        {
            get { return _excelRange; }
            set
            {
                _excelRange = value;
                _values = GetCellValues(_excelRange);
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(ExcelRange)));
            }
        }

        public void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, e);
        }

        List<object> GetCellValues(Range range)
        {
            object[,] values = (object[,])range.Value2;
            var rows = values.GetLength(0);
            var cols = values.GetLength(1);
            var items = new List<object>();

            for (int i = 1; i <= rows; ++i)
                for (int j = 1; j <= cols; ++j)
                    items.Add(values[i, j]);

            Count = items.Count;

            return items;
        }
    }
}
