namespace Gensler.Revit.Excelerator.Models
{
    using Microsoft.Office.Interop.Excel;
    using System.Collections.Generic;
    using System.ComponentModel;

    class ExcelItem : INotifyPropertyChanged
    {
        private ParamField _revitParam;
        private Range _excelRange;
        private List<object> _values;
        private int _count;
        private bool _isActive;

        public event PropertyChangedEventHandler PropertyChanged;

        public bool IsActive
        {
            get => _isActive;
            set
            {
                _isActive = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(IsActive)));
            }
        }

        public List<object> Values
        {
            get => _values;
            set
            {
                _values = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(Values)));
            }
        }

        public int Count
        {
            get => _count;
            set
            {
                _count = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(Count)));
            }
        }

        public ParamField RevitParam
        {
            get => _revitParam;
            set
            {
                _revitParam = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(RevitParam)));
            }
        }

        public Range ExcelRange
        {
            get => _excelRange;
            set
            {
                _excelRange = value;
                _values = GetCellsByColumn(_excelRange);
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(ExcelRange)));
            }
        }

        public void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, e);
        }

        private List<object> GetCellsByRow(Range range)
        {
            var values = (object[,])range.Value2;
            var rows = values?.GetLength(0);
            var cols = values?.GetLength(1);
            var items = new List<object>();

            for (var i = 1; i <= rows; ++i)
                for (var j = 1; j <= cols; ++j)
                    items.Add(values[i, j]);

            Count = items.Count;

            return items;
        }

        private List<object> GetCellsByColumn(Range range)
        {
            var values = (object[,])range.Value2;
            var rows = values?.GetLength(0);
            var cols = values?.GetLength(1);
            var items = new List<object>();

            for (var j = 1; j <= cols; ++j)
                for (var i = 1; i <= rows; ++i)
                    items.Add(values[i, j]);

            Count = items.Count;

            return items;
        }
    }
}
