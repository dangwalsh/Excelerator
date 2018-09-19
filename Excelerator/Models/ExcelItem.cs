namespace Gensler.Revit.Excelerator.Models
{
    using Microsoft.Office.Interop.Excel;
    using System.ComponentModel;

    class ExcelItem : INotifyPropertyChanged
    {
        string _name;
        Range _excelRange;

        public event PropertyChangedEventHandler PropertyChanged;

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(Name)));
            }
        }
        public Range Cells
        {
            get { return _excelRange; }
            set
            {
                _excelRange = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(Cells)));
            }
        }

        public void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, e);
        }
    }
}
