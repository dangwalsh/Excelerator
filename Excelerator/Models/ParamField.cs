

namespace Gensler.Revit.Excelerator.Models
{
    using Autodesk.Revit.DB;
    using System.ComponentModel;

    class ParamField : INotifyPropertyChanged
    {
        string _name;
        SchedulableField _field;

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(Name)));
            }
        }

        public SchedulableField Field
        {
            get { return _field; }
            set
            {
                _field = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(Field)));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, e);
        }
    }
}
