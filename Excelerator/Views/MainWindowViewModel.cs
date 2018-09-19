namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using Gensler.Revit.Excelerator.Models;
    using System.Collections.ObjectModel;
    using System.ComponentModel;

    class MainWindowViewModel : INotifyPropertyChanged
    {
        string _excelPath;
        Importer _importer;

        ObservableCollection<Category> _categoryItems;
        ObservableCollection<Parameter> _parameterItems;
        ObservableCollection<ExcelItem> _excelItems;     

        public event PropertyChangedEventHandler PropertyChanged;


        public Importer Importer { get; }

        public string ExcelPath
        {
            get { return _excelPath; }
            set
            {
                _excelPath = value;
                _importer = new Importer(value);
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(ExcelPath))); 
            }
        }

        public Category RevitCategory { get; set; }

        public ObservableCollection<Category> CategoryItems
        {
            get { return _categoryItems; }
            set
            {
                _categoryItems = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(CategoryItems)));
            }
        }

        public ObservableCollection<Parameter> ParameterItems
        {
            get { return _parameterItems; }
            set
            {
                _parameterItems = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(ParameterItems)));
            }
        }

        public ObservableCollection<ExcelItem> ExcelItems
        {
            get { return _excelItems; }
            set
            {
                _excelItems = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(ExcelItems)));
            }
        }

        public FileCommand FileCommand { get; set; }

        public SelectCommand SelectCommand { get; set; }

        public AddCommand AddCommand { get; set; }

        public EditCommand EditCommand { get; set; }

        public RunCommand RunCommand { get; set; }

        public MainWindowViewModel()
        {
            FileCommand = new FileCommand(this);
            SelectCommand = new SelectCommand(this);
            AddCommand = new AddCommand(this);
            EditCommand = new EditCommand(this);
            RunCommand = new RunCommand(this);

            var categories = new ObservableCollection<Category>();
            foreach (Category cat in RevitCommand.m_Document.Settings.Categories)
                categories.Add(cat);

            CategoryItems = categories;
        }

        void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, e);
        }
    }
}
