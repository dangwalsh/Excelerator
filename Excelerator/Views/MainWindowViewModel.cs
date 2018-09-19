namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using Gensler.Revit.Excelerator.Models;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;

    class MainWindowViewModel : INotifyPropertyChanged
    {
        Importer _importer;

        string _excelPath;
        Category _selectedCategory;
        ParamField _selectedParameter;
        ExcelItem _selectedExcelItem;
        int _numRows;
        int _numCols;
       
        ObservableCollection<Category> _categoryItems;
        ObservableCollection<ParamField> _parameterItems;
        ObservableCollection<ExcelItem> _excelItems;

        public event PropertyChangedEventHandler PropertyChanged;


        public Importer Importer
        {
            get { return _importer; }
        }

        public int NumRows { get { return _numRows; } set { _numRows = value; } }
        public int NumCols { get { return _numCols; } set { _numCols = value; } }

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

        public Category SelectedCategory
        {
            get { return _selectedCategory; }
            set
            {
                _selectedCategory = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(SelectedCategory)));
            }
        }

        public ParamField SelectedParameter
        {
            get { return _selectedParameter; }
            set
            {
                _selectedParameter = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(SelectedParameter)));
            }
        }

        public ExcelItem SelectedExcelItem
        {
            get { return _selectedExcelItem; }
            set
            {
                _selectedExcelItem = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(SelectedExcelItem)));
            }
        }

        public ObservableCollection<Category> CategoryItems
        {
            get
            {
                if (_categoryItems == null)
                    _categoryItems = new ObservableCollection<Category>();
                return _categoryItems;
            }
            set
            {
                _categoryItems = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(CategoryItems)));
            }
        }

        public ObservableCollection<ParamField> ParameterItems
        {
            get
            {
                if (_parameterItems == null)
                    _parameterItems = new ObservableCollection<ParamField>();
                return _parameterItems;
            }
            set
            {
                _parameterItems = value;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(ParameterItems)));
            }
        }

        public ObservableCollection<ExcelItem> ExcelItems
        {
            get
            {
                if (_excelItems == null)
                    _excelItems = new ObservableCollection<ExcelItem>();
                return _excelItems;
            }
            set
            {
                _excelItems = value;
                _numRows = _excelItems.Max(x => x.Count);
                _numCols = _excelItems.Count;
                OnPropertyChanged(new PropertyChangedEventArgs(nameof(ExcelItems)));
            }
        }

        public FileCommand FileCommand { get; set; }

        public SelectCatCommand SelectCatCommand { get; set; }

        public SelectParamCommand SelectParamCommand { get; set; }

        public SelectExcelCommand SelectExcelCommand { get; set; }

        public AddCommand AddCommand { get; set; }

        public RemoveCommand RemoveCommand { get; set; }

        public EditCommand EditCommand { get; set; }

        public RunCommand RunCommand { get; set; }

        public MainWindowViewModel()
        {
            FileCommand = new FileCommand(this);
            SelectCatCommand = new SelectCatCommand(this);
            SelectParamCommand = new SelectParamCommand(this);
            SelectExcelCommand = new SelectExcelCommand(this);
            AddCommand = new AddCommand(this);
            RemoveCommand = new RemoveCommand(this);
            EditCommand = new EditCommand(this);
            RunCommand = new RunCommand(this);

            var categories = new ObservableCollection<Category>();
            foreach (Category cat in RevitCommand._Document.Settings.Categories)
                categories.Add(cat);

            CategoryItems = categories;
        }

        void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, e);
        }
    }
}
