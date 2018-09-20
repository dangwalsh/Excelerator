namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;

    using Models;
    using System;
    using System.Collections.Specialized;

    class MainWindowViewModel : INotifyPropertyChanged
    {
        private string _excelPath;
        private Category _selectedCategory;
        private ParamField _selectedParameter;
        private ExcelItem _selectedExcelItem;
        private ObservableCollection<Category> _categoryItems;
        private ObservableCollection<ParamField> _parameterItems;
        private ObservableCollection<ExcelItem> _excelItems;

        public event PropertyChangedEventHandler PropertyChanged;

        public Importer Importer { get; private set; }

        public int NumRows { get; set; }
        public int NumCols { get; set; }

        public string ExcelPath
        {
            get => _excelPath;
            set
            {
                _excelPath = value;
                Importer = new Importer(value);
                OnPropertyChanged(this, new PropertyChangedEventArgs(nameof(ExcelPath))); 
            }
        }

        public Category SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                _selectedCategory = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs(nameof(SelectedCategory)));
            }
        }

        public ParamField SelectedParameter
        {
            get => _selectedParameter;
            set
            {
                _selectedParameter = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs(nameof(SelectedParameter)));
            }
        }

        public ExcelItem SelectedExcelItem
        {
            get => _selectedExcelItem;
            set
            {
                if (_selectedExcelItem != null)
                    _selectedExcelItem.IsActive = false;
                _selectedExcelItem = value;
                _selectedExcelItem.IsActive = true;
                OnPropertyChanged(this, new PropertyChangedEventArgs(nameof(SelectedExcelItem)));
            }
        }

        public ObservableCollection<Category> CategoryItems
        {
            get => _categoryItems;
            set
            {
                _categoryItems = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs(nameof(CategoryItems)));
            }
        }

        public ObservableCollection<ParamField> ParameterItems
        {
            get => _parameterItems ?? (_parameterItems = new ObservableCollection<ParamField>());
            set
            {
                _parameterItems = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs(nameof(ParameterItems)));
            }
        }

        public ObservableCollection<ExcelItem> ExcelItems
        {
            get
            {
                if (_excelItems != null) return _excelItems;

                _excelItems = new ObservableCollection<ExcelItem>();
                _excelItems.CollectionChanged += ExcelItemsCollectionChanged;

                return _excelItems;
            }
            set
            {
                _excelItems = value;
                NumRows = _excelItems.Max(x => x.Count);
                NumCols = _excelItems.Count;
                OnPropertyChanged(this, new PropertyChangedEventArgs(nameof(ExcelItems)));
            }
        }

        private void ExcelItemsCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
                foreach (ExcelItem item in e.NewItems)
                    item.PropertyChanged += OnPropertyChanged;

            if (e.OldItems != null)
                foreach (ExcelItem item in e.OldItems)
                    item.PropertyChanged -= OnPropertyChanged;
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
            foreach (Category cat in RevitCommand.RevitDocument.Settings.Categories)
                categories.Add(cat);

            CategoryItems = categories;
        }

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(sender, e);
        }
    }
}
