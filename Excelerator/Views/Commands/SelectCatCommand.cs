using System.Collections.ObjectModel;
using System.Linq;

namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using Models;
    using System;
    using System.Windows.Input;

    class SelectCatCommand : ICommand
    {
        private readonly MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var category = parameter as Category;
            var document = RevitCommand.RevitDocument;
            var importer = _viewModel.Importer;
            
            var fields = importer.GetSchedulableFields(document, category);
            var paramItems = new ObservableCollection<ParamField>();

            if (fields != null)
                foreach (var field in fields.OrderBy(x => x.GetName(document)))
                    paramItems.Add(new ParamField {Name = field.GetName(document), Field = field});

            _viewModel.SelectedCategory = category;
            _viewModel.ParameterItems = paramItems;
        }

        public SelectCatCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
