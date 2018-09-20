namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using Models;
    using System;
    using System.Collections.ObjectModel;
    using System.Windows.Input;

    class SelectCatCommand : ICommand
    {
        private readonly MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var category = parameter as Category;
            var document = RevitCommand.RevitDocument;
            var importer = _viewModel.Importer;
            var paramItems = _viewModel.ParameterItems;

            var fields = importer.GetSchedulableFields(document, category);           

            foreach (var field in fields)
                paramItems.Add(new ParamField { Name = field.GetName(document), Field = field });

            _viewModel.SelectedCategory = category;
        }

        public SelectCatCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
