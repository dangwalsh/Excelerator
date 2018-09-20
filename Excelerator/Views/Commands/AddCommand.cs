namespace Gensler.Revit.Excelerator.Views
{
    using Models;
    using System;
    using System.Windows.Input;

    class AddCommand : ICommand
    {
        private readonly MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var revitParam = _viewModel.SelectedParameter;
            var excelItems = _viewModel.ExcelItems;
            var paramItems = _viewModel.ParameterItems;

            excelItems.Add(new ExcelItem { RevitParam = revitParam });
            paramItems.Remove(revitParam);

            
            _viewModel.NumCols++;
        }

        public AddCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
