namespace Gensler.Revit.Excelerator.Views
{
    using System;
    using System.Windows.Input;

    class RemoveCommand : ICommand
    {
        private readonly MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var excelItem = _viewModel.SelectedExcelItem;
            if (excelItem == null) return;

            var excelItems = _viewModel.ExcelItems;
            var paramItems = _viewModel.ParameterItems;

            excelItems.Remove(excelItem);
            paramItems.Add(excelItem.RevitParam);
        }

        public RemoveCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
