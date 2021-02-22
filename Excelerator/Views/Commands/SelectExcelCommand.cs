namespace Gensler.Revit.Excelerator.Views
{
    using Models;
    using System;
    using System.Windows.Input;

    class SelectExcelCommand : ICommand
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
            var selectedExcel = parameter as ExcelItem;

            _viewModel.SelectedExcelItem = selectedExcel;
        }

        public SelectExcelCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
