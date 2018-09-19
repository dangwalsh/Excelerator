namespace Gensler.Revit.Excelerator.Views
{
    using Gensler.Revit.Excelerator.Models;
    using System;
    using System.Windows.Input;

    class SelectExcelCommand : ICommand
    {
        MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

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
