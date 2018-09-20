namespace Gensler.Revit.Excelerator.Views
{
    using System;
    using System.Linq;
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
            var excelItems = _viewModel.ExcelItems;
            var paramItems = _viewModel.ParameterItems;

            excelItems.Remove(excelItem);
            paramItems.Add(excelItem.RevitParam);

            _viewModel.NumRows = excelItems.Max(x => x.Count);
            _viewModel.NumCols--;
        }

        public RemoveCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
