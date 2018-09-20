namespace Gensler.Revit.Excelerator.Views
{
    using System;
    using System.Linq;
    using System.Windows.Input;

    class EditCommand : ICommand
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

            _viewModel.Importer.SelectData(excelItem);
            _viewModel.NumRows = _viewModel.ExcelItems.Max(x => x.Count);
        }

        public EditCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
