namespace Gensler.Revit.Excelerator.Views
{
    using System;
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
        }

        public EditCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
