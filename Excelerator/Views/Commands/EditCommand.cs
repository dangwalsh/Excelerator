namespace Gensler.Revit.Excelerator.Views
{
    using Gensler.Revit.Excelerator.Models;
    using System;
    using System.Windows.Input;

    class EditCommand : ICommand
    {
        private MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var item = parameter as ExcelItem;

            _viewModel.Importer.SelectData(item.Name);
        }

        public EditCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
