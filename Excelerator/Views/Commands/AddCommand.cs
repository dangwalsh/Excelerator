namespace Gensler.Revit.Excelerator.Views
{
    using System;
    using System.Windows.Input;

    class AddCommand : ICommand
    {
        MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            throw new NotImplementedException();
        }

        public AddCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
