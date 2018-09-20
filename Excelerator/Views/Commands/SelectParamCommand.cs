namespace Gensler.Revit.Excelerator.Views
{
    using Models;
    using System;
    using System.Windows.Input;

    class SelectParamCommand : ICommand
    {
        private readonly MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var selectedParam = parameter as ParamField;

            _viewModel.SelectedParameter = selectedParam;
        }

        public SelectParamCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
