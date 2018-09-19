namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using Gensler.Revit.Excelerator.Models;
    using System;
    using System.Windows.Input;

    class SelectParamCommand : ICommand
    {
        MainWindowViewModel _viewModel;

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
