namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using System;
    using System.Collections.ObjectModel;
    using System.Windows.Input;

    class SelectCommand : ICommand
    {
        MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var category = parameter as Category;           
            var paramList = ScheduleFacade.GetParametersInCategory(RevitCommand.m_Document, category);
            var paramItems = new ObservableCollection<Parameter>(paramList);

            _viewModel.ParameterItems = paramItems;
        }

        public SelectCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
