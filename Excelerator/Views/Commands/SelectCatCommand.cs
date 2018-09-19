namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using Gensler.Revit.Excelerator.Models;
    using System;
    using System.Collections.ObjectModel;
    using System.Windows.Input;

    class SelectCatCommand : ICommand
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
            var document = RevitCommand._Document;
            var importer = _viewModel.Importer;

            //var paramList = ScheduleFacade.GetParametersInCategory(document, category);
            var schedule = importer.GetNewSchedule(document, category);
            var fields = importer.GetSchedulableFields(document, schedule);
            var paramItems = new ObservableCollection<ParamField>();

            foreach (var field in fields)
                paramItems.Add(new ParamField { Name = field.GetName(document), Field = field });

            _viewModel.ParameterItems = paramItems;
            _viewModel.SelectedCategory = category;
        }

        public SelectCatCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
