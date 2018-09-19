namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using System;
    using System.Collections.Generic;
    using System.Windows.Input;

    class RunCommand : ICommand
    {
        MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var window = parameter as MainWindow;
            var document = RevitCommand._Document;
            var category = _viewModel.SelectedCategory;
            var importer = _viewModel.Importer;
            var schedule = importer.GetNewSchedule(document, category);
            var fields = new List<SchedulableField>();

            foreach (var item in _viewModel.ExcelItems)
                fields.Add(item.RevitParam.Field);

            importer.AddScheduleFields(document, schedule, fields);
            importer.AddScheduleKeys(document, schedule, _viewModel.NumRows);
            importer.AddDataToKeys(document, schedule, _viewModel.ExcelItems, _viewModel.NumRows, _viewModel.NumCols);

            window.Close();
        }

        public RunCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
