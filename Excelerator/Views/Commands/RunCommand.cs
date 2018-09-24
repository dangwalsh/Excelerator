using System.Collections.ObjectModel;
using System.Linq;
using Gensler.Revit.Excelerator.Models;

namespace Gensler.Revit.Excelerator.Views
{
    using Autodesk.Revit.DB;
    using System;
    using System.Collections.Generic;
    using System.Windows.Input;

    class RunCommand : ICommand
    {
        private readonly MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var window = parameter as MainWindow;
            var document = RevitCommand.RevitDocument;
            var category = _viewModel.SelectedCategory;
            var importer = _viewModel.Importer;
            var schedule = importer.GetNewSchedule(document, category);
            var fields = new List<SchedulableField>();
            var numRows = GetRowCount(_viewModel.ExcelItems);
            var numCols = GetColCount(_viewModel.ExcelItems);

            foreach (var item in _viewModel.ExcelItems)
                fields.Add(item.RevitParam.Field);

            importer.AddFieldsToSchedule(document, schedule, fields);
            importer.HideScheduleKeyName(document, schedule);
            importer.AddKeysToSchedule(document, schedule, numRows);
            importer.AddDataToKeys(document, schedule, _viewModel.ExcelItems, numRows, numCols);

            window?.Close();
        }

        public RunCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        private static int GetColCount(IReadOnlyCollection<ExcelItem> excelItems)
        {
            return excelItems.Count;
        }

        private static int GetRowCount(IEnumerable<ExcelItem> excelItems)
        {
            return excelItems.Max(x => x.Count);
        }
    }
}
