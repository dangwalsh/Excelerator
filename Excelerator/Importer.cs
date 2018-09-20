namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.DB;
    using Models;
    using Microsoft.Office.Interop.Excel;
    using System.Collections;
    using System.Collections.Generic;

    class Importer
    {
        private readonly Application _excelApp;
        private ExcelItem _excelItem;
        

        public Importer(string path)
        {
            _excelApp = new Application();
            _excelApp.Workbooks.Open(path);
        }
        public void SelectData(ExcelItem item)
        {
            _excelItem = item;
            _excelApp.Visible = true;
            _excelApp.SheetSelectionChange += OnSheetSelectionChanged;
        }

        public ViewSchedule GetNewSchedule(Document document, Category category)
        {
            if (document == null) return null;

            return ScheduleFacade.GetNewSchedule(document, category.Id);
        }

        public IList<SchedulableField> GetSchedulableFields(Document document, Category category)
        {
            var schedFields = ScheduleFacade.GetSchedulableFields(document, category);

            return schedFields;
        }

        public void AddFieldsToSchedule(Document document, ViewSchedule schedule, IEnumerable fields)
        {
            foreach (SchedulableField field in fields)
                ScheduleFacade.AddScheduleField(document, schedule, field);
        }

        public void AddKeysToSchedule(Document document, ViewSchedule schedule, int count)
        {
            for (int i = 0; i < count; ++i)
                ScheduleFacade.AddScheduleKey(document, schedule);
        }

        public void AddDataToKeys(Document document, ViewSchedule schedule, ICollection<ExcelItem> excelItems, int numRows, int numCols)
        {
            var keys = ScheduleFacade.GetScheduleKeys(document, schedule);
            var dataRows = ColumnsToRows(excelItems, numRows);

            ScheduleFacade.AddDataToKeys(document, dataRows, keys);
        }

        List<Dictionary<string, string>> ColumnsToRows(ICollection<ExcelItem> excelItems, int numRows)
        {
            var dataRows = new List<Dictionary<string, string>>();

            for (int i = 0; i < numRows; ++i)
            {
                var dataRow = new Dictionary<string, string>();

                foreach (var item in excelItems)
                {
                    var name = item.RevitParam.Name;
                    dataRow.Add(name, item.Values[i] as string);
                }

                dataRows.Add(dataRow);
            }

            return dataRows;
        }

        private void OnSheetSelectionChanged(object sheet, Range range)
        {
            _excelItem.ExcelRange = range;
            _excelApp.Visible = false;
            _excelApp.SheetSelectionChange -= OnSheetSelectionChanged;
        }
    }
}
