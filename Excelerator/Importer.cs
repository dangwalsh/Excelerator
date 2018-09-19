namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.DB;
    using Gensler.Revit.Excelerator.Models;
    using Microsoft.Office.Interop.Excel;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;

    class Importer
    {
        Application _excelApp;
        ExcelItem _excelItem;
        

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

        public void OnSheetSelectionChanged(object sheet, Range range)
        {
            _excelItem.ExcelRange = range;

            _excelApp.Visible = false;
            _excelApp.SheetSelectionChange -= OnSheetSelectionChanged;
        }

        public ViewSchedule GetNewSchedule(Document document, Category category)
        {
            if (document == null) return null;

            return ScheduleFacade.GetNewSchedule(document, category.Id);
        }

        public IList<SchedulableField> GetSchedulableFields(Document document, ViewSchedule schedule)
        {
            var schedFields = schedule.Definition.GetSchedulableFields();

            return schedFields;
        }

        public SchedulableField GetSchedulableFields(Document document, ViewSchedule schedule, string name)
        {
            var schedFields = schedule.Definition.GetSchedulableFields();

            return schedFields.FirstOrDefault(x => x.GetName(document) == name);
        }

        public void AddScheduleFields(Document document, ViewSchedule schedule, IEnumerable fields)
        {
            foreach (SchedulableField field in fields)
                ScheduleFacade.AddScheduleField(document, schedule, field);
        }

        public void AddScheduleKeys(Document document, ViewSchedule schedule, int count)
        {
            for (int i = 0; i < count; ++i)
                ScheduleFacade.AddScheduleKey(document, schedule);
        }

        public void AddDataToKeys(Document document, ViewSchedule schedule, ICollection<ExcelItem> excelItems, int numRows, int numCols)
        {
            var keys = ScheduleFacade.GetScheduleKeys(document, schedule);
            var dataRows = ColumnsToRows(excelItems, numRows, excelItems.Count);

            ScheduleFacade.AddDataToKeys(document, dataRows, keys);
        }

        List<Dictionary<string, string>> ColumnsToRows(ICollection<ExcelItem> excelItems, int numRows, int numCols)
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
    }
}
