namespace Gensler.Revit.Excelerator
{
    using System;
    using System.Linq;
    using System.Collections;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Microsoft.Office.Interop.Excel;
    using Autodesk.Revit.DB;
    using Models;

    class Importer
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private readonly Application _excelApp;
        private readonly string _excelPath;
        private ExcelItem _excelItem;

        public static List<Category> RevitCategories => ScheduleFacade.GetCategories();

        public Importer(string excelPath)
        {
            _excelPath = excelPath;
            _excelApp = new Application {Visible = false};
        }

        public void SelectData(ExcelItem item)
        {
            var caption = _excelApp.Caption;
            var handler = FindWindow(null, caption);

            SetForegroundWindow(handler);

            _excelApp.Workbooks.Open(_excelPath);
            _excelApp.Visible = true;
            _excelApp.SheetSelectionChange += OnSheetSelectionChanged;
            _excelItem = item;
            _excelItem.ExcelRange = _excelApp.Selection;           
        }

        public void Quit()
        {
            _excelApp.Workbooks.Close();
            _excelApp.Quit();
        }

        public ViewSchedule GetNewSchedule(Document document, Category category)
        {
            if (document == null) return null;

            return ScheduleFacade.GetNewSchedule(document, category.Id);
        }

        public IEnumerable<SchedulableField> GetSchedulableFields(Document document, Category category)
        {
            var schedFields = ScheduleFacade.GetSchedulableFields(document, category)?.Where(x => x.FieldType == ScheduleFieldType.Instance);

            return schedFields;
        }

        public void AddFieldsToSchedule(Document document, ViewSchedule schedule, IEnumerable fields)
        {
            foreach (SchedulableField field in fields)
                ScheduleFacade.AddScheduleField(document, schedule, field);
        }

        public void HideScheduleKeyName(Document document, ViewSchedule schedule)
        {
            ScheduleFacade.HideField(document, schedule, 0);
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

        private List<Dictionary<string, string>> ColumnsToRows(ICollection<ExcelItem> excelItems, int numRows)
        {
            var dataRows = new List<Dictionary<string, string>>();

            for (int i = 0; i < numRows; ++i)
            {
                var dataRow = new Dictionary<string, string>();

                foreach (var item in excelItems)
                {
                    var name = item.RevitParam.Name;
                    if (i >= item.Values.Count)
                    {
                        dataRow.Add(name, "");
                        continue;
                    }

                    var obj = item.Values[i];
                    if (obj is null)
                    {
                        dataRow.Add(name, "");
                        continue;
                    }

                    dataRow.Add(name, obj.ToString());
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
