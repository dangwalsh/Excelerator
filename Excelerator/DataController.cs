namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.DB;
    using Gensler.Revit.Excelerator.Models;
    using Microsoft.Office.Interop.Excel;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;

    class DataController
    {
        Application m_ExcelApp;
        List<Column> m_DataColumns;

        public DataController(Document document, string path)
        {
            m_DataColumns = new List<Column>();
            m_ExcelApp = new Application();

            m_ExcelApp.Workbooks.Open(path);
        }

        public void SelectData(string name)
        {
            m_ExcelApp.Visible = true;
            m_DataColumns.Add(ExcelFacade.GetSelection(m_ExcelApp, name));
            m_ExcelApp.Visible = false;
        }

        public ViewSchedule GetNewSchedule(Document document, ElementId category)
        {
            if (document == null) return null;

            return ScheduleFacade.GetNewSchedule(document, category);
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

        public void AddScheduleKeys(Document document, ViewSchedule schedule, ElementId category)
        {
            for (int i = 0; i < m_DataColumns[0].Count; ++i)
                ScheduleFacade.AddScheduleKey(document, schedule);
        }

        public void AddDataToKeys(Document document, ViewSchedule schedule)
        {
            var keys = ScheduleFacade.GetScheduleKeys(document, schedule);
            var dataRows = ColumnsToRows(m_DataColumns);

            ScheduleFacade.AddDataToKeys(document, dataRows, keys);
        }

        List<Dictionary<string, string>> ColumnsToRows(List<Column> dataColumns)
        {
            var numCols = dataColumns.Count;
            var numRows = dataColumns[0].Count;      
            var dataRows = new List<Dictionary<string, string>>();

            for (int i = 0; i < numRows; ++i)
            {
                var dataRow = new Dictionary<string, string>();

                foreach (var col in dataColumns)
                {
                    var name = col.Name;
                    dataRow.Add(name, col.Items[i] as string);
                }

                dataRows.Add(dataRow);
            }

            return dataRows;
        }
    }
}
