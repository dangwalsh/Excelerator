namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.DB;
    using Gensler.Revit.Excelerator.Interfaces;
    using Microsoft.Office.Interop.Excel;
    using System.Collections;

    class Process : IProcess
    {
        Document m_Document;
        ExcelReader m_Reader;
        ScheduleFactory m_Factory;

        public Process(Document document, string path)
        {
            m_Document = document;

            var application = new Application();
            var workbook = application.Workbooks.Open(path);

            application.Visible = true;

            if (workbook != null)
            {
                m_Reader = new ExcelReader(application);
            }
        }

        public void GetSelection()
        {
            if (m_Reader == null) return;

            m_Reader.Select();
        }

        public ViewSchedule CreateSchedule(ElementId category)
        {
            if (m_Document == null) return null;

            m_Factory = new ScheduleFactory();

            return m_Factory.CreateKeySchedule(m_Document, category);
        }

        public void AddFields(ViewSchedule schedule, IEnumerable fields)
        {
            foreach (SchedulableField field in fields)
                m_Factory.AddScheduleField(m_Document, schedule, field);
        }
    }
}
