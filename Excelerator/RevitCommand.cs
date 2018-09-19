namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.Attributes;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Gensler.Revit.Excelerator.Views;
    using System.Collections.Generic;

    [Transaction(TransactionMode.Manual)]
    public class RevitCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {        
            m_Document = commandData.Application.ActiveUIDocument.Document;

            // TODO: UI launch
            var window = new MainWindow();
            window.Show();

            /*
            // TODO: UI File Open Dialog
            var controller = new Importer(@"C:\Users\22791\Desktop\ImportTest.xlsx");

            // TODO: UI GridView with Text Box containing cell range and button to Select in Excel
            var fieldNames = new List<string> { "Name", "Comments" };
            foreach(var name in fieldNames)
                controller.SelectData(name);

            // TODO: UI Dropdown bound to Categories
            var category = new ElementId(BuiltInCategory.OST_Rooms);
            var schedule = controller.GetNewSchedule(m_Document, category);

            // TODO: UI ListView with available Fields and ListView with Fields to Add, (Add Remove buttons between)
            var fields = new List<SchedulableField>();
            foreach (var name in fieldNames)
                fields.Add(controller.GetSchedulableFields(m_Document, schedule, name));

            controller.AddScheduleFields(m_Document, schedule, fields);

            // TODO: UI Button to execute
            controller.AddScheduleKeys(m_Document, schedule, category);
            controller.AddDataToKeys(m_Document, schedule);
            */

            return Result.Succeeded;
        }

        public static Document m_Document;
    }
}
