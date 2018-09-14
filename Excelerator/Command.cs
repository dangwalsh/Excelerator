namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.Attributes;
    using System.Collections.Generic;

    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {        
            var document = commandData.Application.ActiveUIDocument.Document;

            // TODO: UI launch

            // TODO: UI File Open Dialog
            var controller = new DataController(document, @"C:\Users\22791\Desktop\ImportTest.xlsx");

            // TODO: UI GridView with Text Box containing cell range and button to Select in Excel
            var fieldNames = new List<string> { "Name", "Comments" };
            foreach(var name in fieldNames)
                controller.SelectData(name);

            // TODO: UI Dropdown bound to Categories
            var category = new ElementId(BuiltInCategory.OST_Rooms);
            var schedule = controller.GetNewSchedule(document, category);

            // TODO: UI ListView with available Fields and ListView with Fields to Add, (Add Remove buttons between)
            var fields = new List<SchedulableField>();
            foreach (var name in fieldNames)
                fields.Add(controller.GetSchedulableFields(document, schedule, name));

            controller.AddScheduleFields(document, schedule, fields);

            // TODO: UI Button to execute
            controller.AddScheduleKeys(document, schedule, category);
            controller.AddDataToKeys(document, schedule);

            return Result.Succeeded;
        }

        public void MockInterface()
        {

        }
    }
}
