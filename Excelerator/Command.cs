namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.Attributes;
    using System.Collections.Generic;
    using System.Linq;
    using System;

    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // TODO: launch UI
            var document = commandData.Application.ActiveUIDocument.Document;
            
            var process = new Process(document, @"C:\Users\22791\Desktop\ImportTest.xlsx");
            var category = new ElementId(BuiltInCategory.OST_Rooms);

            process.GetSelection();

            var schedule = process.CreateSchedule(category);
            var fields = new List<SchedulableField>
            {
                GetField(document, schedule, "Name")
            };

            process.AddFields(schedule, fields);

            return Result.Succeeded;
        }

        SchedulableField CreateField(Document document, ViewSchedule schedule, ElementId parameterId)
        {          
            return new SchedulableField(ScheduleFieldType.Instance, parameterId);
        }

        SchedulableField GetField(Document document, ViewSchedule schedule, string name)
        {
            var schedFields = schedule.Definition.GetSchedulableFields();

            return schedFields.FirstOrDefault(x => x.GetName(document) == name);
        }
    }
}
