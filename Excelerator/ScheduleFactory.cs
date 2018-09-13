namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.DB;
    using Gensler.Revit.Excelerator.Interfaces;

    class ScheduleFactory : IScheduleFactory
    {
        public void AddScheduleField(Document document, ViewSchedule schedule, SchedulableField field)
        {
            if (IsAlreadyAdded(schedule, field)) return;

            var transaction = new Transaction(document, "Add Field");
            transaction.Start();

            schedule.Definition.AddField(field);

            transaction.Commit();
        }

        public ViewSchedule CreateKeySchedule(Document document, ElementId category)
        {
            var transaction = new Transaction(document, "Create Key Schedule");
            transaction.Start();

            var schedule = ViewSchedule.CreateKeySchedule(document, category);
            transaction.Commit();

            return schedule;
        }

        bool IsAlreadyAdded(ViewSchedule schedule, SchedulableField field)
        {
            var ids = schedule.Definition.GetFieldOrder();
            var isAdded = false;

            foreach (var id in ids)
                if (schedule.Definition.GetField(id).GetSchedulableField() == field)
                {
                    isAdded = true;
                    break;
                }

            return isAdded;
        }
    }
}
