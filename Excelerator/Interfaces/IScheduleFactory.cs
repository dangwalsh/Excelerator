namespace Gensler.Revit.Excelerator.Interfaces
{
    using Autodesk.Revit.DB;

    interface IScheduleFactory
    {
        ViewSchedule CreateKeySchedule(Document document, ElementId category);
        void AddScheduleField(Document document, ViewSchedule schedule, SchedulableField field);
    }
}
