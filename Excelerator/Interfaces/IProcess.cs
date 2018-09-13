namespace Gensler.Revit.Excelerator.Interfaces
{
    using Autodesk.Revit.DB;
    using System.Collections;

    interface IProcess
    {
        ViewSchedule CreateSchedule(ElementId category);
        void AddFields(ViewSchedule schedule, IEnumerable fields);
    }
}
