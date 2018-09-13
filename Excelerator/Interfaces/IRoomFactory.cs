namespace Gensler.Revit.Excelerator.Interfaces
{
    using Autodesk.Revit.DB;
    using Autodesk.Revit.DB.Architecture;

    interface IRoomFactory
    {
        void CreateRooms(string name, int count);
    }
}
