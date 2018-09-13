namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.DB;
    using Autodesk.Revit.DB.Architecture;
    using Gensler.Revit.Excelerator.Interfaces;
    using System.Collections.Generic;

    class RoomFactory : IRoomFactory
    {
        Document m_Document;
        Phase m_Phase;

        public List<Room> Rooms { get; set; }

        public RoomFactory(Document document, Phase phase)
        {
            m_Document = document;
            m_Phase = phase;
        }

        public void CreateRooms(string name, int count)
        {
            Rooms = new List<Room>();
            while (count > 0)
            {
                Rooms.Add(CreateRoom(m_Phase));
                count--;
            }
        }

        Room CreateRoom(Phase phase)
        {
            if (m_Document == null)
                throw new System.Exception("You must initialize the RoomFactory before creating a Room.");

            return m_Document.Create.NewRoom(phase);
        }
    }
}
