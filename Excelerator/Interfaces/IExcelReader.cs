namespace Gensler.Revit.Excelerator.Interfaces
{
    using System.Collections.Generic;

    interface IExcelReader
    {
        Dictionary<string, double> Data { get; set; }
        void Select();
    }
}
