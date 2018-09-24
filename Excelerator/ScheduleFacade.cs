using System;

namespace Gensler.Revit.Excelerator
{
    using Autodesk.Revit.DB;
    using Autodesk.Revit.Exceptions;
    using System.Collections.Generic;
    using System.Linq;

    static class ScheduleFacade
    {
        public static List<Category> GetCategories()
        {
            var document = RevitCommand.RevitDocument;
            var revitCats = document.Settings.Categories;

            var catList = new List<Category>();
            foreach (Category revitCat in revitCats)
                catList.Add(revitCat);
                
            return catList;
        }

        public static void AddScheduleField(Document document, ViewSchedule schedule, SchedulableField field)
        {
            if (IsAlreadyAdded(schedule, field)) return;

            var transaction = new Transaction(document, "Add Field");
            transaction.Start();

            schedule.Definition.AddField(field);

            transaction.Commit();
        }

        public static void HideField(Document document, ViewSchedule schedule, int fieldIndex)
        {
            var transaction = new Transaction(document, "Add Field");
            transaction.Start();

            var field = schedule.Definition.GetField(fieldIndex);
            field.IsHidden = true;

            transaction.Commit();
        }

        public static ViewSchedule GetNewSchedule(Document document, ElementId category)
        {
            var transaction = new Transaction(document, "Create Key Schedule");
            transaction.Start();

            var schedule = ViewSchedule.CreateKeySchedule(document, category);
            
            transaction.Commit();

            return schedule;
        }

        public static IList<SchedulableField> GetSchedulableFields(Document document, Category category)
        {
            var transaction = new Transaction(document, "Create Key Schedule");
            transaction.Start();

            ViewSchedule schedule = null;
            try
            {
                schedule = ViewSchedule.CreateKeySchedule(document, category.Id);
            }
            catch (ArgumentException)
            {

            }

            if (schedule == null)
            {
                transaction.RollBack();
                return null;
            }

            var schedFields = schedule.Definition.GetSchedulableFields();

            transaction.RollBack();

            return schedFields;

        }

        public static void AddScheduleKey(Document document, ViewSchedule schedule)
        {
            var transaction = new Transaction(document, "Add Row");
            transaction.Start();
            
            var tableData = schedule.GetTableData();
            var section = tableData.GetSectionData(SectionType.Body);
            section.InsertRow(section.FirstRowNumber);

            transaction.Commit();
        }

        public static List<Element> GetScheduleKeys(Document document, ViewSchedule schedule)
        {
            var collector = new FilteredElementCollector(document);
            var keys = collector
                .WhereElementIsNotElementType()
                .Where(x => x.OwnerViewId == schedule.Id).ToList();

            return keys;
        }

        public static void AddDataToKeys(Document document, List<Dictionary<string, string>> dataRows, List<Element> keys)
        {
            var transaction = new Transaction(document, "Update Parameters");
            transaction.Start();

            for (int i = 0; i < keys.Count; ++i)
            {
                var key = keys[i];
                var dataRow = dataRows[i];

                foreach (Parameter p in key.Parameters)
                    if (dataRow.ContainsKey(p.Definition.Name))
                    {
                        var dataCell = dataRow[p.Definition.Name];
                        var paraType = p.Definition.ParameterType;

                        if (paraType is ParameterType.Text)
                        {
                            p.Set(dataCell);
                            continue;
                        }

                        double.TryParse(dataCell, out var d);
                        p.Set(d);
                    }
                        
            }

            transaction.Commit();
        }

        static bool IsAlreadyAdded(ViewSchedule schedule, SchedulableField field)
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
