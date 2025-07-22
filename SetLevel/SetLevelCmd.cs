using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;


namespace Dnbim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class SetLevelCmd : IExternalCommand
    {
        //a
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            using (Transaction tx = new Transaction(doc, "Gán Level từ Constraints vào Visibility"))
            {
                tx.Start();
                //a
                int count = 0;

                var accessories = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeAccessory)
                    .WhereElementIsNotElementType()
                    .ToElements();
                //a
                foreach (Element elem in accessories)
                {
                    // 1. Lấy parameter "Level" từ group Constraints
                    Parameter levelParam = GetParameterByNameAndGroup(elem, "Level", BuiltInParameterGroup.PG_CONSTRAINTS);
                    if (levelParam == null || !levelParam.HasValue) continue;

                    string levelName = doc.GetElement(levelParam.AsElementId())?.Name ?? "—";

                    // 2. Tìm parameter "Level" thuộc nhóm Visibility
                    Parameter visibilityParam = GetParameterByNameAndGroup(elem, "Level", BuiltInParameterGroup.PG_VISIBILITY);
                    if (visibilityParam != null && !visibilityParam.IsReadOnly)
                    {
                        visibilityParam.Set(levelName);
                        count++;
                    }
                }

                tx.Commit();
                TaskDialog.Show("Xong", $"Đã gán Level cho {count} đối tượng.");
            }

            return Result.Succeeded;
        }
        //a
        private Parameter GetParameterByNameAndGroup(Element element, string paramName, BuiltInParameterGroup targetGroup)
        {
            foreach (Parameter param in element.Parameters)
            {
                Definition def = param.Definition;
                if (def != null && def.Name == paramName && def.ParameterGroup == targetGroup)
                {
                    return param;
                }
                //a
            }
            return null;
        }
    }
}

