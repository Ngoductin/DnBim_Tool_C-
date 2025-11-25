using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class AlignTagRightCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                IList<Element> selectedTags = uidoc.Selection.GetElementIds()
                    .Select(id => doc.GetElement(id))
                    .Where(e => e is TextNote || e is IndependentTag)
                    .ToList();

                while (selectedTags.Count < 2)
                {
                    Reference r = uidoc.Selection.PickObject(ObjectType.Element, "Chọn tag hoặc text để align right");
                    if (r == null) break;

                    Element tag = doc.GetElement(r);
                    if ((tag is TextNote || tag is IndependentTag) && !selectedTags.Contains(tag))
                        selectedTags.Add(tag);
                }

                if (selectedTags.Count < 2)
                {
                    TaskDialog.Show("Thông báo", "Cần chọn ít nhất 2 tag hoặc text.");
                    return Result.Cancelled;
                }

                double GetRightX(Element e)
                {
                    BoundingBoxXYZ box = e.get_BoundingBox(uidoc.Document.ActiveView);
                    return box?.Max.X ?? double.MinValue;
                }

                double rightX = selectedTags.Max(GetRightX);

                using (Transaction trans = new Transaction(doc, "Align Tags to Right"))
                {
                    trans.Start();

                    foreach (var tag in selectedTags)
                    {
                        double currentX = GetRightX(tag);
                        double deltaX = rightX - currentX;
                        if (Math.Abs(deltaX) > 1e-6)
                        {
                            XYZ move = new XYZ(deltaX, 0, 0);
                            ElementTransformUtils.MoveElement(doc, tag.Id, move);
                        }
                    }

                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
