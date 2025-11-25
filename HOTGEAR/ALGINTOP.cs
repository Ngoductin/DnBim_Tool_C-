using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class AlignTagTopCmd : IExternalCommand
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
                    Reference r = uidoc.Selection.PickObject(ObjectType.Element, "Chọn tag hoặc text để align");
                    if (r == null) break;

                    Element tag = doc.GetElement(r);
                    if ((tag is TextNote || tag is IndependentTag) && !selectedTags.Contains(tag))
                        selectedTags.Add(tag);
                    //A
                }

                if (selectedTags.Count < 2)
                {
                    TaskDialog.Show("Thông báo", "Cần chọn ít nhất 2 tag hoặc text hợp lệ.");
                    return Result.Cancelled;
                }

                double GetTopY(Element e)
                {
                    if (e is TextNote tn)
                        return tn.Coord.Y;
                    if (e is IndependentTag tag && tag.TagHeadPosition != null)
                        return tag.TagHeadPosition.Y;
                    BoundingBoxXYZ box = e.get_BoundingBox(null);
                    return box?.Max.Y ?? double.MinValue;//A
                }

                double topY = selectedTags.Max(GetTopY);

                using (Transaction trans = new Transaction(doc, "Align Tags to Top"))
                {
                    trans.Start();
                    //a

                    foreach (var tag in selectedTags)
                    {
                        double currentY = GetTopY(tag);
                        double deltaY = topY - currentY;
                        if (Math.Abs(deltaY) > 1e-6)
                        {
                            XYZ move = new XYZ(0, deltaY, 0);
                            ElementTransformUtils.MoveElement(doc, tag.Id, move);
                        }
                        //A
                    }

                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
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
    //a
}

/// <summary>
/// Command để thực thi alignment
/// </summary>

