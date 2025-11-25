using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class UPDATE_DETAIL : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
            if (selectedIds == null || selectedIds.Count == 0)
            {
                TaskDialog.Show("Thông báo", "Hãy chọn ít nhất một đối tượng Annotation có tham số STT!");
                return Result.Cancelled;
            }

            List<Element> selectedElements = selectedIds
                .Select(id => doc.GetElement(id))
                .Where(e => e != null)
                .ToList();

            // 🔹 Lọc tất cả Pipe có "CHI TIẾT DETAIL" khác null
            List<Element> pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Element>()
                .Where(e =>
                {
                    Parameter p = e.LookupParameter("CHI TIẾT DETAIL");
                    return p != null && !string.IsNullOrWhiteSpace(p.AsString());
                })
                .ToList();

            using (Transaction t = new Transaction(doc, "Cập nhật chi tiết Detail"))
            {
                t.Start();

                foreach (Element elem in selectedElements)
                {
                    Parameter detailParam = elem.LookupParameter("STT");
                    if (detailParam == null) continue;

                    string detailValue = detailParam.AsString();
                    Element Mainpipe= null;
                    foreach(Element e in pipes) 
                        {
                        Parameter parachitiet = e.LookupParameter("CHI TIẾT DETAIL");
                        if (parachitiet != null && parachitiet.AsString() == detailValue)
                        {
                            Mainpipe = e;

                            break;
                        }
                    }
                    Parameter KT3 = elem.LookupParameter("KT3");

                   
                    if (Mainpipe != null)
                    {
                        Parameter p = Mainpipe.LookupParameter("CHI TIẾT DETAIL");
                        if (p != null && p.AsString() == detailValue)
                        {
                            //MessageBox.Show(p.AsString());
                            ;
                            Parameter lengthpipe = Mainpipe.LookupParameter("Length");
                            double lengthInMM = DetailUtils.RoundFeetToNearest5mm(lengthpipe.AsDouble());
                            //MessageBox.Show(lengthInMM.ToString());



                            KT3.Set(lengthInMM);
                        }

                    }
                      else
                    {
                        KT3.Set(0);
                    }
                    
                }

                t.Commit();
            }

            TaskDialog.Show("✅ Hoàn tất", "Đã cập nhật chiều dài cho các ống trùng STT thành công!");
            return Result.Succeeded;
        }
    }
}
