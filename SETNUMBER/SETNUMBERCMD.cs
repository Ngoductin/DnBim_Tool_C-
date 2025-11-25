using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class SETNUMBERCMD : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)

        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 🟢 1. Lấy danh sách phần tử đang được chọn
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds == null || selectedIds.Count == 0)
                {
                    TaskDialog.Show("Thông báo", "Hãy chọn ít nhất một đối tượng MEP (Pipe, Fitting, Accessory)!");
                    return Result.Cancelled;
                }
                Reference ref0 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Chọn ống trục");
                Element elementstart = doc.GetElement(ref0);
                List<Element> selectedElements = selectedIds
                    .Select(id => doc.GetElement(id))
                    .Where(e => e != null)
                    .ToList();

                // 🟢 2. Sắp xếp lại theo liên kết Connector trong phạm vi các phần tử được chọn
                List<Element> sortedElements = SortByConnectorWithinSelection(selectedElements, elementstart);

                //// 🟢 3. Hiển thị danh sách preview
                //string preview = string.Join("\n", sortedElements
                //    .Select((e, i) => $"{i + 1}. [{e.Category?.Name}] {e.Name}  (ID: {e.Id.IntegerValue})"));

                //TaskDialogResult confirm = TaskDialog.Show(
                //    "Danh sách theo thứ tự liền mạch",
                //    preview + "\n\nBạn có muốn đánh STT theo thứ tự này?",
                //    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                //    TaskDialogResult.Yes);

                //if (confirm == TaskDialogResult.No)
                //    return Result.Cancelled;
                //a
                // 🟢 4. Gán STT vào parameter
                using (Transaction t = new Transaction(doc, "Đánh STT cho đối tượng đã chọn"))
                {
                    t.Start();
                    int counter = 1;
                    foreach (Element e in sortedElements)
                    {
                        Parameter sttParam = e.LookupParameter("STT")
                          ;

                        if (sttParam != null && !sttParam.IsReadOnly)
                        {
                            if (sttParam.StorageType == StorageType.String)
                                sttParam.Set(counter.ToString());

                            counter++;
                        }

                        CopyPipeDataToParameters(doc,e, "ĐỘ DÀI ỐNG", "SYSTEM TYPE ỐNG ĐỒNG", "SYSTEM ABBREVIATION");

                    }
                    t.Commit();
                }

                //TaskDialog.Show("Hoàn tất", "Đã gán STT theo thứ tự liền mạch!");
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi", ex.Message);
                return Result.Failed;
            }
        }

        // 🔧 Sắp xếp list đã chọn theo thứ tự kết nối
        public static List<Element> SortByConnectorWithinSelection(List<Element> selectedElements, Element startElement)
        {
            Queue<Element> queue = new Queue<Element>();
            HashSet<ElementId> visited = new HashSet<ElementId>();
            List<Element> ordered = new List<Element>();

            queue.Enqueue(startElement);
            visited.Add(startElement.Id);

            while (queue.Count > 0)
            {
                Element current = queue.Dequeue();
                ordered.Add(current);

                ConnectorSet conns = GetConnectors(current);
                foreach (Connector conn in conns)
                {
                    foreach (Connector refConn in conn.AllRefs)
                    {
                        Element next = refConn.Owner;
                        // 🔹 chỉ xét nếu nằm trong danh sách được chọn
                        if (next != null && selectedElements.Any(x => x.Id == next.Id) && !visited.Contains(next.Id))
                        {
                            visited.Add(next.Id);
                            queue.Enqueue(next);
                        }
                    }
                }
            }

            // Nếu có phần tử cô lập chưa duyệt, thêm vào cuối
            foreach (var e in selectedElements)
            {
                if (!visited.Contains(e.Id))
                    ordered.Add(e);
            }

            return ordered;
        }

        // 🔧 Lấy connector của phần tử MEP
        public static ConnectorSet GetConnectors(Element e)
        {
            if (e is MEPCurve mc)
                return mc.ConnectorManager.Connectors;
            if (e is FamilyInstance fi && fi.MEPModel != null)
                return fi.MEPModel.ConnectorManager.Connectors;
            return new ConnectorSet();
        }
        private void CopyPipeDataToParameters(
     Document doc,
     Element e,
     string targetLengthParam,
     string targetSystemTypeParam,
     string targetSystemAbbrParam)
        {
            // --- 1️⃣ Nếu là ỐNG (Pipe) ---
            if (e is Pipe pipe)
            {
                // 🟩 Sao chép Length sang "ĐỘ DÀI ỐNG"
                Parameter len = pipe.LookupParameter("Length");
                Parameter targetLen = pipe.LookupParameter(targetLengthParam);
                if (len != null && targetLen != null && !targetLen.IsReadOnly)
                {
                    double roundedFeet = DetailUtils.RoundFeetToNearest5mm(len.AsDouble());
                    if (targetLen.StorageType == StorageType.Double)
                        targetLen.Set(roundedFeet);
                    else
                        targetLen.Set($"{roundedFeet * 304.8:F0} mm");
                }

                // 🟩 Sao chép System Type & Abbreviation
                MEPSystem system = pipe.MEPSystem;
                if (system != null)
                {
                    ElementId sysTypeId = system.GetTypeId();
                    Element sysTypeElem = doc.GetElement(sysTypeId);
                    string sysTypeName = sysTypeElem?.Name ?? "(Không xác định)";

                    string sysAbbr = "";
                    Parameter abbrParam = system.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
                    if (abbrParam != null)
                        sysAbbr = abbrParam.AsString();

                    Parameter pType = pipe.LookupParameter(targetSystemTypeParam);
                    Parameter pAbbr = pipe.LookupParameter(targetSystemAbbrParam);

                    if (pType != null && !pType.IsReadOnly)
                        pType.Set(sysTypeName);

                    if (pAbbr != null && !pAbbr.IsReadOnly)
                        pAbbr.Set(sysAbbr ?? "");
                }
            }

            // --- 2️⃣ Nếu là Phụ kiện / Co / Van / Bộ chia v.v. (FamilyInstance) ---
            else if (e is FamilyInstance fi)
            {
                Parameter lenParam = fi.LookupParameter(targetLengthParam);
                if (lenParam != null && !lenParam.IsReadOnly)
                {
                    // 🔹 Lấy tên Family
                    string familyName = fi.Symbol.Family.Name.ToUpperInvariant();
                    double newLen = 0;

                    // 🔹 ️Nếu là "BỘ CHIA GAS" → 380 mm
                    if (familyName.Contains("Fitting-Joint"))
                    {
                        newLen = 380/304.8;
                        
                    }
                    else
                    {
                        // 🔹 Lấy góc nếu có
                        Parameter angleParam = fi.LookupParameter("ANGLE") ?? fi.Symbol.LookupParameter("ANGLE");
                        if (angleParam != null)
                        {
                            double angleDeg = angleParam.AsDouble() * (180 / Math.PI);
                            newLen = (angleDeg > 55) ? 245 : 120;
                        }
                        else
                        {
                            // 🔹 Nếu không có góc, gán mặc định 120 mm
                            newLen = 120;
                        }
                    }

                    // 🔹 Nếu đã có giá trị trước đó → chỉ ghi đè nếu khác biệt đáng kể
                    bool needUpdate = true;
                    if (lenParam.HasValue)
                    {
                        double currentFeet = lenParam.StorageType == StorageType.Double
                            ? lenParam.AsDouble()
                            : 0;
                        double currentMM = currentFeet * 304.8;
                        if (Math.Abs(currentMM - newLen) < 1.0)
                            needUpdate = false;
                    }

                    if (needUpdate)
                    {
                        if (lenParam.StorageType == StorageType.Double)
                            lenParam.Set(newLen / 304.8);
                        else
                            lenParam.Set($"{newLen} mm");
                    }
                }

               
                }
            }
        }

    }


