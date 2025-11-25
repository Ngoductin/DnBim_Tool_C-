using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class AutoNumberGasSystem : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 🟢 1️⃣ Lấy danh sách phần tử đã chọn
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds == null || selectedIds.Count == 0)
                {
                    TaskDialog.Show("Thông báo", "Hãy chọn toàn bộ hệ thống ống gas (Pipe, Fitting, Bộ chia gas)!");
                    return Result.Cancelled;
                }

                // 🟢 2️⃣ Pick ống bắt đầu (ống trục chính)
                Reference ref0 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Chọn ống trục bắt đầu");
                Element start = doc.GetElement(ref0);
                if (start == null)
                {
                    TaskDialog.Show("Lỗi", "Không tìm thấy ống bắt đầu!");
                    return Result.Failed;
                }

                // 🟢 3️⃣ Danh sách toàn hệ thống
                List<Element> allElements = selectedIds
                    .Select(id => doc.GetElement(id))
                    .Where(e => e != null)
                    .ToList();

                using (Transaction t = new Transaction(doc, "Đánh STT và Copy Parameter hệ thống ống gas"))
                {
                    t.Start();

                    // 🔹 Xóa toàn bộ STT cũ
                    foreach (var e in allElements)
                    {
                        Parameter p = e.LookupParameter("STT");
                        if (p != null && !p.IsReadOnly)
                            p.Set(string.Empty);
                    }

                    // 🔹 STT đầu tiên luôn = 1
                    SetSTT(start, "1");
                    CopyPipeDataToParameters(doc, start, "ĐỘ DÀI ỐNG", "SYSTEM TYPE ỐNG ĐỒNG");

                    int counter = 1;
                    TraverseSystem(doc, start, allElements, "1", ref counter);

                    t.Commit();
                }

                TaskDialog.Show("✅ Hoàn tất", "Đã đánh STT và sao chép dữ liệu cho toàn hệ thống ống gas!");
            }
            catch (Exception ex)
            {
                //TaskDialog.Show("⚠️ Lỗi", ex.Message);
            }

            return Result.Succeeded;
        }

        // =====================================================
        // 🔁 HÀM DUYỆT HỆ THỐNG THEO CONNECTOR
        // =====================================================
        private void TraverseSystem(Document doc, Element current, List<Element> allElements, string prefix, ref int counter)
        {
            double currentConMay = GetConMayDouble(current);
            ConnectorSet conns = GetConnectors(current);

            foreach (Connector conn in conns)
            {
                foreach (Connector refConn in conn.AllRefs)
                {
                    Element next = refConn.Owner;
                    if (next == null || next.Id == current.Id) continue;
                    if (!allElements.Any(e => e.Id == next.Id)) continue;
                    if (!string.IsNullOrEmpty(GetSTT(next))) continue;

                    double nextConMay = GetConMayDouble(next);
                    string newSTT;

                    // 🔹 Nếu cùng CON MÁY => tăng cấp cùng nhánh
                    if (nextConMay == currentConMay)
                    {
                        counter++;
                        newSTT = IncrementSTT(prefix);
                    }
                    else
                    {
                        // 🔹 Khác CON MÁY => nhánh phụ, reset đếm
                        int subCounter = 1;
                        newSTT = prefix + "." + subCounter;
                    }

                    // Gán STT và Copy Data
                    SetSTT(next, newSTT);
                    CopyPipeDataToParameters(doc, next, "ĐỘ DÀI ỐNG", "SYSTEM TYPE ỐNG ĐỒNG");

                    // 🔸 Nếu là Bộ chia gas => duyệt sâu nhánh phụ
                    if (next is FamilyInstance fi && fi.Symbol.Family.Name.ToUpper().Contains("BỘ CHIA GAS"))
                    {
                        int subCounter = 1;
                        TraverseSystem(doc, next, allElements, newSTT, ref subCounter);
                    }
                    else
                    {
                        TraverseSystem(doc, next, allElements, newSTT, ref counter);
                    }
                }
            }
        }


        // =====================================================
        // ⚙️ HÀM COPY PARAMETER
        // =====================================================
        

        // =====================================================
        // 🧩 HÀM TIỆN ÍCH
        // =====================================================
        private ConnectorSet GetConnectors(Element e)
        {
            if (e is FamilyInstance fi && fi.MEPModel != null)
                return fi.MEPModel.ConnectorManager?.Connectors;
            if (e is MEPCurve mc)
                return mc.ConnectorManager?.Connectors;
            return new ConnectorSet();
        }

        private double GetConMayDouble(Element e)
        {
            Parameter p = e.LookupParameter("CON MÁY");
            if (p == null || p.StorageType != StorageType.Double)
                return double.NaN;
            return p.AsDouble();
        }

        private void SetSTT(Element e, string value)
        {
            Parameter p = e.LookupParameter("STT");
            if (p != null && !p.IsReadOnly)
                p.Set(value);
        }

        private string GetSTT(Element e)
        {
            Parameter p = e.LookupParameter("STT");
            return p?.AsString() ?? "";
        }

        private string IncrementSTT(string stt)
        {
            var parts = stt.Split('.').Select(int.Parse).ToList();
            parts[parts.Count - 1]++;
            return string.Join(".", parts);
        }

        private void CopyPipeDataToParameters(
       Document doc,
       Element e,
       string targetLengthParam,
       string targetSystemTypePara)
        {
            // ===================== 🟩 1️⃣ ĐỐI TƯỢNG LÀ PIPE =====================
            if (e is Pipe pipe)
            {
                // --- Sao chép Length sang "ĐỘ DÀI ỐNG"
                Parameter len = pipe.LookupParameter("Length");
                Parameter targetLen = pipe.LookupParameter(targetLengthParam);
                if (len != null && targetLen != null && !targetLen.IsReadOnly)
                {
                    double roundedFeet = Math.Round(len.AsDouble() / (5 / 304.8)) * (5 / 304.8); // Làm tròn 5mm
                    if (targetLen.StorageType == StorageType.Double)
                        targetLen.Set(roundedFeet);
                    else
                        targetLen.Set($"{roundedFeet * 304.8:F0} mm");
                }

                // --- Sao chép System Type + Abbreviation
                string sysTypeName = "(Không xác định)";
                string sysAbbr = "";

                // Ưu tiên lấy trực tiếp qua parameter RBS_PIPING_SYSTEM_TYPE_PARAM
                Parameter sysTypeParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                if (sysTypeParam != null && !string.IsNullOrEmpty(sysTypeParam.AsString()))
                {
                    sysTypeName = sysTypeParam.AsString();
                }
                else if (pipe.MEPSystem != null)
                {
                    // Nếu không có parameter trực tiếp → lấy qua MEPSystem
                    MEPSystem sys = pipe.MEPSystem;
                    Element sysTypeElem = doc.GetElement(sys.GetTypeId());
                    sysTypeName = sysTypeElem?.Name ?? "(Không xác định)";

                    Parameter abbrParam = sys.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
                    if (abbrParam != null)
                        sysAbbr = abbrParam.AsString();
                }

                // Gán vào parameter đích
                Parameter pType = pipe.LookupParameter(targetSystemTypePara);

                if (pType != null && !pType.IsReadOnly)
                    pType.Set(sysTypeName);

            }

            // ===================== 🟦 2️⃣ ĐỐI TƯỢNG LÀ FAMILY INSTANCE =====================
            else if (e is FamilyInstance fi)
            {
                Parameter lenParam = fi.LookupParameter(targetLengthParam);
                if (lenParam != null && !lenParam.IsReadOnly)
                {
                    // 🔹 Lấy tên Family
                    string familyName = fi.Symbol.Family.Name.ToUpperInvariant();
                    double newLen = 0;

                    // 🔹 ️Nếu là "BỘ CHIA GAS" → 380 mm
                    if (familyName.Contains("BỘ CHIA GAS"))
                    {
                        newLen = 380;
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
                            lenParam.Set($"{newLen:F0} mm");
                    }
                }

                // --- Sao chép System Type + Abbreviation qua Connector
                ConnectorSet connectors = fi.MEPModel?.ConnectorManager?.Connectors;
                if (connectors != null)
                {
                    foreach (Connector c in connectors)
                    {
                        if (c.MEPSystem != null)
                        {
                            MEPSystem sys = c.MEPSystem;

                            string sysTypeName = "(Không xác định)";
                            string sysAbbr = "";

                            // Thử lấy trực tiếp qua Parameter
                            Parameter sysTypeParam = fi.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                            if (sysTypeParam != null && !string.IsNullOrEmpty(sysTypeParam.AsString()))
                            {
                                sysTypeName = sysTypeParam.AsString();
                            }
                            else
                            {
                                Element sysTypeElem = doc.GetElement(sys.GetTypeId());
                                sysTypeName = sysTypeElem?.Name ?? "(Không xác định)";
                            }

                            Parameter abbrParam = sys.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
                            if (abbrParam != null)
                                sysAbbr = abbrParam.AsString();

                            Parameter pType = fi.LookupParameter(targetSystemTypePara);

                            if (pType != null && !pType.IsReadOnly)
                                pType.Set(sysTypeName);

                            break; // Chỉ cần lấy 1 hệ thống đầu tiên
                        }
                    }
                }
            }
        }

    }
}
