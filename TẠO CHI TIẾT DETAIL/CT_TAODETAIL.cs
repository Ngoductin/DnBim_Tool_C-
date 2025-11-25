using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace DnBim_Tool
{
    public static class DetailUtils
    {
        /// <summary>
        /// 🔹 Hàm tìm giá trị lớn nhất (Max) của một parameter bất kỳ trong toàn bộ model.
        /// Hỗ trợ các kiểu: Double, Integer, String (tự parse sang số).
        /// </summary>
        /// <param name="doc">Tài liệu Revit hiện tại (Document)</param>
        /// <param name="paramName">Tên parameter cần tìm</param>
        /// <returns>Giá trị lớn nhất (double). Nếu không tìm thấy, trả về 0.</returns>
        public static double TIMGIATRIDETAILMAX(Document doc, string paramName)
        {
            try
            {
                if (doc == null || string.IsNullOrWhiteSpace(paramName))
                    return 0;

                // --- 1️⃣ Thu thập toàn bộ các phần tử MEP cần duyệt ---
                var collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .Where(e =>
                        e.Category != null &&
                        (
                            e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves ||
                            e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting ||
                            e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeAccessory ||
                            e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves ||
                            e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctFitting ||
                            e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory ||
                            e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_FlexDuctCurves
                        ));

                if (!collector.Any())
                    return 0;

                // --- 2️⃣ Thu thập dữ liệu (ID + ValueText) ---
                var elementData = new List<string>();

                foreach (var e in collector)
                {
                    Parameter p = e.LookupParameter(paramName);
                    if (p == null) continue;

                    string txt = "";
                    switch (p.StorageType)
                    {
                        case StorageType.Double:
                            txt = p.AsDouble().ToString(CultureInfo.InvariantCulture);
                            break;
                        case StorageType.Integer:
                            txt = p.AsInteger().ToString();
                            break;
                        case StorageType.String:
                            txt = p.AsString();
                            break;
                    }

                    if (!string.IsNullOrWhiteSpace(txt))
                        elementData.Add(txt);
                }

                if (elementData.Count == 0)
                    return 0;

                // --- 3️⃣ Chạy song song (Parallel LINQ) để tìm giá trị lớn nhất ---
                double maxVal = 0;
                Task.Run(() =>
                {
                    maxVal = elementData
                        .AsParallel()
                        .Select(v =>
                        {
                            if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                                return d;
                            return 0;
                        })
                        .DefaultIfEmpty(0)
                        .Max();
                }).Wait();

                return maxVal;
            }
            catch (Exception)
            {
                return 0;
            }
        }
        public static double RoundFeetToNearest5mm(double feetValue)
        {
            double mm = feetValue * 304.8;
            mm = Math.Round(mm / 5.0) * 5;   // làm tròn bội 5 mm
            return mm / 304.8;               // trả về feet để set vào Revit
        }
    }
}
