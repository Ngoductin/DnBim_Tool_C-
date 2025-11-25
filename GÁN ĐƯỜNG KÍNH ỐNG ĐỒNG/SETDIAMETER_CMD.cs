using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Collections.Generic;
using Autodesk.Revit.DB.Plumbing;
using System.Globalization;
using System.Windows;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class setdia_cmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // --- Lấy các phần tử được chọn ---
            var selectedIds = uidoc.Selection.GetElementIds();
            if (!selectedIds.Any())
            {
                TaskDialog.Show("Thông báo", "Vui lòng chọn các ống hoặc co/tê cần tách đường kính.");
                return Result.Cancelled;
            }

            List<Element> elementsList = selectedIds
                .Select(id => doc.GetElement(id))
                .Where(e => e is Pipe || e is FamilyInstance)
                .ToList();

            if (!elementsList.Any())
            {
                TaskDialog.Show("Thông báo", "Không tìm thấy ống hoặc fitting nào trong lựa chọn.");
                return Result.Cancelled;
            }

            using (Transaction t = new Transaction(doc, "Tách đường kính ống gas/hơi (double)"))
            {
                t.Start();

                foreach (Element e in elementsList)
                {
                    Parameter dkParam = e.LookupParameter("Đường kính óng gas (hơi)");
                    if (dkParam == null) continue;

                    string value = dkParam.AsString();
                    if (string.IsNullOrWhiteSpace(value) || !value.Contains("/"))
                        continue;

                    // --- Tách chuỗi theo dấu "/" ---
                    string[] parts = value.Split('/');
                    if (parts.Length < 2) continue;

                    string gasStr = parts[0].Trim();
                    string hoiStr = parts[1].Trim();

                    // --- Chuyển về double ---
                    double gasMm, hoiMm;
                    if (!double.TryParse(gasStr, NumberStyles.Any, CultureInfo.InvariantCulture, out gasMm))
                        continue;
                    if (!double.TryParse(hoiStr, NumberStyles.Any, CultureInfo.InvariantCulture, out hoiMm))
                        continue;

                    // --- Quy đổi mm chuẩn danh định ---
                    gasMm = Quydoi(gasMm);
                   
                    hoiMm = Quydoi(hoiMm);
                    string valuegas = gasMm.ToString();
                    string valuehoi = hoiMm.ToString();


                    // --- Gán vào parameter đích ---
                    Parameter pGas = e.LookupParameter("ĐƯỜNG GAS");
                    Parameter pHoi = e.LookupParameter("ĐƯỜNG HƠI");

                    if (pGas != null && !pGas.IsReadOnly && pGas.StorageType == StorageType.String)
                        pGas.Set(valuegas);

                    if (pHoi != null && !pHoi.IsReadOnly && pHoi.StorageType == StorageType.String)
                        pHoi.Set(valuehoi);
                }

                t.Commit();

            }

            //TaskDialog.Show("Hoàn tất", "Đã tách & gán giá trị double cho toàn bộ ống và fitting được chọn.");
            return Result.Succeeded;
        }

        private double Quydoi(double mmValue)
        {
            switch (mmValue)
            {
                case 6.4: return 6.35;
                case 9.5: return 9.52;
                case 15.9: return 15.88;
                case 19.1: return 19.05;
                case 22.2: return 22.22;
                case 28.6: return 28.58;
                case 34.9: return 34.93;
                case 41.3: return 41.28;
                default: return mmValue;
            }
        }
    }
}
