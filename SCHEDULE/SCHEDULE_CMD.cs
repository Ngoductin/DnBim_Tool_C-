using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Windows;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class ScheduleKey_Update : IExternalCommand
    {
        //public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        //{
        //    UIDocument uidoc = commandData.Application.ActiveUIDocument;
        //    Document doc = uidoc.Document;

        //    // 🔹 1. Lấy Schedule chứa dữ liệu thực tế
        //    ViewSchedule dataSchedule = new FilteredElementCollector(doc)
        //        .OfClass(typeof(ViewSchedule))
        //        .Cast<ViewSchedule>()
        //        .FirstOrDefault(v => v.Name == "GỘP KHỐI LƯỢNG ỐNG ĐỒNG");

        //    if (dataSchedule == null)
        //    {
        //        TaskDialog.Show("Lỗi", "Không tìm thấy Schedule 'KHỐI LƯỢNG ỐNG ĐỒNG'");
        //        return Result.Failed;
        //    }

        //    TableData tableData = dataSchedule.GetTableData();
        //    TableSectionData body = tableData.GetSectionData(SectionType.Body);

        //    // 🔹 2. Gom nhóm dữ liệu theo đường kính (không phân biệt gas/hơi)
        //    Dictionary<string, double> tongHop = new Dictionary<string, double>();

        //    int rows = body.NumberOfRows;
        //    int colGas = 3; // Cột ĐƯỜNG GAS
        //    int colHoi = 4; // Cột ĐƯỜNG HƠI
        //    int colLen = 2; // Cột ĐỘ DÀI ỐNG

        //    for (int i = 0; i < rows; i++)
        //    {
        //        string gas = body.GetCellText(i, colGas);
        //        string hoi = body.GetCellText(i, colHoi);
        //        string lenStr = body.GetCellText(i, colLen);

        //        if (!double.TryParse(lenStr, out double len))
        //            continue;

        //        AddOrSum(tongHop, gas, len);
        //        AddOrSum(tongHop, hoi, len);
        //    }

        //    // 🔹 3. Lấy Key Schedule đích
        //    ViewSchedule keySchedule = new FilteredElementCollector(doc)
        //        .OfClass(typeof(ViewSchedule))
        //        .Cast<ViewSchedule>()
        //        .FirstOrDefault(v => v.Name == "TỔNG THỐNG KÊ ỐNG ĐỒNG");

        //    if (keySchedule == null)
        //    {
        //        TaskDialog.Show("Lỗi", "Không tìm thấy Key Schedule 'DANH MỤC ỐNG ĐỒNG'");
        //        return Result.Failed;
        //    }

        //    TableSectionData keyBody = keySchedule.GetTableData().GetSectionData(SectionType.Body);

        //    using (Transaction t = new Transaction(doc, "Cập nhật Key Schedule"))
        //    {
        //        t.Start();

        //        // 🔹 Xóa toàn bộ dòng cũ trừ header
        //        for (int r = keyBody.NumberOfRows - 1; r >= 0; r--)
        //        {
        //            //keyBody.RemoveRow(r);
        //        }

        //        // 🔹 4. Ghi dữ liệu tổng hợp
        //        foreach (var kvp in tongHop.OrderBy(k => k.Key))
        //        {
        //            keyBody.InsertRow(keyBody.NumberOfRows);
        //            int newRow = keyBody.NumberOfRows - 1;
        //            keyBody.SetCellText(newRow, 0, $"Ø{kvp.Key}");
        //            keyBody.SetCellText(newRow, 1, $"{Math.Round(kvp.Value, 0)}");
        //        }

        //        t.Commit();
        //    }

        //    TaskDialog.Show("Thành công", $"Đã tổng hợp {tongHop.Count} loại ống đồng vào Key Schedule!");
        //    return Result.Succeeded;
        //}
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 🔹 Tìm schedule có tên cụ thể
            ViewSchedule schedule = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(v => v.Name == "GỘP KHỐI LƯỢNG ỐNG ĐỒNG");

            if (schedule == null)
            {
                TaskDialog.Show("Lỗi", "Không tìm thấy schedule 'KHỐI LƯỢNG ỐNG ĐỒNG'.");
                return Result.Failed;
            }

            // 🔹 Lấy dữ liệu thân bảng
            TableData tableData = schedule.GetTableData();
            TableSectionData body = tableData.GetSectionData(SectionType.Body);

            if (body.NumberOfRows == 0)
            {
                TaskDialog.Show("Thông báo", "Schedule này không có dòng dữ liệu nào.");
                return Result.Cancelled;
            }

            // 🔹 Giả sử cột ĐƯỜNG GAS là cột thứ 3 (index = 3)
            int colGas = 3;
            //string firstGasValue = body.GetCellCalculatedValue(1, colGas);

            // 🔹 Hiển thị kết quả


            return Result.Succeeded;
        }

        // 🔹 Hàm cộng dồn dữ liệu
        private void AddOrSum(Dictionary<string, double> dict, string key, double len)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            key = key.Trim();
            if (!dict.ContainsKey(key))
                dict[key] = 0;
            dict[key] += len;
        }
    }
}
