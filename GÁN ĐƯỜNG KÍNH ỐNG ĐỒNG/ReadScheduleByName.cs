using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Linq;
using Microsoft.VisualBasic;
using Autodesk.Revit.Attributes;
using System.Windows;
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;  // ⚠️ cần thêm reference: Microsoft.VisualBasic.dll
namespace DnBim_Tool

{
    [Transaction(TransactionMode.Manual)]
   
    public class ReadScheduleByName : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (Transaction t = new Transaction(doc, "Tạo bảng Excel ảo readable"))
            {
               
                t.Start();
                //}
                // 🔹 Nếu người dùng đang mở một Schedule → lấy luôn nó
                View activeView = uidoc.ActiveView;
                ViewSchedule schedule = activeView as ViewSchedule;

                // 🔹 Đọc dữ liệu từ schedule
                TableData tableData = schedule.GetTableData();
                TableSectionData tableSection = tableData.GetSectionData(SectionType.Body);
                int r = tableSection.NumberOfRows;
                int c = tableSection.NumberOfColumns;
                Parameter schedulecon = schedule.LookupParameter("SCHEDULE CON");
                List<string> listSchedule = new List<string>();

                if (schedulecon != null)
                {
                    string tenschedulecon = schedulecon.AsString() ?? "";

                    // 🔹 Nếu chuỗi rỗng hoặc null → thêm chuỗi trống
                    if (string.IsNullOrWhiteSpace(tenschedulecon))
                    {
                        listSchedule.Add("");
                    }
                    // 🔹 Nếu không chứa dấu '|' → thêm duy nhất một giá trị
                    else if (!tenschedulecon.Contains("|"))
                    {
                        listSchedule.Add(tenschedulecon.Trim());
                    }
                    // 🔹 Nếu có nhiều giá trị ngăn cách bằng '|'
                    else
                    {
                        // Tách và giữ nguyên thứ tự xuất hiện
                        string[] parts = tenschedulecon.Split('|');
                        foreach (string part in parts)
                        {
                            string trimmed = part.Trim();
                            if (!string.IsNullOrEmpty(trimmed))
                                listSchedule.Add(trimmed);
                        }
                    }
                }
                else
                {
                    // 🔸 Nếu parameter không tồn tại, thêm chuỗi trống để tránh lỗi
                    listSchedule.Add("");
                }

                //// 🔹 In ra kiểm tra thứ tự
                //TaskDialog.Show("📋 Danh sách Schedule con", string.Join("\n", listSchedule));

                Dictionary<int, Dictionary<int, string>> tableMap = new Dictionary<int, Dictionary<int, string>>();
                Dictionary<int, Dictionary<int, string>> tableMapTongongdong = new Dictionary<int, Dictionary<int, string>>();

                string cellText = "";

                // 🔹 Dictionary chính: key = "Cột1|Cột2", value = danh sách các độ dài
                Dictionary<string, List<double>> mergedData = new Dictionary<string, List<double>>();
                Dictionary<string, (string col1, string col2)> rowLabel = new Dictionary<string, (string, string)>();

                // 🔹 Bảng quy chuẩn: key = đường kính ngoài, value = độ dày (mm)
                Dictionary<string, string> Quycachongdong = new Dictionary<string, string>
{
    { "6.35", "0.71" },
    { "9.52", "0.81" },
    { "12.7", "0.81" },
    { "15.88", "1.02" },
    { "19.05", "1.02" },
    { "22.22", "1.02" },
    { "28.58", "1.02" },
    { "34.93", "1.22" },
    { "41.28", "1.42" }
};
                Hashtable pipeTable = new Hashtable();

                for (int i = 1; i < r; i++)
                {
                    string col1 = tableSection.GetCellText(i, 0); // STT MÁY
                    string col2 = tableSection.GetCellText(i, 1); // ĐƯỜNG ỐNG
                    string col3 = schedule.GetCellText(SectionType.Body, i, 2); // ĐỘ DÀI ỐNG
                 

                    // Ép kiểu độ dài
                    if (!double.TryParse(col3, out double length))
                        continue;

                    // 🔹 Dùng Regex tách ra tất cả kích thước số trong chuỗi col2
                    var matches = System.Text.RegularExpressions.Regex.Matches(col2, @"\d+(\.\d+)?");

                    foreach (System.Text.RegularExpressions.Match match in matches)
                    {
                        string size = match.Value; // ví dụ: "19.05" hoặc "41.28"

                        if (pipeTable.ContainsKey(size))
                        {
                            double current = (double)pipeTable[size];
                            pipeTable[size] = current + length; // cộng dồn nếu đã có
                        }
                        else
                        {
                            pipeTable[size] = length; // thêm mới
                        }
                    }
                    // 🔹 Nếu là dòng tiêu đề hoặc dòng trống thì bỏ qua (hoặc xử lý riêng)
                    if (col1.ToUpper().Contains("STT") || col2.ToUpper().Contains("ĐƯỜNG") || col3.ToUpper().Contains("ĐỘ DÀI"))
                    {
                        continue; // không tính dòng tiêu đề
                    }

                    if (string.IsNullOrWhiteSpace(col1) || string.IsNullOrWhiteSpace(col2))
                        continue; // bỏ dòng trống

                    string key = $"{col1}|{col2}";

                   

                    // 🔹 Nếu key đã có → thêm giá trị mới
                    if (mergedData.ContainsKey(key))
                    {
                        mergedData[key].Add(length);
                    }
                    else
                    {
                        mergedData[key] = new List<double> { length };
                        rowLabel[key] = (col1, col2);
                    }
                }  // 🔹 Tạo kết quả hiển thị
                   // 🔹 Chuyển Hashtable sang danh sách và sắp xếp theo đường kính (key)
                var sortedPipeTable = pipeTable.Cast<DictionaryEntry>()
                    .OrderBy(e => double.TryParse(e.Key.ToString(), out double val) ? val : double.MaxValue)
                    .ToList();

                StringBuilder sb1 = new StringBuilder();
                int row = 0;

                foreach (var entry in sortedPipeTable)
                {
                    string size = entry.Key.ToString();
                    double total = ((double)entry.Value) /1000;

                    string thick = Quycachongdong.ContainsKey(size) ? Quycachongdong[size] : "?";
                    string calc = $"= {total:F3}";
                    string label = $"ỐNG ĐỒNG {size} x Φ{thick}mm";

                    var rowDict = new Dictionary<int, string>
    {
        { 0, label },
        { 1, "m" },
        { 2, calc }
    };

                    tableMapTongongdong[row++] = rowDict;
                }

                // ===============================================================
                // 🔹 Chuyển kết quả gộp sang Dictionary<int, Dictionary<int, string>>
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                int rowIndex = 0;

                foreach (var kvp in mergedData)
                {
                    var label = rowLabel[kvp.Key];
                    var values = kvp.Value;
                    double total = values.Sum();

                    string calc = (values.Count == 1)
                        ? (total != 0 ? $"= {total:F0}" : "")
                        : string.Join(" + ", values.Select(v => v.ToString("F0"))) + $" = {total:F0}";

                    // 🔸 Mỗi dòng có 4 cột: col1, col2, col3="Đơn vị", calc
                    var rowDict = new Dictionary<int, string>
    {
        { 0, label.col1 }, // STT MÁY
        { 1, label.col2 }, // ĐƯỜNG ỐNG
         { 2, "mm"}, // Đơn vị
        { 3, calc }        // CỘNG DỒN ĐỘ DÀI
    };

                    tableMap[rowIndex++] = rowDict;
                   
                }



                string c1 = tableSection.GetCellText(0, 0); // STT MÁY
                string c2 = tableSection.GetCellText(0, 1); // ĐƯỜNG ỐNG
                string c3 = tableSection.GetCellText(0, 2); // ĐỘ DÀI ỐNGUnitUtils.ConvertToInternalUnits(maxLength, UnitTypeId.Millimeters)
                string tieude = $"{c1}|{c2}|ĐƠN VỊ|{c3}";
                string tieudeTong = $"ỐNG ĐỒNG |ĐƠN VỊ|ĐỘ DÀI ỐNG";


                Dictionary<string, double> tongChieuDaiTheoLoai = new Dictionary<string, double>();

              

                CREATESCHEDULE_ultis.CREATE_FAKE_SCHEDULE(doc, tableMap, listSchedule[0], tieude);
                CREATESCHEDULE_ultis.CREATE_FAKE_SCHEDULE_Tongongdong(doc, tableMapTongongdong, listSchedule[1], tieudeTong);



                t.Commit();
            } 
                return Result.Succeeded;
        }
    }

}