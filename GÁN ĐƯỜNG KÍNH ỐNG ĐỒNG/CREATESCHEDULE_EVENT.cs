using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Document = Autodesk.Revit.DB.Document;
using System.Windows;

namespace DnBim_Tool
{
    public static  class CREATESCHEDULE_ultis

    {
        public static void CREATE_FAKE_SCHEDULE(Document doc, Dictionary<int, Dictionary<int, string>> tableMap, string name,string tieude)
        {
           
                // 🔹 1. Tạo schedule trống (Generic Model)
                ElementId catId = new ElementId(BuiltInCategory.OST_GenericModel);
                ViewSchedule sched;

            bool themdong = true;
            bool xoadongcuoi = false;

                // 🧱 2. Nếu người dùng không nhập tên → tạo schedule mới (Table1, Table2, ...)
                if (string.IsNullOrWhiteSpace(name))
                {
                    // Lấy toàn bộ schedule hiện có
                    var allschedules = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSchedule))
                        .Cast<ViewSchedule>()
                        .ToList();

                    // Đếm bao nhiêu schedule bắt đầu bằng "Table"
                    int count  = allschedules.Count(v => v.Name.StartsWith("Table", StringComparison.OrdinalIgnoreCase));

                    string newName = $"Table{count + 1}";

                    sched = ViewSchedule.CreateSchedule(doc, catId);
                    sched.Name = newName;

                    TaskDialog.Show("🆕 Tạo mới", $"Đã tạo schedule mới: {newName}");
                }
                else
                {
                    // 🧱 3. Người dùng đã nhập tên → tìm trong project xem có tồn tại không
                    sched = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSchedule))
                        .Cast<ViewSchedule>()
                        .FirstOrDefault(v => v.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                themdong = false;
                xoadongcuoi = true;
                if (sched == null)
                    { themdong=false;
                    // Nếu không tìm thấy, tạo mới với tên đó
                    sched = ViewSchedule.CreateSchedule(doc, catId);
                        sched.Name = name;
                        themdong= true;
                    xoadongcuoi = false;
                    
                }
                    else
                    {
                        
                    }
                }
            sched.Definition.ShowHeaders = false;
            // 🔹 2. Gán text style
            ElementId textTypeId = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .FirstOrDefault(e => e.Name.Contains("2.5mm Arial"))?.Id;

            if (textTypeId != null)
            {
                // Áp dụng cho title, header, body
                sched.TitleTextTypeId = textTypeId;
                sched.HeaderTextTypeId = textTypeId;
                sched.BodyTextTypeId = textTypeId;
            }




            // 🔹 Lấy dữ liệu bảng
            TableData tableData = sched.GetTableData();




                // 🔹 3. Số dòng và cột
                // 🧱 Giả sử bạn đã có schedule
                TableData table = sched.GetTableData();
                TableSectionData header = table.GetSectionData(SectionType.Header);

             

                // 🔹 Kiểm tra số cột và set lại độ rộng cột đầu tiên (cột A)
                // 🔹 3. Xóa hoặc chỉnh riêng ô đầu tiên (Assembly Code)
                if (header.NumberOfColumns > 0)
                {
                    // Xóa nội dung ô “Assembly Code”
                    header.ClearCell(0, 0);

                    // Đặt lại độ rộng cột đầu tiên (300 mm)
                    header.SetColumnWidth(0, 30 / 304.8);
                }
                //A
               

                // 🧱 Xóa dữ liệu cũ (nếu cần)
                for (int i = 0; i < header.NumberOfRows; i++)
                {
                    for (int j = 0; j < header.NumberOfColumns; j++)
                    {
                        header.ClearCell(i, j);
                    }
                }

            // 🧱 3️⃣ Lấy số hàng/cột tối đa trong dictionary
            int maxRow = tableMap.Keys.Max();
            int maxCol = tableMap.Values.Max(dict => dict.Keys.Max());
            if(themdong == true)
            {
                for (int c = 0; c < maxCol; c++)
                    header.InsertColumn(c);
                for (int r = 0; r < maxRow; r++)
                    header.InsertRow(r);
            }
            //MessageBox.Show($"Max Row: {maxRow}, Max Col: {maxCol}");
           
            
           


            foreach (var rowPair in tableMap)
            {
                int rowIndex = rowPair.Key;          // 🔹 Hàng hiện tại
                var rowDict = rowPair.Value;         // 🔹 Dữ liệu của hàng đó

                // 🔹 Lấy dữ liệu từng cột
                string col1 = rowDict.ContainsKey(0) ? rowDict[0] : "";
                string col2 = rowDict.ContainsKey(1) ? rowDict[1] : "";
                string col3 = rowDict.ContainsKey(3) ? rowDict[3] : "";

                // 🔹 Đảm bảo schedule có đủ hàng
              

                // 🔹 Ghi dữ liệu vào đúng dòng
                header.SetCellText(rowIndex, 0, col1);
                header.SetCellText(rowIndex, 1, col2);
                header.SetCellText(rowIndex, 3, col3);
            }

            for (int r = 0; r < tableMap.Count; r++)
            {
                foreach (var col in tableMap[r])
                {
                    header.SetCellText(r, col.Key, col.Value);
                }
            }

            // 🔹 Duyệt toàn bộ ô trong Header

            header.InsertRow(0);
            // 🟩 Tách chuỗi dựa theo ký tự phân tách "|"
            string[] parts = tieude.Split('|');

            // 🟦 Loại bỏ khoảng trắng thừa (nếu có)
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = parts[i].Trim();
            }

            // 🧩 Đưa từng phần vào các ô trong header
            for (int j = 0; j < parts.Length; j++)
            {
                header.SetCellText(0, j, parts[j]); // 0 là hàng đầu tiên
            }
            
            try
            {
                // 🔹 8. Thêm field Assembly Code
                ScheduleDefinition definition = sched.Definition;
                SchedulableField assemblyField = null;

                foreach (SchedulableField f in definition.GetSchedulableFields())
                {
                    if (f.GetName(doc) == "Assembly Code")
                    {
                        assemblyField = f;
                        break;
                    }
                }
                // 🔹 Lấy định nghĩa schedule


                if (assemblyField != null)
                {
                    definition.AddField(assemblyField);
                }
                else
                {
                    TaskDialog.Show("⚠️ Cảnh báo", "Không tìm thấy field Assembly Code trong category Generic Model.");
                }


                // 🔹 9. Thêm 2 filter (Filter by Assembly Code)
                if (assemblyField != null)
                {
                    ScheduleField field = definition.GetField(0); // Assembly Code
                    ScheduleFilter filter1 = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, "NO VALUES FOUND");
                    ScheduleFilter filter2 = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, "ALL VALUES FOUND");

                    definition.AddFilter(filter1);
                    definition.AddFilter(filter2);
                }
              
              
            }

            catch
            {

            }
            if (xoadongcuoi == true)
            {
                MessageBox.Show("Xóa dòng cuối");
                int dongcanxoa = header.NumberOfRows-1 ;
                header.RemoveRow(dongcanxoa);
            }

        }
        public static void CREATE_FAKE_SCHEDULE_Tongongdong(Document doc, Dictionary<int, Dictionary<int, string>> tableMapTongongdong, string name, string tieude)
        {

            // 🔹 1. Tạo schedule trống (Generic Model)
            ElementId catId = new ElementId(BuiltInCategory.OST_GenericModel);
            ViewSchedule sched;
            bool themdong = true;
            bool xoadongcuoi = false;

            // 🧱 2. Nếu người dùng không nhập tên → tạo schedule mới (Table1, Table2, ...)
            if (string.IsNullOrWhiteSpace(name))
            {
                // Lấy toàn bộ schedule hiện có
                var allschedules = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSchedule))
                    .Cast<ViewSchedule>()
                    .ToList();

                // Đếm bao nhiêu schedule bắt đầu bằng "Table"
                int count = allschedules.Count(v => v.Name.StartsWith("Table", StringComparison.OrdinalIgnoreCase));

                string newName = $"Table{count + 1}";

                sched = ViewSchedule.CreateSchedule(doc, catId);
                sched.Name = newName;

                TaskDialog.Show("🆕 Tạo mới", $"Đã tạo schedule mới: {newName}");
            }
            else
            {
                // 🧱 3. Người dùng đã nhập tên → tìm trong project xem có tồn tại không
                sched = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSchedule))
                    .Cast<ViewSchedule>()
                    .FirstOrDefault(v => v.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                themdong = false;
                xoadongcuoi = true;

                if (sched == null)
                {
                    // Nếu không tìm thấy, tạo mới với tên đó
                    sched = ViewSchedule.CreateSchedule(doc, catId);
                    sched.Name = name;
                    themdong = true;
                    xoadongcuoi = false;

                }
                else
                {

                }
            }
            sched.Definition.ShowHeaders = false;
            // 🔹 2. Gán text style
            ElementId textTypeId = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .FirstOrDefault(e => e.Name.Contains("2.5mm Arial"))?.Id;

            if (textTypeId != null)
            {
                // Áp dụng cho title, header, body
                sched.TitleTextTypeId = textTypeId;
                sched.HeaderTextTypeId = textTypeId;
                sched.BodyTextTypeId = textTypeId;
            }




            // 🔹 Lấy dữ liệu bảng
            TableData tableData = sched.GetTableData();




            // 🔹 3. Số dòng và cột
            // 🧱 Giả sử bạn đã có schedule
            TableData table = sched.GetTableData();
            TableSectionData header = table.GetSectionData(SectionType.Header);



            // 🔹 Kiểm tra số cột và set lại độ rộng cột đầu tiên (cột A)
            // 🔹 3. Xóa hoặc chỉnh riêng ô đầu tiên (Assembly Code)
            if (header.NumberOfColumns > 0)
            {
                // Xóa nội dung ô “Assembly Code”
                header.ClearCell(0, 0);
                

                // Đặt lại độ rộng cột đầu tiên (300 mm)
                header.SetColumnWidth(0, 30 / 304.8);
            }


            // 🧱 Xóa dữ liệu cũ (nếu cần)
            for (int i = 0; i < header.NumberOfRows; i++)
            {
                for (int j = 0; j < header.NumberOfColumns; j++)
                {
                    header.ClearCell(i, j);
                    header.ResetCellOverride(i, j);
                }
            }

            // 🧱 3️⃣ Lấy số hàng/cột tối đa trong dictionary
            int maxRow = tableMapTongongdong.Keys.Max();
            int maxCol = tableMapTongongdong.Values.Max(dict => dict.Keys.Max());

            //MessageBox.Show($"Max Row: {maxRow}, Max Col: {maxCol}");
            if(themdong)
            {
                for (int c = 0; c < maxCol; c++)
                    header.InsertColumn(c);
                for (int r = 0; r < maxRow; r++)
                    header.InsertRow(r);

            }
          




            foreach (var rowPair in tableMapTongongdong)
            {
                int rowIndex = rowPair.Key;          // 🔹 Hàng hiện tại
                var rowDict = rowPair.Value;         // 🔹 Dữ liệu của hàng đó

                // 🔹 Lấy dữ liệu từng cột
                string col1 = rowDict.ContainsKey(0) ? rowDict[0] : "";
                string col2 = rowDict.ContainsKey(1) ? rowDict[1] : "";
                string col3 = rowDict.ContainsKey(3) ? rowDict[3] : "";

                // 🔹 Đảm bảo schedule có đủ hàng


                // 🔹 Ghi dữ liệu vào đúng dòng
                header.SetCellText(rowIndex, 0, col1);
                header.SetCellText(rowIndex, 1, col2);
                header.SetCellText(rowIndex, 2, col3);
            }

            for (int r = 0; r < tableMapTongongdong.Count; r++)
            {
                foreach (var col in tableMapTongongdong[r])
                {
                    header.SetCellText(r, col.Key, col.Value);
                }
            }

            // 🔹 Duyệt toàn bộ ô trong Header

            header.InsertRow(0);
            // 🟩 Tách chuỗi dựa theo ký tự phân tách "|"
            string[] parts = tieude.Split('|');

            // 🟦 Loại bỏ khoảng trắng thừa (nếu có)
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = parts[i].Trim();
            }

            // 🧩 Đưa từng phần vào các ô trong header
            for (int j = 0; j < parts.Length; j++)
            {
                header.SetCellText(0, j, parts[j]); // 0 là hàng đầu tiên
            }
            try
            {
                // 🔹 8. Thêm field Assembly Code
                ScheduleDefinition definition = sched.Definition;
                SchedulableField assemblyField = null;

                foreach (SchedulableField f in definition.GetSchedulableFields())
                {
                    if (f.GetName(doc) == "Assembly Code")
                    {
                        assemblyField = f;
                        break;
                    }
                }
                // 🔹 Lấy định nghĩa schedule


                if (assemblyField != null)
                {
                    definition.AddField(assemblyField);
                }
                else
                {
                    TaskDialog.Show("⚠️ Cảnh báo", "Không tìm thấy field Assembly Code trong category Generic Model.");
                }

                // 🔹 9. Thêm 2 filter (Filter by Assembly Code)
                if (assemblyField != null)
                {
                    ScheduleField field = definition.GetField(0); // Assembly Code
                    ScheduleFilter filter1 = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, "NO VALUES FOUND");
                    ScheduleFilter filter2 = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, "ALL VALUES FOUND");

                    definition.AddFilter(filter1);
                    definition.AddFilter(filter2);
                }
            }
            catch
            {

            }
            if (xoadongcuoi == true)
            {
                //MessageBox.Show("Xóa dòng cuối");
                int dongcanxoa = header.NumberOfRows - 1;
                header.RemoveRow(dongcanxoa);
            }

        }
    }
}
