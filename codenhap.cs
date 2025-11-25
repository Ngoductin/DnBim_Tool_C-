using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Linq;
using System.Text;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]

    public class CreateExcelStyleSchedule : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            using (Transaction t = new Transaction(doc, "Tạo bảng Excel ảo readable"))
            {
                t.Start();
                // 🔹 1. Tạo schedule trống (Generic Model)
                ElementId catId = new ElementId(BuiltInCategory.OST_GenericModel);
                ViewSchedule sched = ViewSchedule.CreateSchedule(doc, catId);
                // 🔹 Tìm tất cả các schedule hiện có trong project
                var allSchedules = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSchedule))
                    .Cast<ViewSchedule>()
                    .ToList();

                // 🔹 Đếm xem đã có bao nhiêu schedule tên bắt đầu bằng "Table"
                int count = allSchedules.Count(v => v.Name.StartsWith("Table", StringComparison.OrdinalIgnoreCase));

                // 🔹 Tạo tên mới: Table1, Table2, Table3...
                string newName = $"Table{count + 1}";

                // 🔹 Đặt tên cho schedule mới
                sched.Name = newName;


                // 🔹 Lấy dữ liệu bảng
                TableData tableData = sched.GetTableData();

            
                

                

                // 🔹 2. Lấy vùng Header (nơi được phép chèn dòng/cột)
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
                // 🔹 3. Số dòng và cột
                int rows = 6, cols = 4;

                for (int r = 0; r < header.NumberOfRows; r++)
                {
                    for (int c = 0; c < header.NumberOfColumns; c++)
                    {
                        header.ClearCell(r, c);
                    }
                }
                for (int c = 0; c < cols; c++)
                    header.InsertColumn(c);
                for (int r = 0; r < rows; r++)
                    header.InsertRow(r);

                // 🔹 4. Đặt chiều rộng & chiều cao ô để tránh lỗi “too large”
                double colWidth = 30 / 304.8; // 30 mm
                double rowHeight = 8 / 304.8; // 8 mm
                                              // 🔹 Clear toàn bộ ô đang có trong Header
               

               
                for (int r = 0; r < rows+1; r++)
                {
                    header.SetRowHeight(r, rowHeight);
                    for (int c = 0; c < cols+1; c++)
                    {
                        string text = $"   R{r }-C{c }   ";
                        header.SetCellText(r, c, text);
                        header.SetColumnWidth(c, colWidth);
                    }
                }

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

                t.Commit();
            }

            //TaskDialog.Show("✅ Hoàn tất", "Bảng Excel ảo readable đã tạo thành công kèm 2 bộ lọc Assembly Code!");
            return Result.Succeeded;
        }
    }
}
