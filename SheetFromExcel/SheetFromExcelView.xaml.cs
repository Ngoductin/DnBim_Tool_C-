using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;
using System.Data;
using System.IO;
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;

namespace DnBim_Tool
{
    // Class ánh xạ dữ liệu từng dòng từ Excel
    public class SheetData
    {
        public string SheetNumber { get; set; }
        public string SheetName { get; set; }
        public string ViewName { get; set; }
        public string LevelName { get; set; }
        public string ScopeBoxName { get; set; }
        public string ViewTemplateName { get; set; }
        public string TitleOnSheet { get; set; }
        public string SheetGroup { get; set; }
    }

    /// <summary>
    /// Interaction logic for SheetFromExcelView.xaml
    /// </summary>
    public partial class SheetFromExcelView : Window
    {
        private Document doc;
        public SheetFromExcelView(Document doc)
        {
            InitializeComponent();
            this.doc = doc;
            cbb_TitleBlocks.ItemsSource = GetListTitleBlocks();
            cbb_TitleBlocks.SelectedIndex = 0;
        }

        private void bt_Browse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Chọn file Excel";
            dialog.Filter = "Excel Files| *xls; *xlsx; *xlsm";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                tb_FilePath.Text = dialog.FileName;
            }
            else tb_FilePath.Text = "";
        }

        private List<string> GetListTitleBlocks()
        {
            var collector = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsElementType()
                        .Cast<FamilySymbol>()
                        .ToList();

            var listNames = (collector.Select(x => $"{x.FamilyName}: {x.Name}")).ToList();
            listNames.Sort();
            return listNames;
        }

        private ElementId GetBlockId(string blockName)
        {
            var collector = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsElementType()
                        .Cast<FamilySymbol>()
                        .ToList();
            var id = collector.Find(x => $"{x.FamilyName}: {x.Name}" == blockName)?.Id;
            return id;
        }

        private void bt_Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // Đọc dữ liệu Excel thành List<SheetData>
        private List<SheetData> ReadExcelData(string filePath)
        {
            var result = new List<SheetData>();
            using (FileStream fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(fileStream))
                {
                    var data = reader.AsDataSet();
                    if (data == null || data.Tables.Count == 0) return result;
                    var table = data.Tables[0];
                    if (table.Rows.Count < 2) return result;

                    // Đọc header
                    var headers = new List<string>();
                    foreach (DataColumn col in table.Columns)
                        headers.Add(col.ColumnName);

                    // Đọc từng dòng dữ liệu
                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];
                        var sd = new SheetData
                        {
                            SheetNumber = row[0]?.ToString(),
                            SheetName = row[1]?.ToString(),
                            ViewName = row[2]?.ToString(),
                            LevelName = row[3]?.ToString(),
                            ScopeBoxName = row[4]?.ToString(),
                            ViewTemplateName = row[5]?.ToString(),
                            TitleOnSheet = row[6]?.ToString(),
                            SheetGroup = row[7]?.ToString()
                        };
                        // Bỏ qua dòng trống
                        if (string.IsNullOrWhiteSpace(sd.SheetNumber) && string.IsNullOrWhiteSpace(sd.SheetName) && string.IsNullOrWhiteSpace(sd.ViewName))
                            continue;
                        result.Add(sd);
                    }
                }
            }
            return result;
        }

        private void bt_Ok_Click(object sender, RoutedEventArgs e)
        {
            // 1. Chọn Title Block
            string blockName = cbb_TitleBlocks.SelectedValue?.ToString();
            ElementId blockId = GetBlockId(blockName);

            // 2. Kiểm tra đường dẫn file Excel
            string filePath = tb_FilePath.Text;
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Chọn file excel hoặc copy/paste đường dẫn!", "Thông báo");
                return;
            }
            if (filePath.Contains("\"")) filePath = filePath.Replace("\"", "");

            // 3. Đọc dữ liệu từ file Excel
            var excelData = ReadExcelData(filePath);
            if (excelData == null || excelData.Count == 0)
            {
                MessageBox.Show("File Excel không có dữ liệu hợp lệ!", "Thông báo");
                return;
            }

            // 4. Đóng file Excel nếu đang mở
            CloseExcelFile(filePath);

            // 5. Kiểm tra sheet number trùng và thông báo
            List<string> duplicateSheets = new List<string>();
            List<SheetData> validSheets = new List<SheetData>();
            foreach (var row in excelData)
            {
                if (!string.IsNullOrEmpty(row.SheetNumber) && IsSheetNumberExists(row.SheetNumber))
                {
                    duplicateSheets.Add(row.SheetNumber);
                }
                else
                {
                    validSheets.Add(row);
                }
            }
            if (duplicateSheets.Count > 0)
            {
                string msg = "Các Sheet Number sau đã tồn tại và được bỏ qua:\n" + string.Join("\n", duplicateSheets);
                MessageBox.Show(msg, "Sheet Number trùng lặp", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            // Loại bỏ sheet trùng trong file Excel
            validSheets = validSheets
                .GroupBy(s => s.SheetNumber, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();

            List<(Viewport viewport, SheetData data)> viewportList = new List<(Viewport, SheetData)>();

            // Sử dụng TransactionGroup để rollback nếu có lỗi
            using (TransactionGroup tg = new TransactionGroup(doc, "Tạo Sheet, View, Viewport và gán Scope Box"))
            {
                tg.Start();
                try
                {
                    using (var t = new Transaction(doc, "Tạo Sheet, View, Viewport và gán Scope Box"))
                    {
                        t.Start();
                        foreach (var row in validSheets)
                        {
                            ViewSheet vs = ViewSheet.Create(doc, blockId);
                            if (vs == null) continue;

                            // Gán các thuộc tính cơ bản
                            if (!string.IsNullOrEmpty(row.SheetNumber))
                            {
                                try
                                {
                                    vs.SheetNumber = row.SheetNumber;
                                }
                                catch (Autodesk.Revit.Exceptions.ArgumentException)
                                {
                                    throw new Exception($"Sheet Number đã tồn tại: {row.SheetNumber}");
                                }
                            }
                            if (!string.IsNullOrEmpty(row.SheetName))
                                vs.Name = row.SheetName;

                            // Gán các parameter khác nếu có
                            if (!string.IsNullOrEmpty(row.SheetGroup))
                                SetParameters(vs, "Sheet Group", row.SheetGroup);
                            if (!string.IsNullOrEmpty(row.TitleOnSheet))
                                SetParameters(vs, "Title on Sheet", row.TitleOnSheet);

                            // Tạo view nếu có ViewName, Level, ScopeBox, ViewTemplate
                            Autodesk.Revit.DB.View view = null;
                            bool needDraftingView = string.IsNullOrEmpty(row.LevelName) && string.IsNullOrEmpty(row.ScopeBoxName) && string.IsNullOrEmpty(row.ViewTemplateName);

                            if (!string.IsNullOrEmpty(row.ViewName) && !needDraftingView)
                            {
                                view = CreateOrGetView(row.ViewName, row.LevelName);

                                // Gán View Template nếu có
                                if (!string.IsNullOrEmpty(row.ViewTemplateName) && view != null)
                                {
                                    var template = new FilteredElementCollector(doc)
                                        .OfClass(typeof(Autodesk.Revit.DB.View))
                                        .Cast<Autodesk.Revit.DB.View>()
                                        .FirstOrDefault(vt => vt.IsTemplate && vt.Name == row.ViewTemplateName);
                                    if (template != null)
                                        view.ViewTemplateId = template.Id;
                                }
                            }
                            else
                            {
                                // Tạo Drafting View nếu thiếu các giá trị
                                view = CreateDraftingView(doc, row.ViewName ?? "Drafting View Auto");
                            }

                            // Tạo viewport ở vị trí tạm thời (0,0,0)
                            Viewport viewport = null;
                            if (vs != null && view != null && Viewport.CanAddViewToSheet(doc, vs.Id, view.Id))
                            {
                                viewport = Viewport.Create(doc, vs.Id, view.Id, new XYZ(0, 0, 0));

                                // Gán Scope Box cho viewport nếu có
                                if (!string.IsNullOrEmpty(row.ScopeBoxName))
                                {
                                    var scopeBox = new FilteredElementCollector(doc)
                                        .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                                        .WhereElementIsNotElementType()
                                        .FirstOrDefault(se => se.Name.Equals(row.ScopeBoxName, StringComparison.InvariantCultureIgnoreCase));

                                    if (scopeBox != null)
                                    {
                                        var param = viewport.LookupParameter("Scope Box");
                                        if (param != null && !param.IsReadOnly)
                                        {
                                            param.Set(scopeBox.Id);
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception($"Không tìm thấy Scope Box: {row.ScopeBoxName}");
                                    }
                                }

                                // Gán Title on Sheet nếu có
                                if (!string.IsNullOrEmpty(row.TitleOnSheet))
                                {
                                    var titleParam = viewport.LookupParameter("Title on Sheet");
                                    if (titleParam != null)
                                    {
                                        titleParam.Set(row.TitleOnSheet);
                                    }
                                }

                                viewportList.Add((viewport, row));
                            }
                        }
                        t.Commit();
                    }

                    // 6. Người dùng chọn 2 điểm trên sheet
                    UIDocument uidoc = new UIDocument(doc);
                    var points = PickRectangle(uidoc);
                    if (points == null)
                    {
                        throw new Exception("Bạn đã hủy thao tác chọn điểm!");
                    }

                    XYZ pt1 = points.Item1;
                    XYZ pt2 = points.Item2;
                    double z = pt1.Z;
                    XYZ center = new XYZ(
                        (pt1.X + pt2.X) / 2,
                        (pt1.Y + pt2.Y) / 2,
                        z);

                    // 7. Di chuyển từng viewport vào giữa 2 điểm đã chọn
                    int movedCount = 0;
                    using (var t = new Transaction(doc, "Di chuyển Viewport vào giữa 2 điểm"))
                    {
                        t.Start();
                        foreach (var item in viewportList)
                        {
                            var viewport = item.viewport;
                            if (viewport != null)
                            {
                                var boxCenter = viewport.GetBoxCenter();
                                var translation = center.Subtract(boxCenter);
                                ElementTransformUtils.MoveElement(doc, viewport.Id, translation);
                                movedCount++;
                            }
                        }
                        t.Commit();
                    }

                    tg.Assimilate(); // Commit tất cả nếu thành công
                    MessageBox.Show($"{viewportList.Count} sheet đã được tạo!\n{movedCount} viewport đã được di chuyển vào giữa 2 điểm chọn.", "Kết quả");
                    Close();
                }
                catch (Exception ex)
                {
                    tg.RollBack(); // Hoàn tác toàn bộ thay đổi
                    MessageBox.Show($"Có lỗi xảy ra: {ex.Message}\nMọi thay đổi đã được hủy.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        // Kiểm tra sheet number đã tồn tại chưa
        private bool IsSheetNumberExists(string sheetNumber)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Any(s => s.SheetNumber.Equals(sheetNumber, StringComparison.OrdinalIgnoreCase));
        }

        private void CloseExcelFile(string filePath)
        {
            Excel.Application app = new Excel.Application();
            if (app != null)
            {
                foreach (Workbook workbook in app.Workbooks)
                {
                    if (workbook.FullName == filePath)
                    {
                        workbook.Close(false);
                        break;
                    }
                }
                Marshal.ReleaseComObject(app);
            }
        }

        // Tạo Floor Plan hoặc Drafting View nếu chưa có
        private Autodesk.Revit.DB.View CreateOrGetView(string viewName, string levelName)
        {
            // Kiểm tra view đã tồn tại chưa
            var view = new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .FirstOrDefault(v => v.Name == viewName && !v.IsTemplate);

            if (view != null)
                return view;

            // Nếu có levelName thì tạo Floor Plan, không thì tạo Drafting View
            if (!string.IsNullOrEmpty(levelName))
            {
                // Tìm Level
                var level = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .FirstOrDefault(l => l.Name == levelName);
                if (level == null) return null;

                // Tìm ViewFamilyType cho FloorPlan
                var vft = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(x => x.ViewFamily == ViewFamily.FloorPlan);
                if (vft == null) return null;

                var newView = ViewPlan.Create(doc, vft.Id, level.Id);
                newView.Name = viewName;
                return newView;
            }
            else
            {
                // Tìm ViewFamilyType cho Drafting
                var vft = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(x => x.ViewFamily == ViewFamily.Drafting);
                if (vft == null) return null;

                var newView = ViewDrafting.Create(doc, vft.Id);
                newView.Name = viewName;
                return newView;
            }
        }

        // Tạo Drafting View nếu không có Level, Scope Box, View Template
        private Autodesk.Revit.DB.View CreateDraftingView(Document doc, string viewName)
        {
            var vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(x => x.ViewFamily == ViewFamily.Drafting);

            if (vft == null) return null;

            var draftingView = ViewDrafting.Create(doc, vft.Id);
            draftingView.Name = viewName;

            // Thêm TextNote "DRAFT"
            var textType = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .FirstElement() as TextNoteType;

            if (textType != null)
            {
                TextNote.Create(doc, draftingView.Id, new XYZ(0, 0, 0), "DRAFT", textType.Id);
            }

            // Thêm một đường line dài 1000mm
            XYZ startPoint = new XYZ(-0.5, 0, 0); // -500mm
            XYZ endPoint = new XYZ(0.5, 0, 0);    // +500mm
            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);
            doc.Create.NewDetailCurve(draftingView, line);

            return draftingView;
        }

        private void SetParameters(ViewSheet vs, string parameterName, string parameterValue)
        {
            try
            {
                Autodesk.Revit.DB.Parameter p = vs.LookupParameter(parameterName);
                if (p != null && !p.IsReadOnly)
                {
                    var paraType = p.StorageType;
                    if (paraType == StorageType.Integer)
                    {
                        p.Set(int.Parse(parameterValue));
                    }
                    else if (paraType == StorageType.Double)
                    {
                        p.Set(double.Parse(parameterValue));
                    }
                    else if (paraType == StorageType.String)
                    {
                        p.Set(parameterValue);
                    }
                }
            }
            catch { }
        }

        public void GanScopeBoxVaoView(UIDocument uidoc, Document doc, string tenScopeBox, Autodesk.Revit.DB.View view)
        {
            var scopeBox = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                .WhereElementIsNotElementType()
                .FirstOrDefault(e => e.Name == tenScopeBox);

            if (scopeBox == null)
            {
                TaskDialog.Show("Thông báo", $"Không tìm thấy Scope Box tên: {tenScopeBox}");
                return;
            }

            using (Transaction t = new Transaction(doc, "Gán Scope Box vào View"))
            {
                t.Start();
                var param = view.LookupParameter("Scope Box");
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(scopeBox.Id);
                }
                t.Commit();
            }

            TaskDialog.Show("Thông báo", $"Đã gán Scope Box '{tenScopeBox}' vào View.");
        }

        public void GanScopeBoxVaoView(UIDocument uidoc, Document doc, string tenScopeBox)
        {
            Autodesk.Revit.DB.View activeView = doc.ActiveView;

            var scopeBox = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                .WhereElementIsNotElementType()
                .FirstOrDefault(e => e.Name == tenScopeBox);

            if (scopeBox == null)
            {
                TaskDialog.Show("Thông báo", $"Không tìm thấy Scope Box tên: {tenScopeBox}");
                return;
            }

            using (Transaction t = new Transaction(doc, "Gán Scope Box vào View"))
            {
                t.Start();
                var param = activeView.LookupParameter("Scope Box");
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(scopeBox.Id);
                }
                t.Commit();
            }

            TaskDialog.Show("Thông báo", $"Đã gán Scope Box '{tenScopeBox}' vào View.");
        }

        /// <summary>
        /// Chọn 2 điểm trên Sheet View để xác định vùng hình chữ nhật
        /// </summary>
        private Tuple<XYZ, XYZ> PickRectangle(UIDocument uidoc)
        {
            try
            {
                XYZ pt1 = uidoc.Selection.PickPoint(ObjectSnapTypes.None, "Chọn điểm đầu của hình chữ nhật (theo Title Block đã chọn)");
                XYZ pt2 = uidoc.Selection.PickPoint(ObjectSnapTypes.None, "Chọn điểm đối diện của hình chữ nhật");
                return new Tuple<XYZ, XYZ>(pt1, pt2);
            }
            catch
            {
                return null;
            }
        }

        private void cbb_TitleBlocks_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }
    }
}