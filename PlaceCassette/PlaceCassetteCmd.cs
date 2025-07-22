using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Dnbim_Tool.Placecassette;
using DnBim_Tool;
using ExcelDataReader;
using static System.Windows.Forms.LinkLabel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;
using Line = Autodesk.Revit.DB.Line;
using Transform = Autodesk.Revit.DB.Transform;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]


    public class PlaceCassetteCmd : IExternalCommand
    
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;
            AddAllSharedParametersToRoom(doc);



          
            var window = new PlaceCassetteView();


            window.ShowDialog();
            string filepath = window.tb_FilePath.Text.ToString();
            double khoangcachtuong = 0;
           

            string sheetNamesisSelected = window.cbbSheetName.SelectedValue.ToString();

            if (window.DialogResult == true)
            {
                if (string.IsNullOrEmpty(filepath))
                {
                    MessageBox.Show("Chọn file excel hoặc copy/paste đường dẫn!", "Message");
                }
                else
                {
                    if (window.cbbOption.SelectedValue.ToString() == "Kinh Tế")
                    {
                        khoangcachtuong = 7000 / 304.8;
                    }
                    else
                    {
                        khoangcachtuong = 4000 / 304.8;
                    }
                    //MessageBox.Show((khoangcachtuong*304.8).ToString(), "Message");
                    if (filepath.Contains("\"")) filepath = filepath.Replace("\"", "");

                    System.Data.DataTable excelData = new System.Data.DataTable();
                    using (FileStream fileStream = File.Open(filepath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(fileStream))
                        {
                            IDictionary sortedCapacityModelTable= new Hashtable();
                            var data = reader.AsDataSet();
                            if (data != null)
                            {
                                IList<string> AllcolumNames = new List<string>();
                                IList<string> AllValue = new List<string>();

                                for (int i = 0; i < data.Tables.Count; i++)
                                {
                                    excelData = data.Tables[i];
                                    if (excelData.TableName == "CASSETTE")
                                    {
                                        using (Transaction t = new Transaction(doc, "CASSETTE"))
                                        {
                                            t.Start();
                                            sortedCapacityModelTable = ReadExcelData(excelData);
                                           
                                            /* Xuất toàn bộ Hashtable trong một lần*/
                                            //DisplayHashtable(sortedCapacityModelTable);
                                            t.Commit();
                                        }
                                    }
                                }
                                for (int i = 0; i < data.Tables.Count; i++)
                                {
                                    excelData = data.Tables[i];
                                    if (excelData.ToString() == sheetNamesisSelected)
                                    {
                                        //CreateSheetFromExcel(doc, excelData, blockId, pick1, pick2);
                                        using (Transaction t = new Transaction(doc, " "))
                                        {

                                            t.Start();
                                            try
                                            {
                                                SetParamterRoomNames(doc, excelData, view, khoangcachtuong , sortedCapacityModelTable);
                                            }
                                            catch { }
                                            t.Commit();
                                        }


                                    }
                                }


                            }
                        }
                    }

                    // Lấy tất cả FamilyInstance trong dự án
                    FilteredElementCollector collector = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance));

                    double totalPrice = 0;

                    foreach (FamilyInstance fi in collector)
                    {
                        // Kiểm tra loại FamilyInstance có phải là Mechanical Equipment không
                        if (fi.Symbol.Family.FamilyCategory.Name == "Mechanical Equipment")
                        {
                            // Kiểm tra tên FamilyInstance bắt đầu bằng 'S' và kết thúc bằng 'N'
                            string familyName = fi.Symbol.Family.Name;
                            if (familyName.StartsWith("S") && familyName.EndsWith("N"))
                            {
                                // Lấy Parameter "PRICE" của FamilyInstance
                                Parameter priceParam = fi.LookupParameter("PRICE");

                                // Kiểm tra xem Parameter PRICE có tồn tại và có giá trị không
                                if (priceParam != null && priceParam.HasValue)
                                {
                                    double priceValue = priceParam.AsDouble(); // Lấy giá trị PRICE (giả sử là double)
                                    totalPrice += priceValue;  // Cộng dồn vào tổng
                                }
                            }
                        }//a
                    }

                    // Hiển thị tổng giá trị PRICE
                    TaskDialog.Show("Dự toán", $"Tổng chi phí thiết bị của phương án này là: {totalPrice}");
                //a
            }

            }
            return Result.Succeeded;
        }

        public void AddAllSharedParametersToRoom(Document doc)
        {
            using (Transaction trans = new Transaction(doc, "Thêm tất cả Shared Parameters vào Room"))
            {
                trans.Start();

                // Lấy file Shared Parameter hiện tại
                DefinitionFile defFile = doc.Application.OpenSharedParameterFile();
                if (defFile == null)
                {
                    TaskDialog.Show("Lỗi", "Không tìm thấy Shared Parameters file. Hãy kiểm tra lại.");
                    return;
                }

                // Lấy BindingMap để kiểm tra và thêm Parameter vào Room
                BindingMap bindingMap = doc.ParameterBindings;

                // Duyệt qua tất cả các nhóm trong file Shared Parameters
                foreach (DefinitionGroup group in defFile.Groups)
                {
                    foreach (Definition definition in group.Definitions)
                    {
                        string parameterName = definition.Name;

                        // Kiểm tra xem Parameter đã được gán vào Room chưa
                        bool isParameterBound = false;
                        DefinitionBindingMapIterator iterator = bindingMap.ForwardIterator();
                        while (iterator.MoveNext())
                        {
                            if (iterator.Key.Name == parameterName)
                            {
                                isParameterBound = true;
                                break;
                            }
                        }

                        if (!isParameterBound)
                        {
                            // Nếu chưa có, thêm Parameter vào Room
                            Category roomCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rooms);
                            CategorySet categorySet = doc.Application.Create.NewCategorySet();
                            categorySet.Insert(roomCategory);

                            // Gán Parameter vào Group "Other" (PG_OTHER)
                            InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categorySet);
                            bindingMap.Insert(definition, binding, BuiltInParameterGroup.PG_TEXT);
                        }
                    }
                }

                trans.Commit();
            }

            //TaskDialog.Show("Thông báo", "Tất cả Shared Parameters đã được thêm vào Room.");
        }
        public void SetParamterRoomNames(Document doc, System.Data.DataTable excelData, View view,double khoangcach, IDictionary sortedCapacityModelTable)
        {
            string roomInfo = "";
            IList<FamilyInstance> cassettes= new List<FamilyInstance>();
            double tongtien = 0;
            // Lấy tất cả các Room trong dự án
            FilteredElementCollector roomCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            foreach (Element room in roomCollector)
            {
                double coolingcapacity = 0;
            
                Parameter nameParam = room.LookupParameter("Name"); // Lấy giá trị Parameter "Name"

                string nameParameter = nameParam.ToString();
                string valueName=nameParam.AsString();
                using (Transaction trans = new Transaction(doc, "Create View Sheet"))
                {
                    int rowCount = excelData.Rows.Count;
                    int colCount = excelData.Columns.Count;
                    int colNameroom = 1;
                    for (int r = 1; r < rowCount; r++)
                    {

                        string header = excelData.Columns[colNameroom].ColumnName;
                        string parameterName = excelData.Rows[0][header].ToString();
                        string parameterValue = excelData.Rows[r][header].ToString();
                        
                        if (parameterValue == valueName)
                        {
                            for (int j = 0; j < colCount; j++)
                            {
                                
                                string tencot = excelData.Columns[j].ColumnName;
                                string tenPara = excelData.Rows[0][tencot].ToString();
                                string giatriPara = excelData.Rows[r][tencot].ToString();
                                //Parameter nameParam1 = room.LookupParameter(tenPara);
                                //nameParam1.Set(giatriPara);
                                SetParameters(room, tenPara, giatriPara);
                            }
                        }
                    }
                    Autodesk.Revit.DB.Parameter p1 = room.LookupParameter("Area");
                    Autodesk.Revit.DB.Parameter p2 = room.LookupParameter("HeatLoad");
                    if (p1 != null && p1.HasValue&& p2 != null && p2.HasValue)
                    {
                        // Lấy giá trị của parameter (đơn vị là feet²)
                        double areaInFeet = p1.AsDouble();
                        double HeatLoad = p2.AsDouble();

                        // Chuyển đổi từ feet² sang m² (1 feet² = 0.092903 m²)
                        double areaInSquareMeters = areaInFeet * 0.092903;

                        /* In ra kết quả*/
                        //TaskDialog.Show("Area", "Area in square meters: " + areaInSquareMeters);
                       
                         coolingcapacity= areaInSquareMeters*HeatLoad/1000;
                        //MessageBox.Show(coolingcapacity.ToString());
                        Autodesk.Revit.DB.Parameter p3 = room.LookupParameter("Cooling capacity");
                        p3.Set(coolingcapacity);  
                    }
                    else
                    {
                        //TaskDialog.Show("Error", "Parameter 'Area' is not found or has no value.");
                    }
                    //Autodesk.Revit.DB.Parameter p1 = room.LookupParameter("Cooling capacity");

                }

                if(room is Room room1)
                {
                    GetMidLineRooom(doc, room1, coolingcapacity, view, khoangcach, sortedCapacityModelTable, cassettes);
                }
               
            }
         

        }

        public void ProcessFixedSheet(Document doc, System.Data.DataTable excelData, View view, Parameter nameParam)
        {
            string nameParameter = nameParam.ToString();
            string valueName = nameParam.AsString();
            using (Transaction trans = new Transaction(doc, "Create View Sheet"))
            {
                int rowCount = excelData.Rows.Count;
                int colCount = excelData.Columns.Count;
                int colNameroom = 1;
                for (int r = 1; r < rowCount; r++)
                {

                    string header = excelData.Columns[colNameroom].ColumnName;
                    string parameterName = excelData.Rows[0][header].ToString();
                    string parameterValue = excelData.Rows[r][header].ToString();

                    if (parameterValue == valueName)
                    {
                        for (int j = 0; j < colCount; j++)
                        {

                            string tencot = excelData.Columns[j].ColumnName;
                            string tenPara = excelData.Rows[0][tencot].ToString();
                            string giatriPara = excelData.Rows[r][tencot].ToString();
                            

                        }
                    }
                }
            } 
            }

        public static void SetParameters(Element room, string parameterName, string parameterValue)
        {
            try
            {
                Autodesk.Revit.DB.Parameter p = room.LookupParameter(parameterName);

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
        public static void GetMidLineRooom(Document doc, Room room
                                         , double coolingcapacity,
                                        View view, double khoangcach
                                        , IDictionary sortedCapacityModelTable,
                                        IList<FamilyInstance> cassettes
                                        
                               )
        {

            // Lấy Transform của View hiện tại (nếu cần chuyển đổi hệ tọa độ từ View)
            Transform viewTransform = view.CropBox.Transform;
            if (room == null)
            {
                TaskDialog.Show("Error", "No room selected.");
                return;
            }
            Level roomLevel = room.Level;
            string levelName = roomLevel.Name;

            // Tạo tùy chọn lấy biên dạng
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();

            // Lấy danh sách đường biên của Room
            IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);

            if (boundaries.Count == 0)
            {
                //TaskDialog.Show("Info", "No boundary segments found.");
                return;
            }

            IList<Line> Chieudai = new List<Line>();

            // Duyệt qua các đường biên và lấy thông tin
            foreach (IList<BoundarySegment> boundaryList in boundaries)
            {
                //MessageBox.Show(boundaryList.Count.ToString());





                IList<XYZ> points = new List<XYZ>();
                IList<Line> boundlines = new List<Line>();
                foreach (IList<BoundarySegment> boundaryList1 in boundaries)
                {
                    foreach (BoundarySegment segment in boundaryList1)
                    {
                        Curve curve = segment.GetCurve();

                        if (curve is Line line)
                        {
                            //DetailLine detailLine = (DetailLine)doc.Create.NewDetailCurve(view, line);
                            boundlines.Add(line);
                            for (int i = 0; i < 2; i++)
                            {
                                points.Add(line.GetEndPoint(i));
                            }
                        }
                    }
                }
                List<XYZ> filteredPoints = new List<XYZ>();

                foreach (XYZ point in points)
                {
                    bool isDuplicate = false;

                    // Kiểm tra với danh sách đã lọc trước đó
                    foreach (XYZ uniquePoint in filteredPoints)
                    {
                        if (point.DistanceTo(uniquePoint) < 0.0001)
                        {
                            isDuplicate = true;
                            break;
                        }
                    }

                    // Nếu không trùng thì thêm vào danh sách kết quả
                    if (!isDuplicate)
                    {
                        filteredPoints.Add(point);
                    }
                }
                double xSum = points.Sum(p => p.X);
                double ySum = points.Sum(p => p.Y);
                double zSum = points.Sum(p => p.Z);

                int count = points.Count;

                XYZ trongtam = new XYZ(xSum / count, ySum / count, zSum / count);

                Line Y = Line.CreateBound(trongtam + new XYZ(0, 500, 0), (trongtam + new XYZ(0, -500, 0)));
                Line X = Line.CreateBound(trongtam + new XYZ(500, 0, 0), (trongtam + new XYZ(-500, 0, 0)));
                //DetailLine detailLine = (DetailLine)doc.Create.NewDetailCurve(view, X);
                //DetailLine detailLine1 = (DetailLine)doc.Create.NewDetailCurve(view, Y);
                Plane plane1 = Plane.CreateByNormalAndOrigin(XYZ.BasisX, trongtam);
                Plane plane2 = Plane.CreateByNormalAndOrigin(XYZ.BasisY, trongtam);

                IList<XYZ> newpointX = new List<XYZ>();

                //MessageBox.Show(boundlines.Count.ToString());
                foreach (Line line in boundlines)
                {
                    IntersectionResultArray results1;
                    SetComparisonResult result1 = X.Intersect(line, out results1);
                    if (result1 == SetComparisonResult.Overlap && results1 != null && results1.Size > 0)
                    {
                        XYZ intersectionPoint = results1.get_Item(0).XYZPoint;
                        newpointX.Add(intersectionPoint);

                    }

                }

                IList<XYZ> newpointY = new List<XYZ>();
                foreach (Line line in boundlines)
                {
                    IntersectionResultArray results1;
                    SetComparisonResult result1 = Y.Intersect(line, out results1);
                    if (result1 == SetComparisonResult.Overlap && results1 != null && results1.Size > 0)
                    {
                        XYZ intersectionPoint = results1.get_Item(0).XYZPoint;
                        newpointY.Add(intersectionPoint);

                    }

                }
                /*TrungdiemX*/
                XYZ trungdiemX = new XYZ(newpointX.Sum(p => p.X) / newpointX.Count, newpointX.Sum(p => p.Y) / newpointX.Count, newpointX.Sum(p => p.Z) / newpointX.Count);
                /*TrungdiemY*/
                XYZ trungdiemY = new XYZ(newpointY.Sum(p => p.X) / newpointY.Count, newpointY.Sum(p => p.Y) / newpointY.Count, newpointY.Sum(p => p.Z) / newpointY.Count);

                XYZ diemgiuaphong = new XYZ(trungdiemX.X, trungdiemY.Y, trungdiemX.Z);


                IList<XYZ> diemtrucX = new List<XYZ>();
                IList<XYZ> diemtrucY = new List<XYZ>();

                Line X1 = Line.CreateBound(diemgiuaphong + new XYZ(500, 0, 0), (diemgiuaphong + new XYZ(-500, 0, 0)));
                Line Y1 = Line.CreateBound(diemgiuaphong + new XYZ(0, 500, 0), (diemgiuaphong + new XYZ(0, -500, 0)));

                foreach (Line line in boundlines)
                {
                    IntersectionResultArray results1;
                    SetComparisonResult result1 = X1.Intersect(line, out results1);
                    if (result1 == SetComparisonResult.Overlap && results1 != null && results1.Size > 0)
                    {
                        XYZ intersectionPoint = results1.get_Item(0).XYZPoint;
                        diemtrucX.Add(intersectionPoint);

                    }

                }
                foreach (Line line in boundlines)
                {
                    IntersectionResultArray results1;
                    SetComparisonResult result1 = Y1.Intersect(line, out results1);
                    if (result1 == SetComparisonResult.Overlap && results1 != null && results1.Size > 0)
                    {
                        XYZ intersectionPoint = results1.get_Item(0).XYZPoint;
                        diemtrucY.Add(intersectionPoint);

                    }
                    //a
                }
                Line trucX = Line.CreateBound(diemtrucX[0], diemtrucX[1]);
                Line trucY = Line.CreateBound(diemtrucY[0], diemtrucY[1]);
                //DetailLine detailLine = (DetailLine)doc.Create.NewDetailCurve(view, trucX);
                //DetailLine detailLine1 = (DetailLine)doc.Create.NewDetailCurve(view, trucY);

                IList<XYZ> Diemtruc = new List<XYZ>();



                if (trucY.Length > trucX.Length)
                {



                    int sodiemtrentruc = (int)(trucY.Length / khoangcach);

                    double khoangcachtuong = (trucY.Length - khoangcach * sodiemtrentruc) / 2;
                    if (sodiemtrentruc == 0)
                    {

                        Diemtruc.Add(diemgiuaphong);
                    }
                    if (khoangcachtuong < (1500 / 304.8) && sodiemtrentruc == 1)
                    {

                        Diemtruc.Add(diemgiuaphong);
                    }

                    else
                    {

                        if (sodiemtrentruc % 2 == 0)
                        {
                            sodiemtrentruc = sodiemtrentruc - 1;
                        }
                        //MessageBox.Show(sodiemtrentruc.ToString());
                        for (int i = 1; i <= sodiemtrentruc; i++)
                        {
                            if ((trucX.Length - ((2 * i - 1) * khoangcach)) / 2 > 1500 / 304.8)
                            {  // Nếu chưa đạt điều kiện, tiếp tục thêm điểm vào danh sách
                                Diemtruc.Add(CT.FindTwoPointsOnLineFromPoint(diemgiuaphong, trucY, (2 * i - 1) * khoangcach).Item1);
                                Diemtruc.Add(CT.FindTwoPointsOnLineFromPoint(diemgiuaphong, trucY, (2 * i - 1) * khoangcach).Item2);
                            }
                        }
                           


                    }
                }
                if (trucY.Length < trucX.Length)
                {
                    int sodiemtrentruc = (int)(trucX.Length / khoangcach);

                    double khoangcachtuong = (trucX.Length - khoangcach * sodiemtrentruc) / 2;

                    if (sodiemtrentruc == 0)
                    {
                        //MessageBox.Show(sodiemtrentruc.ToString());
                        Diemtruc.Add(diemgiuaphong);
                    }
                    if (khoangcachtuong < (1500 / 304.8) && sodiemtrentruc == 1)
                    {
                        Diemtruc.Add(diemgiuaphong);
                    }
                    else
                    {
                        if (sodiemtrentruc % 2 == 0)
                        {
                            sodiemtrentruc = sodiemtrentruc - 1;
                        }
                        //MessageBox.Show(sodiemtrentruc.ToString());

                        for (int i = 1; i <= sodiemtrentruc; i++)

                        {  // Nếu chưa đạt điều kiện, tiếp tục thêm điểm vào danh sách

                            //MessageBox.Show(((trucX.Length - ((2 * i - 1) * khoangcach)) > 1500 / 304.8).ToString());
                            if ((trucX.Length - ((2 * i - 1) * khoangcach))/2 > 1500 / 304.8)
                            {
                               
                                Diemtruc.Add(CT.FindTwoPointsOnLineFromPoint(diemgiuaphong, trucX, (2 * i - 1) * khoangcach).Item1);
                                Diemtruc.Add(CT.FindTwoPointsOnLineFromPoint(diemgiuaphong, trucX, (2 * i - 1) * khoangcach).Item2);
                            }

                        }

                    }
                }




                //MessageBox.Show(Diemtruc.Count.ToString());

                IList<XYZ> Diemdat = new List<XYZ>();

                foreach (XYZ p in Diemtruc)
                {
                    Line line = null;
                    if (trucY.Length > trucX.Length)
                    {
                        line = Line.CreateBound(p + new XYZ(500, 0, 0),
                           (p + new XYZ(-500, 0, 0)));
                    }
                    else
                    {
                        line = Line.CreateBound(p + new XYZ(0, 500, 0),
                            (p + new XYZ(0, -500, 0)));
                    }

                    IList<XYZ> exp = new List<XYZ>();//Điểm sau khi mở rộng giao với biên
                    foreach (Line linebound in boundlines)
                    {
                        IntersectionResultArray results1;
                        SetComparisonResult result1 = line.Intersect(linebound, out results1);
                        if (result1 == SetComparisonResult.Overlap && results1 != null && results1.Size > 0)
                        {
                            XYZ intersectionPoint = results1.get_Item(0).XYZPoint;
                            exp.Add(intersectionPoint);

                        }

                    }
                    Line newline = Line.CreateBound(exp[0], exp[1]);
                   
                    //DetailLine detailLine = (DetailLine)doc.Create.NewDetailCurve(view, newline);
                    int sodiemtrentruc = (int)(newline.Length / khoangcach);

                    double khoangcachtuong = (newline.Length - khoangcach * sodiemtrentruc) / 2;
                    //MessageBox.Show(((newline.Length- khoangcach)/2 * 304.8).ToString());
                    if (sodiemtrentruc == 0)
                    {
                        //MessageBox.Show(sodiemtrentruc.ToString());
                        Diemdat.Add(p);
                    }
                    if (khoangcachtuong < (1500 / 304.8) && sodiemtrentruc == 1)
                    {
                        //MessageBox.Show("toi bi ngu");
                        Diemdat.Add(p);
                    }
                    else
                    {

                        if (sodiemtrentruc % 2 == 0)
                        {
                            sodiemtrentruc = sodiemtrentruc - 1;
                        }
                        for (int i = 1; i <= sodiemtrentruc; i++)
                            if ((newline.Length - ((2 * i - 1) * khoangcach)) / 2 > 1500 / 304.8)
                            {  // Nếu chưa đạt điều kiện, tiếp tục thêm điểm vào danh sách
                                Diemdat.Add(CT.FindTwoPointsOnLineFromPoint(p, newline, (2 * i - 1) * khoangcach).Item1);
                                Diemdat.Add(CT.FindTwoPointsOnLineFromPoint(p, newline, (2 * i - 1) * khoangcach).Item2);
                            }

                    }
                }

                double congsuatmoimay = coolingcapacity / Diemdat.Count;
                //MessageBox.Show(congsuatmoimay.ToString());

                // Khởi tạo biến để lưu kết quả
                Dictionary<string, string> smallestLargerEntryValue = null;
                double smallestLargerEntryKey = double.NaN;

                // Duyệt qua từng phần tử trong IDictionary đã sắp xếp
                foreach (DictionaryEntry entry in sortedCapacityModelTable)
                {
                    // Ép kiểu entry.Key và entry.Value về đúng kiểu
                    double key = (double)entry.Key;
                    Dictionary<string, string> value = (Dictionary<string, string>)entry.Value;
                    //aBB
                    // Kiểm tra nếu key của entry lớn hơn target
                    if (key > congsuatmoimay) // là target của bạn
                    {
                        // Nếu chưa tìm thấy giá trị lớn hơn, hoặc nếu giá trị tìm thấy nhỏ hơn giá trị trước đó, cập nhật kết quả
                        smallestLargerEntryKey = key;
                        smallestLargerEntryValue = value;
                        break; // Sau khi tìm thấy giá trị lớn nhất nhỏ hơn target, thoát khỏi vòng lặp
                    }
                }
                string tenmay = null;
                string giatien=null;    
                // Kiểm tra kết quả
                if (smallestLargerEntryValue != null)
                {
                    //MessageBox.Show($"Key: {smallestLargerEntryKey}");
                    foreach (var param in smallestLargerEntryValue)
                    {
                        //MessageBox.Show($"{param.Key}: {param.Value}");
                        if (param.Key == "Column0")
                        {
                            tenmay = param.Value;

                        }
                        if (param.Key == "Column2")
                        {
                             giatien = param.Value;
                            //MessageBox.Show(giatien);

                        }
                    }
                    foreach (var p in Diemdat)
                    {
                        //MessageBox.Show("Toibingu");
                        XYZ point = new XYZ(p.X, p.Y, 3200 / 304.8);
                        FamilySymbol symbol = CT.GetFamilySymbol(doc, tenmay, tenmay);
                        if (!symbol.IsActive) symbol.Activate();
                        //FamilyInstance support = doc.Create.NewFamilyInstance(point, symbol,level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        FamilyInstance support = doc.Create.NewFamilyInstance(point, symbol, roomLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                        support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(3200 / 304.8);

                        SetParameters(support, "PRICE", giatien);
                       
                       cassettes.Add(support);




                    }
                }


                if (smallestLargerEntryValue == null)
                {
                    IList<XYZ> Diemtruc1 = new List<XYZ>();


                    khoangcach = 4000 / 304.8;
                    if (trucY.Length > trucX.Length)
                    {



                        int sodiemtrentruc = (int)(trucY.Length / khoangcach);

                        double khoangcachtuong = (trucY.Length - khoangcach * sodiemtrentruc) / 2;
                        if (sodiemtrentruc == 0)
                        {

                            Diemtruc1.Add(diemgiuaphong);
                        }
                        if (khoangcachtuong < (1500 / 304.8) && sodiemtrentruc == 1)
                        {

                            Diemtruc1.Add(diemgiuaphong);
                        }

                        else
                        {

                            if (sodiemtrentruc % 2 == 0)
                            {
                                sodiemtrentruc = sodiemtrentruc - 1;
                            }
                            //MessageBox.Show(sodiemtrentruc.ToString());
                            for (int i = 1; i <= sodiemtrentruc; i++)
                                if ((trucY.Length - ((2 * i - 1) * khoangcach)) / 2 > 1500 / 304.8)
                                {  // Nếu chưa đạt điều kiện, tiếp tục thêm điểm vào danh sách
                                    Diemtruc1.Add(CT.FindTwoPointsOnLineFromPoint(diemgiuaphong, trucY, (2 * i - 1) * khoangcach).Item1);
                                    Diemtruc1.Add(CT.FindTwoPointsOnLineFromPoint(diemgiuaphong, trucY, (2 * i - 1) * khoangcach).Item2);
                                }

                        }
                    }
                    if (trucY.Length < trucX.Length)
                    {
                        int sodiemtrentruc = (int)(trucX.Length / khoangcach);

                        double khoangcachtuong = (trucX.Length - khoangcach * sodiemtrentruc) / 2;

                        if (sodiemtrentruc == 0)
                        {
                            //MessageBox.Show(sodiemtrentruc.ToString());
                            Diemtruc1.Add(diemgiuaphong);
                        }
                        if (khoangcachtuong < (1500 / 304.8) && sodiemtrentruc == 1)
                        {
                            Diemtruc1.Add(diemgiuaphong);
                        }
                        else
                        {
                            if (sodiemtrentruc % 2 == 0)
                            {
                                sodiemtrentruc = sodiemtrentruc - 1;
                            }
                            //MessageBox.Show(sodiemtrentruc.ToString());

                            for (int i = 1; i <= sodiemtrentruc; i++)

                            {  // Nếu chưa đạt điều kiện, tiếp tục thêm điểm vào danh sách
                                if ((trucX.Length - ((2 * i - 1) * khoangcach)) / 2 > 1500 / 304.8)
                                {
                                    Diemtruc1.Add(CT.FindTwoPointsOnLineFromPoint(diemgiuaphong, trucX, (2 * i - 1) * khoangcach).Item1);
                                    Diemtruc1.Add(CT.FindTwoPointsOnLineFromPoint(diemgiuaphong, trucX, (2 * i - 1) * khoangcach).Item2);
                                }

                            }

                        }
                    }




                    //MessageBox.Show(Diemtruc.Count.ToString());

                    IList<XYZ> Diemdat1 = new List<XYZ>();

                    foreach (XYZ p in Diemtruc1)
                    {
                        Line line = null;
                        if (trucY.Length > trucX.Length)
                        {
                            line = Line.CreateBound(p + new XYZ(500, 0, 0),
                               (p + new XYZ(-500, 0, 0)));
                        }
                        else
                        {
                            line = Line.CreateBound(p + new XYZ(0, 500, 0),
                                (p + new XYZ(0, -500, 0)));
                        }

                        IList<XYZ> exp = new List<XYZ>();//Điểm sau khi mở rộng giao với biên
                        foreach (Line linebound in boundlines)
                        {
                            IntersectionResultArray results1;
                            SetComparisonResult result1 = line.Intersect(linebound, out results1);
                            if (result1 == SetComparisonResult.Overlap && results1 != null && results1.Size > 0)
                            {
                                XYZ intersectionPoint = results1.get_Item(0).XYZPoint;
                                exp.Add(intersectionPoint);

                            }

                        }
                        Line newline = Line.CreateBound(exp[0], exp[1]);
                        //DetailLine detailLine = (DetailLine)doc.Create.NewDetailCurve(view, newline);
                        int sodiemtrentruc = (int)(newline.Length / khoangcach);

                        double khoangcachtuong = (newline.Length - khoangcach * sodiemtrentruc) / 2;

                        if (sodiemtrentruc == 0)
                        {
                            //MessageBox.Show(sodiemtrentruc.ToString());
                            Diemdat1.Add(p);
                        }
                        if (khoangcachtuong < (1500 / 304.8) && sodiemtrentruc == 1)
                        {
                            //MessageBox.Show("toi bi ngu");
                            Diemdat1.Add(p);
                        }
                        else
                        {

                            if (sodiemtrentruc % 2 == 0)
                            {
                                sodiemtrentruc = sodiemtrentruc - 1;
                            }
                            for (int i = 1; i <= sodiemtrentruc; i++)
                                if ((newline.Length - ((2 * i - 1) * khoangcach)) / 2 > 1500 / 304.8)
                                {  // Nếu chưa đạt điều kiện, tiếp tục thêm điểm vào danh sách
                                    Diemdat1.Add(CT.FindTwoPointsOnLineFromPoint(p, newline, (2 * i - 1) * khoangcach).Item1);
                                    Diemdat1.Add(CT.FindTwoPointsOnLineFromPoint(p, newline, (2 * i - 1) * khoangcach).Item2);
                                }

                        }
                    }

                    double congsuatmoimay1 = coolingcapacity / Diemdat1.Count;
                    //MessageBox.Show(congsuatmoimay.ToString());

                    // Khởi tạo biến để lưu kết quả
                    Dictionary<string, string> smallestLargerEntryValue1 = null;
                    double smallestLargerEntryKey1 = double.NaN;

                    // Duyệt qua từng phần tử trong IDictionary đã sắp xếp
                    foreach (DictionaryEntry entry in sortedCapacityModelTable)
                    {
                        // Ép kiểu entry.Key và entry.Value về đúng kiểu
                        double key = (double)entry.Key;
                        Dictionary<string, string> value = (Dictionary<string, string>)entry.Value;

                        // Kiểm tra nếu key của entry lớn hơn target
                        if (key > congsuatmoimay1) // là target của bạn
                        {
                            // Nếu chưa tìm thấy giá trị lớn hơn, hoặc nếu giá trị tìm thấy nhỏ hơn giá trị trước đó, cập nhật kết quả
                            smallestLargerEntryKey1 = key;
                            smallestLargerEntryValue1 = value;
                            break; // Sau khi tìm thấy giá trị lớn nhất nhỏ hơn target, thoát khỏi vòng lặp
                        }
                    }
                    string tenmay1 = null;
                    string giatien1 = null;


                    // Kiểm tra kết quả
                    if (smallestLargerEntryValue1 != null)
                    {
                        //MessageBox.Show($"Key: {smallestLargerEntryKey1}");
                        foreach (var param in smallestLargerEntryValue1)
                        {
                            //MessageBox.Show($"{param.Key}: {param.Value}");
                            if (param.Key == "Column0")
                            {
                                tenmay1 = param.Value;
                                //MessageBox.Show(tenmay1);
                            }

                        }
                        foreach (var param in smallestLargerEntryValue1)
                        {
                            //MessageBox.Show($"{param.Key}: {param.Value}");
                            if (param.Key == "Column2")
                            {
                                giatien1 = param.Value;
                                //MessageBox.Show(giatien1);
                            }

                        }
                        foreach (var p in Diemdat1)
                        {
                            //MessageBox.Show("Toibingu");
                            XYZ point = new XYZ(p.X, p.Y, 3200 / 304.8);
                            FamilySymbol symbol = CT.GetFamilySymbol(doc, tenmay1, tenmay1);
                            if (!symbol.IsActive) symbol.Activate();
                            //FamilyInstance support = doc.Create.NewFamilyInstance(point, symbol,level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            FamilyInstance support = doc.Create.NewFamilyInstance(point, symbol, roomLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                            support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(3200 / 304.8);
                            SetParameters(support, "PRICE", giatien1);
                            cassettes.Add(support);


                        }

                    }



                }
            }

        }


        private SortedList<double, Dictionary<string, string>> ReadExcelData(DataTable excelData)
        {
            SortedList<double, Dictionary<string, string>> sortedList = new SortedList<double, Dictionary<string, string>>();

            int rowCount = excelData.Rows.Count;
            int colCount = excelData.Columns.Count;
            int colNameroom = 1; // Cột chứa key (Capacity)

            for (int r = 1; r < rowCount; r++) // Duyệt từng hàng, bỏ qua hàng tiêu đề
            {
                // Lấy giá trị của cột Key (Capacity)
                string header = excelData.Columns[colNameroom].ColumnName;
                string parameterValue = excelData.Rows[r][header].ToString(); // Giá trị của hàng

                if (double.TryParse(parameterValue, out double key)) // Chuyển về số nếu có thể
                {
                    Dictionary<string, string> parameters = new Dictionary<string, string>();

                    for (int j = 0; j < colCount; j++) // Duyệt từng cột trong hàng
                    {
                        string columnName = excelData.Columns[j].ColumnName;
                        string paramValue = excelData.Rows[r][columnName].ToString(); // Lấy giá trị cột

                        if (!string.IsNullOrEmpty(paramValue))
                        {
                            parameters[columnName] = paramValue;
                        }
                    }

                    // Lưu vào SortedList với Capacity làm key, còn value là Dictionary chứa các thông tin khác
                    sortedList[key] = parameters;
                }
            }

            return sortedList;
        }
       



    }
}
        
    

