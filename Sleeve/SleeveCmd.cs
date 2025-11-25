using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using System.Windows.Media;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Visual;
using Autodesk.Revit.UI;
using DnBim_Tool;
using Autodesk.Revit.UI.Selection;
using System.Windows;
using System.Xml.Linq;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System.Security.Cryptography;
using System.Windows.Shapes;
using Line = Autodesk.Revit.DB.Line;
using System.IO;
using System.Windows.Documents;
using Dnbim_Tool.Sleeve;
using System.Collections;
using System.Drawing;
using Autodesk.Revit.DB.Structure;
using System.Windows.Media.Media3D;
using Autodesk.Revit.DB.ExtensibleStorage;
using System.Collections.Concurrent;
using System.Threading;


namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    //a
    public class SleeveCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var window = new SleeveView();
            window.ShowDialog();

            if (window.DialogResult == true)
            {
                try
                {
                    CreateSleeve(uiDoc, doc, window.isCeilingsChecked, window.isFloorsChecked, window.isWallsChecked, window.isPipeChecked, window.isDuctChecked, window.isCableTrayChecked);
                }
                catch { }

            }







            return Result.Succeeded;
        }







        public static void ShowHashtable(Hashtable hashtable)
        {
            StringBuilder sb = new StringBuilder();

            foreach (DictionaryEntry entry in hashtable)
            {
                sb.AppendLine($"Key: {entry.Key}, Value: {entry.Value}");
            }

            MessageBox.Show(sb.ToString(), "Nội dung Hashtable");
        }

        public static XYZ CalculateDuctFaceAreas(Document doc, Element element)
        {
            Duct duct = element as Duct;
            // Lấy GeometryElement của đối tượng ống gió
            Options geomOptions = new Options
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Medium,
            };

            GeometryElement geomElement = duct.get_Geometry(geomOptions);
            if (geomElement == null)
            {
                TaskDialog.Show("Error", "GeometryElement is null. Cannot process geometry.");
                return null;
            }

            Face largestFace = null;
            double largestArea = 0;
            XYZ largestFaceNormal = null;

            foreach (GeometryObject geomObj in geomElement)
            {
                if (geomObj is GeometryInstance geomInstance)
                {
                    GeometryElement instanceGeometry = geomInstance.GetInstanceGeometry();
                    if (instanceGeometry == null) continue;

                    foreach (GeometryObject instanceObj in instanceGeometry)
                    {
                        if (instanceObj is Solid solid)
                        {
                            if (solid == null || solid.Faces.Size == 0) continue;

                            foreach (Face face in solid.Faces)
                            {
                                double area = face.Area;

                                if (area > largestArea)
                                {
                                    largestArea = area;
                                    largestFace = face;
                                }
                            }
                        }
                    }
                }
                else if (geomObj is Solid solidObj)
                {
                    if (solidObj == null || solidObj.Faces.Size == 0) continue;

                    foreach (Face face in solidObj.Faces)
                    {
                        double area = face.Area;

                        if (area > largestArea)
                        {
                            largestArea = area;
                            largestFace = face;
                        }
                    }
                }
            }

            // Kiểm tra nếu không tìm thấy mặt nào
            if (largestFace == null)
            {
                TaskDialog.Show("Error", "No faces were found in the duct geometry.");
                return null;
            }

            // Lấy BoundingBoxUV để tìm một điểm hợp lệ trên mặt
            BoundingBoxUV boundingBox = largestFace.GetBoundingBox();
            if (boundingBox == null)
            {
                TaskDialog.Show("Error", "BoundingBoxUV is null for the largest face.");
                return null;
            }

            // Lấy một điểm giữa trong BoundingBoxUV
            UV midPoint = new UV(
                (boundingBox.Min.U + boundingBox.Max.U) / 2,
                (boundingBox.Min.V + boundingBox.Max.V) / 2
            );

            // Tính vector pháp tuyến tại điểm đã chọn


            try
            {
                largestFaceNormal = largestFace.ComputeNormal(midPoint);
                return largestFaceNormal;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to compute normal: {ex.Message}");
                return null;
            }

            // Hiển thị kết quả
            //string result = $\"Largest Face Area: {largestArea:F2} square units\\nNormal Vector: ({largestFaceNormal.X:F2}, {largestFaceNormal.Y:F2}, {largestFaceNormal.Z:F2})\";
            //TaskDialog.Show("Largest Face Info", result);
        }


        public static Solid GetMEPSolid(Element element)
        {

            Options options = new Options()
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Fine,
            };

            GeometryElement geometryElement = element.get_Geometry(options);

            foreach (GeometryObject geoOb in geometryElement)
            {
                if (geoOb is Solid solid) return solid;
            }

            return null;
        }

        public static void CreateSleeve(UIDocument uiDoc, Document doc
                                         , bool isCeilingsChecked, bool isFloorsChecked, bool isWallsChecked,
                                        bool isPipeChecked, bool isDuctChecked, bool isCableTrayChecked)


        {




            /*Floor*/
            //List<Element> mainElements = new List<Element>();
            //List<Element> linkedFloors = new List<Element>(); // Danh sách lưu các đối tượng Floors

            // Giả sử Element có định nghĩa các phương thức Equals và GetHashCode đúng đắn
            HashSet<Element> mainElements = new HashSet<Element>();
            HashSet<Element> linkedFloors = new HashSet<Element>();
            IList<Element> fileterlocationelement = new List<Element>();

            //IList<XYZ> diem = new List<XYZ>();






            // Thu thập tất cả các đối tượng RevitLinkInstance trong tài liệu chính
            var linkInstances = new FilteredElementCollector(doc)
                                 .OfClass(typeof(RevitLinkInstance))
                                 .Cast<RevitLinkInstance>()
                                 .Where(link => link.GetLinkDocument() != null) // Chỉ lấy file đang được load
                                 .GroupBy(link => link.Name) // Nhóm theo Name
                                 .Select(group => group.First()) // Lấy phần tử đầu tiên của mỗi nhóm
                                 .ToList();




            //MessageBox.Show(linkInstances.Count.ToString());


            foreach (RevitLinkInstance linkInstance in linkInstances)
            {

                // Lấy tài liệu liên kết
                Document linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null) continue; // Bỏ qua nếu không thể truy cập tài liệu liên kết

                // Thu thập tất cả các đối tượng trong tài liệu liên kết
                var flielinkedElements = new FilteredElementCollector(linkedDoc)
                    .WhereElementIsNotElementType(); // Lấy tất cả các instance, không lấy type

                // Danh sách các category cần lấy (Floors, Walls)
                IList<BuiltInCategory> categoriesToInclude = new List<BuiltInCategory>();

                if (isFloorsChecked == true)
                {
                    IList<BuiltInCategory> categoriesFloor = new List<BuiltInCategory>
                        {
                            BuiltInCategory.OST_Floors,
                        };
                    // Thêm các phần tử từ additionalCategories vào categoriesToInclude bằng vòng lặp
                    foreach (var category in categoriesFloor)
                    {
                        categoriesToInclude.Add(category);
                    }
                }
                if (isCeilingsChecked == true)
                {

                    IList<BuiltInCategory> categoriesCelling = new List<BuiltInCategory>
                        {
                            BuiltInCategory.OST_Ceilings
                        };
                    foreach (var category in categoriesCelling)
                    {
                        categoriesToInclude.Add(category);
                    }
                }
                if (isWallsChecked == true)
                {
                    IList<BuiltInCategory> categoriesWall = new List<BuiltInCategory>
                        {
                            BuiltInCategory.OST_Walls
                        };
                    foreach (var category in categoriesWall)
                    {
                        categoriesToInclude.Add(category);
                    }
                }

                // Lọc ra các đối tượng có Category là Floors hoặc Walls
                IList<Element> floorsAndWalls = flielinkedElements
                    .Where(e => e.Category != null && categoriesToInclude.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
                    .ToList();

                // Thêm các đối tượng Floors và Walls vào danh sách kết quả, đảm bảo không trùng lặp
                foreach (var element in floorsAndWalls)
                {
                    if (!linkedFloors.Any(f => f.Id == element.Id)) // Kiểm tra nếu chưa tồn tại
                    {
                        linkedFloors.Add(element);
                    }
                }





            }

            if (isDuctChecked == true)
            {

                mainElements.UnionWith(
                    new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .OfClass(typeof(Duct))
                        .ToList()
                );
            }
            if (isPipeChecked == true)
            {
                mainElements.UnionWith(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(Pipe))
                    .ToList()
            );
            }
            if (isCableTrayChecked == true)
            {
                mainElements.UnionWith(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(CableTray))
                    .ToList()
            );
            }





            foreach (var element in mainElements)
            {
                Line locationLine = null;

                if (element is Duct duct)
                {

                    locationLine = (duct.Location as LocationCurve).Curve as Line;
                    if (locationLine.Direction.AngleTo(XYZ.BasisZ) != 0 || locationLine.Direction.AngleTo(XYZ.BasisZ) != 180)
                    {
                        fileterlocationelement.Add(element);
                    }

                }
                if (element is Pipe pipe)
                {

                    locationLine = (pipe.Location as LocationCurve).Curve as Line;
                    if (locationLine.Direction.AngleTo(XYZ.BasisZ) != 0 || locationLine.Direction.AngleTo(XYZ.BasisZ) != 180)
                    {
                        fileterlocationelement.Add(element);
                    }

                }
                if (element is CableTray cableTray)
                {

                    locationLine = (cableTray.Location as LocationCurve).Curve as Line;
                    if (locationLine.Direction.AngleTo(XYZ.BasisZ) != 0 || locationLine.Direction.AngleTo(XYZ.BasisZ) != 180)
                    {
                        fileterlocationelement.Add(element);
                    }

                }
            }



            ConcurrentBag<Hashtable> bag = new ConcurrentBag<Hashtable>();
            AutoResetEvent autoEvent1 = new AutoResetEvent(false);
            Hashtable point = new Hashtable();
            Task t1 = Task.Factory.StartNew(() =>
            {
                foreach (Element mainElement in mainElements)
                {
                    ElementId elementId = mainElement.Id;
                    int elementIdInt = mainElement.Id.IntegerValue;
                    Line locationLine = null;
                    Hashtable hashtable = new Hashtable();


                    // Kiểm tra loại đối tượng trong tài liệu chính
                    if (mainElement is Duct duct)
                    {
                        locationLine = (duct.Location as LocationCurve).Curve as Line;

                    }
                    else if (mainElement is Pipe pipe)
                    {
                        locationLine = (pipe.Location as LocationCurve).Curve as Line;

                    }
                    else if (mainElement is CableTray cableTray)
                    {
                        locationLine = (cableTray.Location as LocationCurve).Curve as Line;
                    }

                    if (locationLine != null)
                    {
                        foreach (Element linkedElement in linkedFloors)
                        {

                            double segmentLength = 0;
                            Options geomOptions = new Options
                            {
                                ComputeReferences = true,
                                DetailLevel = ViewDetailLevel.Fine
                            };
                            List<XYZ> intersectionPoints = new List<XYZ>();
                            GeometryElement geomElement = linkedElement.get_Geometry(geomOptions);
                            if (geomElement != null)
                            {
                                foreach (GeometryObject geomObj in geomElement)
                                {

                                    if (geomObj is Solid solid)
                                    {
                                        foreach (Face face in solid.Faces)
                                        {
                                            IntersectionResultArray results;
                                            if (face.Intersect(locationLine, out results) == SetComparisonResult.Overlap)
                                            {
                                                foreach (IntersectionResult result in results)
                                                {
                                                    intersectionPoints.Add(result.XYZPoint);
                                                }
                                            }
                                        }
                                    }
                                }

                                // 5. Xử lý nếu tìm thấy điểm giao
                                if (intersectionPoints.Count == 2)
                                {

                                    for (int i = 0; i < intersectionPoints.Count - 1; i += 2)
                                    {
                                        XYZ p1 = intersectionPoints[i];
                                        XYZ p2 = intersectionPoints[i + 1];
                                        segmentLength = p1.DistanceTo(p2);


                                        XYZ diemdat = new XYZ(
                                                           (p1.X + p2.X) / 2, // Lấy đến 3 chữ số thập phân
                                                           (p1.Y + p2.Y) / 2, // Lấy đến 3 chữ số thập phân
                                                           (p1.Z + p2.Z) / 2  // Lấy đến 3 chữ số thập phân
                                                       );



                                        hashtable.Add(diemdat, segmentLength);






                                    }


                                }
                            }
                        }
                    }
                    point.Add(elementIdInt, hashtable);
                    bag.Add(point);
                }
            });

            t1.Wait();






            Hashtable newhashtable = RemoveDuplicateKeys(point);

            



            foreach (DictionaryEntry entry0 in newhashtable) // Lặp qua DictionaryEntry của newhashtable
            {
                using (Transaction trans = new Transaction(doc, "Place Family Instances"))
                {
                    trans.Start();
                    ElementId elementId = new ElementId((int)entry0.Key);
                    Element element = doc.GetElement(elementId);

                    // Lấy Hashtable con từ Value
                    Hashtable pointTable = (Hashtable)entry0.Value;

                    foreach (DictionaryEntry entry1 in pointTable)
                    {
                        XYZ diemsupport = (XYZ)entry1.Key;    // Lấy tọa độ
                        double value = (double)entry1.Value;  // Lấy giá trị

                        double height = 0;
                        double width = 0;

                        if (element is Duct duct)
                        {//a
                            Line locationLine = (element.Location as LocationCurve).Curve as Line;

                            height = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble()
                                    + 2 * duct.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble();
                            width = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble()
                                    + 2 * duct.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble();
                            FamilySymbol symbol = CT.GetFamilySymbol(doc, "Sleeve_Rectangular", "Sleeve_Rectangular");
                            
                            if (!symbol.IsActive) symbol.Activate();
                            if (!CT.Checkexist(doc, diemsupport))
                            {
                                double anglesleeve = 0;
                                if (Math.Round((locationLine.Direction.AngleTo(XYZ.BasisZ) * 180 / Math.PI), 0) == 0 || Math.Round((locationLine.Direction.AngleTo(XYZ.BasisZ) * 180 / Math.PI), 0) == 180)
                                {
                                    anglesleeve = 90 * Math.PI / 180;
                                }
                                else
                                {
                                    anglesleeve = 90 * Math.PI / 180;
                                }
                                // Tạo FamilyInstance
                                FamilyInstance support = doc.Create.NewFamilyInstance(
                                    diemsupport,
                                    symbol,
                                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural
                                );

                                support.LookupParameter("Sleeve_Width").Set(width);
                                support.LookupParameter("Sleeve_Height").Set(height);
                                support.LookupParameter("Sleeve_Length").Set(value);
                                support.LookupParameter("Sleeve_Angle").Set(anglesleeve);


                                XYZ vecto = CalculateDuctFaceAreas(doc, element);
                                double angle = Math.Round(vecto.AngleTo(XYZ.BasisY) * 180 / Math.PI, 0);
                                double angleZ = Math.Round(locationLine.Direction.AngleTo(XYZ.BasisZ) * 180 / Math.PI, 0);
                                if (angle == 90)

                                {


                                    ElementTransformUtils.RotateElement(doc, support.Id, locationLine, Math.PI / 2);

                                }
                                if (height > width)

                                {


                                    ElementTransformUtils.RotateElement(doc, support.Id, locationLine, Math.PI / 2);

                                }
                                if (angleZ != 90)

                                {



                                }
                                else
                                {

                                    ElementTransformUtils.RotateElement(doc, support.Id, locationLine, Math.PI / 2);
                                }
                            }
                        }
                        if (element is Pipe pipe)
                        {
                            double diameter = Math.Round((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() * 1
                                                + 2 * pipe.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                            double anglesleeve = 0;
                            Line locationLine = (element.Location as LocationCurve).Curve as Line;
                            if (Math.Round((locationLine.Direction.AngleTo(XYZ.BasisZ) * 180 / Math.PI), 0) == 0 || Math.Round((locationLine.Direction.AngleTo(XYZ.BasisZ) * 180 / Math.PI), 0) == 180)
                            {
                                anglesleeve = 90 * Math.PI / 180;
                            }
                            else
                            {
                                anglesleeve = 0;
                            }

                            FamilySymbol symbol = CT.GetFamilySymbol(doc, "Sleeve_Round", "Sleeve_Round");
                            if (!symbol.IsActive) symbol.Activate();


                            if (!CT.Checkexist(doc, diemsupport))
                            {

                                FamilyInstance support = doc.Create.NewFamilyInstance(
                                diemsupport,
                                symbol,
                                Autodesk.Revit.DB.Structure.StructuralType.NonStructural
                            );

                                //MessageBox.Show(linkedElement.Id.ToString());
                                support.LookupParameter("Sleeve_Diameter").Set(diameter);
                                support.LookupParameter("Sleeve_Length").Set(value);
                                support.LookupParameter("Sleeve_Angle").Set(anglesleeve);


                                if (Math.Round((locationLine.Direction.AngleTo(XYZ.BasisX) * 180 / Math.PI), 0) == 0 || Math.Round((locationLine.Direction.AngleTo(XYZ.BasisX) * 180 / Math.PI), 0) == 180)

                                {


                                    Line linephu = Line.CreateBound(diemsupport, diemsupport + new XYZ(0, 0, 1));
                                    ElementTransformUtils.RotateElement(doc, support.Id, linephu, Math.PI / 2);

                                }
                            }
                        }



                    }
                    trans.Commit();
                }

            }
        }

        static string JoinHashtable(Hashtable hashtable, int indentLevel = 0)
        {
            StringBuilder sb = new StringBuilder();
            string indent = new string(' ', indentLevel * 4); // Tạo thụt lề

            foreach (DictionaryEntry entry in hashtable)
            {
                if (entry.Value is Hashtable nestedHashtable)
                {
                    // Nếu giá trị là Hashtable, gọi đệ quy
                    sb.AppendLine($"{indent}🔹 Key: {entry.Key}, Value: (Nested Hashtable)");
                    sb.Append(JoinHashtable(nestedHashtable, indentLevel + 1)); // Tăng mức indent
                }
                else
                {
                    // Giá trị không phải Hashtable, in bình thường
                    sb.AppendLine($"{indent}🔹 Key: {entry.Key}, Value: {entry.Value}");
                }
            }

            return sb.ToString();
        }

        static Hashtable RemoveDuplicateKeys(Hashtable hashtable)
        {
            Hashtable result = new Hashtable();
            HashSet<string> seenValues = new HashSet<string>(); // Sử dụng HashSet để theo dõi các giá trị đã gặp

            foreach (DictionaryEntry entry in hashtable)
            {
                // Chuyển giá trị thành chuỗi để so sánh (có thể tùy chỉnh nếu giá trị phức tạp hơn)
                string keyString = entry.Key.ToString();

                // Nếu chưa gặp giá trị này, thêm vào Hashtable kết quả
                if (!seenValues.Contains(keyString))
                {
                    seenValues.Add(keyString);
                    result.Add(entry.Key, entry.Value);
                }
            }

            return result;
        }
    }

}




//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Collections;
//using System.Collections.Concurrent;
//using System.Threading;
//using System.Windows;

//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.DB.Visual;
//using Autodesk.Revit.DB.Structure;
//using Autodesk.Revit.UI;
//using Autodesk.Revit.UI.Selection;

//// alias tránh đụng tên
//using Line = Autodesk.Revit.DB.Line;
//using Autodesk.Revit.DB.Electrical;
//using Autodesk.Revit.DB.Mechanical;
//using Autodesk.Revit.DB.Plumbing;
//using Dnbim_Tool.Sleeve;

//namespace DnBim_Tool
//{
//    [Transaction(TransactionMode.Manual)]
//    public class SleeveCmd : IExternalCommand
//    {
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
//            Document doc = uiDoc.Document;

//            var window = new SleeveView();
//            window.ShowDialog();

//            if (window.DialogResult == true)
//            {
//                try
//                {
//                    CreateSleeve(
//                        uiDoc,
//                        doc,
//                        window.isCeilingsChecked,
//                        window.isFloorsChecked,
//                        window.isWallsChecked,
//                        window.isPipeChecked,
//                        window.isDuctChecked,
//                        window.isCableTrayChecked
//                    );
//                }
//                catch (Exception ex)
//                {
//                    TaskDialog.Show("Sleeve", $"Đã xảy ra lỗi: {ex.Message}");
//                }
//            }

//            return Result.Succeeded;
//        }

//        // --------------------------
//        //  A. Collector theo View
//        // --------------------------
//        private static IList<Element> GetVisibleMepElementsInView(
//            Document doc, View view,
//            bool includeDucts, bool includePipes, bool includeCableTrays)
//        {
//            var catFilters = new List<ElementFilter>();
//            if (includeDucts) catFilters.Add(new ElementCategoryFilter(BuiltInCategory.OST_DuctCurves));
//            if (includePipes) catFilters.Add(new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves));
//            if (includeCableTrays) catFilters.Add(new ElementCategoryFilter(BuiltInCategory.OST_CableTray));

//            if (catFilters.Count == 0) return new List<Element>();

//            ElementFilter catOr = (catFilters.Count == 1)
//                ? catFilters[0]
//                : (ElementFilter)new LogicalOrFilter(catFilters);

//            // Collector theo view: đã tôn trọng discipline, view range, crop/section box, filters...
//            var inView = new FilteredElementCollector(doc, view.Id)
//                .WhereElementIsNotElementType()
//                .WherePasses(catOr)
//                .ToElements();

//            var result = new List<Element>();
//            foreach (var e in inView)
//            {
//                if (e.IsHidden(view)) continue; // hide theo element
//                var cat = e.Category;
//                if (cat != null && view.GetCategoryHidden(cat.Id)) continue; // tắt category
//                if (!PassesView3DSectionBox(view, e)) continue; // siết theo section box (nếu có)
//                result.Add(e);
//            }
//            return result;
//        }

//        private static bool PassesView3DSectionBox(View view, Element e)
//        {
//            var v3 = view as View3D;
//            if (v3 == null || !v3.IsSectionBoxActive) return true;
//            var bb = e.get_BoundingBox(view);
//            if (bb == null) return false;
//            // Kiểm tra giao sơ bộ bằng AABB trong section box
//            var sb = v3.GetSectionBox();
//            return AabbIntersects(sb, bb);
//        }

//        private static bool AabbIntersects(BoundingBoxXYZ box, BoundingBoxXYZ bb)
//        {
//            return !(bb.Max.X < box.Min.X || bb.Min.X > box.Max.X
//                  || bb.Max.Y < box.Min.Y || bb.Min.Y > box.Max.Y
//                  || bb.Max.Z < box.Min.Z || bb.Min.Z > box.Max.Z);
//        }

//        private static IList<RevitLinkInstance> GetVisibleLinkInstancesInView(Document doc, View view)
//        {
//            var links = new FilteredElementCollector(doc, view.Id)
//                .OfClass(typeof(RevitLinkInstance))
//                .Cast<RevitLinkInstance>()
//                .ToList();

//            var visible = new List<RevitLinkInstance>();
//            foreach (var li in links)
//            {
//                if (li.IsHidden(view)) continue;
//                var cat = li.Category;
//                if (cat != null && view.GetCategoryHidden(cat.Id)) continue;
//                if (!PassesView3DSectionBox(view, li)) continue;
//                visible.Add(li);
//            }
//            return visible;
//        }

//        // Overload: dùng bbox của link
//        private static bool PassesView3DSectionBox(View view, RevitLinkInstance li)
//        {
//            var v3 = view as View3D;
//            if (v3 == null || !v3.IsSectionBoxActive) return true;
//            var bb = li.get_BoundingBox(view);
//            if (bb == null) return false;
//            var sb = v3.GetSectionBox();
//            return AabbIntersects(sb, bb);
//        }

//        // --------------------------
//        //  B. Hình học duct lớn nhất
//        // --------------------------
//        public static XYZ CalculateDuctFaceNormal(Document doc, Element element)
//        {
//            if (!(element is  Duct duct)) return null;

//            Options geomOptions = new Options
//            {
//                ComputeReferences = true,
//                DetailLevel = ViewDetailLevel.Medium,
//            };

//            GeometryElement geomElement = duct.get_Geometry(geomOptions);
//            if (geomElement == null) return null;

//            Face largestFace = null;
//            double largestArea = 0;

//            foreach (GeometryObject geomObj in geomElement)
//            {
//                if (geomObj is GeometryInstance gi)
//                {
//                    var instanceGeometry = gi.GetInstanceGeometry();
//                    if (instanceGeometry == null) continue;
//                    foreach (GeometryObject instanceObj in instanceGeometry)
//                    {
//                        if (instanceObj is Solid s && s.Faces.Size > 0)
//                        {
//                            foreach (Face f in s.Faces)
//                            {
//                                double area = f.Area;
//                                if (area > largestArea)
//                                {
//                                    largestArea = area;
//                                    largestFace = f;
//                                }
//                            }
//                        }
//                    }
//                }
//                else if (geomObj is Solid so && so.Faces.Size > 0)
//                {
//                    foreach (Face f in so.Faces)
//                    {
//                        double area = f.Area;
//                        if (area > largestArea)
//                        {
//                            largestArea = area;
//                            largestFace = f;
//                        }
//                    }
//                }
//            }

//            if (largestFace == null) return null;

//            var bb = largestFace.GetBoundingBox();
//            if (bb == null) return null;

//            UV mid = new UV((bb.Min.U + bb.Max.U) * 0.5, (bb.Min.V + bb.Max.V) * 0.5);
//            try
//            {
//                return largestFace.ComputeNormal(mid);
//            }
//            catch
//            {
//                return null;
//            }
//        }

//        // --------------------------
//        //  C. Tạo Sleeve
//        // --------------------------
//        public static void CreateSleeve(
//            UIDocument uiDoc, Document doc,
//            bool isCeilingsChecked, bool isFloorsChecked, bool isWallsChecked,
//            bool isPipeChecked, bool isDuctChecked, bool isCableTrayChecked)
//        {
//            var activeView = uiDoc.ActiveView;

//            // 1) Chỉ lấy Duct/Pipe/CableTray đang hiển thị trong view
//            var mainElements = GetVisibleMepElementsInView(doc, activeView,
//                isDuctChecked, isPipeChecked, isCableTrayChecked);

//            // 2) Chỉ lấy Link hiển thị trong view
//            var linkInstances = GetVisibleLinkInstancesInView(doc, activeView);

//            // 3) Gom các đối tượng (Floors/Walls/Ceilings) từ link + kèm transform link => để chuyển hình học sang hệ toạ độ host
//            var categoriesToInclude = new List<BuiltInCategory>();
//            if (isFloorsChecked) categoriesToInclude.Add(BuiltInCategory.OST_Floors);
//            if (isCeilingsChecked) categoriesToInclude.Add(BuiltInCategory.OST_Ceilings);
//            if (isWallsChecked) categoriesToInclude.Add(BuiltInCategory.OST_Walls);

//            var linkedSolids = new List<(Solid SolidInHost, ElementId LinkedElementId)>(); // solid đã transform về host

//            if (categoriesToInclude.Count > 0)
//            {
//                ElementFilter orFilter;
//                if (categoriesToInclude.Count == 1)
//                {
//                    orFilter = new ElementCategoryFilter(categoriesToInclude[0]);
//                }
//                else
//                {
//                    IList<ElementFilter> filters = categoriesToInclude
//                        .Select(c => (ElementFilter)new ElementCategoryFilter(c))
//                        .ToList();

//                    orFilter = new LogicalOrFilter(filters);
//                }

//                foreach (var linkInstance in linkInstances)
//                {
//                    var linkedDoc = linkInstance.GetLinkDocument();
//                    if (linkedDoc == null) continue;

//                    var elems = new FilteredElementCollector(linkedDoc)
//                        .WhereElementIsNotElementType()
//                        .WherePasses(orFilter)
//                        .ToElements();

//                    // transform link => host
//                    Transform t = linkInstance.GetTotalTransform();
//                    Options go = new Options
//                    {
//                        ComputeReferences = true,
//                        DetailLevel = ViewDetailLevel.Fine
//                    };

//                    foreach (var el in elems)
//                    {
//                        var ge = el.get_Geometry(go);
//                        if (ge == null) continue;

//                        foreach (GeometryObject g in ge)
//                        {
//                            if (g is Solid s && s.Volume > 1e-9 && s.Faces.Size > 0)
//                            {
//                                // chuyển solid của link về hệ toạ độ host
//                                Solid sHost = SolidUtils.CreateTransformed(s, t);
//                                linkedSolids.Add((sHost, el.Id));
//                            }
//                            else if (g is GeometryInstance gi)
//                            {
//                                var ig = gi.GetInstanceGeometry();
//                                if (ig == null) continue;
//                                foreach (GeometryObject igg in ig)
//                                {
//                                    if (igg is Solid si && si.Volume > 1e-9 && si.Faces.Size > 0)
//                                    {
//                                        Solid sHost = SolidUtils.CreateTransformed(si, t);
//                                        linkedSolids.Add((sHost, el.Id));
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//            }

//            // 4) Dò giao tuyến: line của MEP với các mặt của solid link
//            //    Với mỗi ME P element: lấy LocationCurve -> Line, nếu không vertical hoàn toàn thì xét giao với các faces
//            var elementToSleeveSegments = new Dictionary<int, List<(XYZ Mid, double Length)>>();

//            foreach (var mainElement in mainElements)
//            {
//                Line locLine = GetLocationLine(mainElement);
//                if (locLine == null) continue;

//                // Nếu muốn bỏ ống hoàn toàn thẳng đứng:
//                double angleToZ = locLine.Direction.AngleTo(XYZ.BasisZ);
//                bool isVertical = (Math.Abs(angleToZ) < 1e-6) || (Math.Abs(angleToZ - Math.PI) < 1e-6);
//                // => tuỳ ý: nếu không muốn đặt sleeve cho ống đứng, bật continue dòng dưới
//                // if (isVertical) continue;

//                var intersections = new List<XYZ>();

//                foreach (var (solidHost, _) in linkedSolids)
//                {
//                    foreach (Face f in solidHost.Faces)
//                    {
//                        if (f == null) continue;
//                        if (f.Intersect(locLine, out IntersectionResultArray ira) == SetComparisonResult.Overlap && ira != null && ira.Size > 0)
//                        {
//                            foreach (IntersectionResult ir in ira)
//                            {
//                                if (ir != null) intersections.Add(ir.XYZPoint);
//                            }
//                        }
//                    }
//                }

//                // gom thành từng cặp p1-p2 (vào-ra), tính mid & segment length
//                if (intersections.Count >= 2)
//                {
//                    // sắp xếp theo tham số dọc theo line để ghép cặp ổn định
//                    var sorted = intersections
//                        .Select(p => new { P = p, Param = locLine.Project(p).Parameter })
//                        .OrderBy(x => x.Param)
//                        .Select(x => x.P)
//                        .ToList();

//                    for (int i = 0; i + 1 < sorted.Count; i += 2)
//                    {
//                        XYZ p1 = sorted[i];
//                        XYZ p2 = sorted[i + 1];
//                        double segLen = p1.DistanceTo(p2);
//                        XYZ mid = (p1 + p2) * 0.5;

//                        int idInt = mainElement.Id.IntegerValue;
//                        if (!elementToSleeveSegments.ContainsKey(idInt))
//                            elementToSleeveSegments[idInt] = new List<(XYZ, double)>();

//                        elementToSleeveSegments[idInt].Add((mid, segLen));
//                    }
//                }
//            }

//            // 5) Đặt family instances
//            using (Transaction tr = new Transaction(doc, "Place Sleeves (Visible in View)"))
//            {
//                tr.Start();

//                foreach (var kv in elementToSleeveSegments)
//                {
//                    var eid = new ElementId(kv.Key);
//                    var mepElement = doc.GetElement(eid);
//                    if (mepElement == null) continue;

//                    var segments = kv.Value;
//                    foreach (var (pt, len) in segments)
//                    {
//                        // chống trùng nếu bạn có CT.Checkexist
//                        if (CT.Checkexist(doc, pt)) continue;

//                        if (mepElement is Duct duct)
//                        {
//                            PlaceRectSleeveForDuct(doc, duct, pt, len);
//                        }
//                        else if (mepElement is Pipe pipe)
//                        {
//                            PlaceRoundSleeveForPipe(doc, pipe, pt, len);
//                        }
//                        else if (mepElement is CableTray tray)
//                        {
//                            PlaceRectSleeveForCableTray(doc, tray, pt, len);
//                        }
//                    }
//                }

//                tr.Commit();
//            }
//        }

//        // --------------------------
//        //  D. Helpers đặt family
//        // --------------------------
//        private static Line GetLocationLine(Element e)
//        {
//            var lc = (e.Location as LocationCurve)?.Curve as Line;
//            return lc;
//        }

//        private static void PlaceRectSleeveForDuct(Document doc, Duct duct, XYZ point, double sleeveLen)
//        {
//            var sym = CT.GetFamilySymbol(doc, "Sleeve_Rectangular", "Sleeve_Rectangular");
//            if (!sym.IsActive) sym.Activate();

//            // tính kích thước phủ bảo ôn
//            double h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble()
//                     + 2 * duct.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble();
//            double w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble()
//                     + 2 * duct.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble();

//            var fi = doc.Create.NewFamilyInstance(point, sym, StructuralType.NonStructural);

//            fi.LookupParameter("Sleeve_Width")?.Set(w);
//            fi.LookupParameter("Sleeve_Height")?.Set(h);
//            fi.LookupParameter("Sleeve_Length")?.Set(sleeveLen);

//            // góc nghiêng sleeve theo hướng ống
//            Line axis = GetLocationLine(duct);
//            double angleZdeg = RadianToDeg(axis.Direction.AngleTo(XYZ.BasisZ));
//            double sleeveAngle = (Math.Abs(angleZdeg) < 1 || Math.Abs(angleZdeg - 180) < 1) ? DegToRadian(90) : DegToRadian(90); // giữ logic cũ của bạn
//            fi.LookupParameter("Sleeve_Angle")?.Set(sleeveAngle);

//            // căn chỉnh xoay thêm nếu cần
//            var n = CalculateDuctFaceNormal(doc, duct);
//            if (n != null)
//            {
//                double a = Math.Round(RadianToDeg(n.AngleTo(XYZ.BasisY)));
//                if (Math.Abs(a - 90) < 1) ElementTransformUtils.RotateElement(doc, fi.Id, axis, Math.PI / 2);
//            }

//            if (h > w) ElementTransformUtils.RotateElement(doc, fi.Id, axis, Math.PI / 2);

//            // nếu ống nằm ngang (song song X), xoay thêm để hướng trục sleeve hợp lý
//            double angleToX = RadianToDeg(axis.Direction.AngleTo(XYZ.BasisX));
//            if (Math.Abs(angleToX) < 1 || Math.Abs(angleToX - 180) < 1)
//            {
//                var aux = Line.CreateBound(point, point + XYZ.BasisZ);
//                ElementTransformUtils.RotateElement(doc, fi.Id, aux, Math.PI / 2);
//            }
//        }

//        private static void PlaceRoundSleeveForPipe(Document doc, Pipe pipe, XYZ point, double sleeveLen)
//        {
//            var sym = CT.GetFamilySymbol(doc, "Sleeve_Round", "Sleeve_Round");
//            if (!sym.IsActive) sym.Activate();

//            double dOut = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
//            double ins = pipe.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble();
//            double dia = Math.Round(dOut + 2 * ins, 6);

//            Line axis = GetLocationLine(pipe);
//            double angleZdeg = RadianToDeg(axis.Direction.AngleTo(XYZ.BasisZ));
//            double sleeveAngle = (Math.Abs(angleZdeg) < 1 || Math.Abs(angleZdeg - 180) < 1) ? DegToRadian(90) : 0;

//            var fi = doc.Create.NewFamilyInstance(point, sym, StructuralType.NonStructural);
//            fi.LookupParameter("Sleeve_Diameter")?.Set(dia);
//            fi.LookupParameter("Sleeve_Length")?.Set(sleeveLen);
//            fi.LookupParameter("Sleeve_Angle")?.Set(sleeveAngle);

//            // nếu pipe song song trục X: xoay để hướng sleeve đứng
//            double angleToX = RadianToDeg(axis.Direction.AngleTo(XYZ.BasisX));
//            if (Math.Abs(angleToX) < 1 || Math.Abs(angleToX - 180) < 1)
//            {
//                var aux = Line.CreateBound(point, point + XYZ.BasisZ);
//                ElementTransformUtils.RotateElement(doc, fi.Id, aux, Math.PI / 2);
//            }
//        }

//        private static void PlaceRectSleeveForCableTray(Document doc, CableTray tray, XYZ point, double sleeveLen)
//        {
//            // Dùng cùng family Rectangular (hoặc đổi tên theo library của bạn)
//            var sym = CT.GetFamilySymbol(doc, "Sleeve_Rectangular", "Sleeve_Rectangular");
//            if (!sym.IsActive) sym.Activate();

//            // Lấy kích thước khay (param có thể khác nhau theo template, bạn điều chỉnh nếu tên khác)
//            double h = tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).AsDouble();
//            double w = tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).AsDouble();

//            var fi = doc.Create.NewFamilyInstance(point, sym, StructuralType.NonStructural);

//            fi.LookupParameter("Sleeve_Width")?.Set(w);
//            fi.LookupParameter("Sleeve_Height")?.Set(h);
//            fi.LookupParameter("Sleeve_Length")?.Set(sleeveLen);

//            // góc
//            var axis = GetLocationLine(tray);
//            double angleZdeg = RadianToDeg(axis.Direction.AngleTo(XYZ.BasisZ));
//            double sleeveAngle = (Math.Abs(angleZdeg) < 1 || Math.Abs(angleZdeg - 180) < 1) ? DegToRadian(90) : DegToRadian(90);
//            fi.LookupParameter("Sleeve_Angle")?.Set(sleeveAngle);

//            if (h > w) ElementTransformUtils.RotateElement(doc, fi.Id, axis, Math.PI / 2);

//            // nếu nằm ngang theo X, xoay cho thẩm mỹ
//            double angleToX = RadianToDeg(axis.Direction.AngleTo(XYZ.BasisX));
//            if (Math.Abs(angleToX) < 1 || Math.Abs(angleToX - 180) < 1)
//            {
//                var aux = Line.CreateBound(point, point + XYZ.BasisZ);
//                ElementTransformUtils.RotateElement(doc, fi.Id, aux, Math.PI / 2);
//            }
//        }

//        // --------------------------
//        //  E. Tiện ích
//        // --------------------------
//        private static double RadianToDeg(double r) => r * 180.0 / Math.PI;
//        private static double DegToRadian(double d) => d * Math.PI / 180.0;

//        // (tuỳ chọn) bạn vẫn có thể giữ các hàm dưới nếu cần
//        public static Solid GetMEPSolid(Element element)
//        {
//            Options options = new Options()
//            {
//                ComputeReferences = true,
//                DetailLevel = ViewDetailLevel.Fine,
//            };
//            GeometryElement geometryElement = element.get_Geometry(options);
//            foreach (GeometryObject geoOb in geometryElement)
//            {
//                if (geoOb is Solid solid && solid.Volume > 1e-9)
//                    return solid;
//            }
//            return null;
//        }
//    }
//}
