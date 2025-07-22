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


namespace Dnbim_Tool
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
                    CreateSleeve(uiDoc, doc,window.isCeilingsChecked,window.isFloorsChecked,window.isWallsChecked,window.isPipeChecked,window.isDuctChecked,window.isCableTrayChecked);
                }
                catch { }

            }







            return Result.Succeeded;
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

                if(isFloorsChecked==true)
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

                                       
                                               XYZ   diemdat = new XYZ(
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
    

    

