using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Line = Autodesk.Revit.DB.Line;
using System.Xml.Linq;
using System.Windows.Media.Imaging;
using System.Drawing.Imaging;
using System.Security.Cryptography;
using System.Linq.Expressions;
using Autodesk.Revit.DB.Electrical;
using System.Windows.Controls;
using Color = Autodesk.Revit.DB.Color;
using Autodesk.Revit.DB.Architecture;
using static Dnbim_Tool.TreeNodeEvent;
using System.Collections;
using Autodesk.Revit.UI.Selection;
using System.Windows.Media.Media3D;

namespace DnBim_Tool
{
    public class CT
    {
        public static double ToRadian(double degree)
        {
            return degree * Math.PI / 180;
        }
        public static PlanarFace FindPlanarFace(Element element, XYZ pos)
        {
            // 1. Lấy GeometryElement của đối tượng
            Options ops = new Options
            {
                ComputeReferences = true,  // Cho phép lấy tham chiếu đến các mặt
                DetailLevel = ViewDetailLevel.Fine  // Mức độ chi tiết cao
            };

            GeometryElement geometryElement = element.get_Geometry(ops);

            // 2. Duyệt qua các GeometryObject trong GeometryElement
            foreach (GeometryObject geomObj in geometryElement)
            {
                if (geomObj is Solid solid && solid != null)
                {
                    // 3. Duyệt qua tất cả các Face trong Solid
                    foreach (Face face in solid.Faces)
                    {
                        try
                        {
                            // 4. Chiếu điểm đã chọn (pos) lên mặt phẳng
                            IntersectionResult result = face.Project(pos);
                            double space = Math.Round(result.XYZPoint.DistanceTo(pos), 0);
                            if (result != null && space == 0)
                            {
                                // 5. Trả về mặt phẳng (PlanarFace) nếu tìm thấy
                                return face as PlanarFace;
                            }
                        }
                        catch (NullReferenceException)
                        {

                        }
                    }
                }
            }

            // 6. Trả về null nếu không tìm thấy mặt phẳng phù hợp
            return null;
        }
        public static Line FindDirectionOfElement(Reference r1, LocationCurve locationCurve)
        {
            Line pipeLine = locationCurve.Curve as Line;
            XYZ p1 = pipeLine.GetEndPoint(0);
            XYZ p2 = pipeLine.GetEndPoint(1);
            XYZ p3 = r1.GlobalPoint;

            double d1 = p3.DistanceTo(p1);
            double d2 = p3.DistanceTo(p2);
            if (d2 < d1)
            {
                Line pipeline1 = Line.CreateBound(p2, p1);
                return pipeline1;
            }

            return pipeLine;
        }
        public static XYZ FindPointOnLineFromStartPoint(Line line, double distance)
        {
            //get point
            XYZ A = line.GetEndPoint(0);
            XYZ B = line.GetEndPoint(1);
            XYZ AB = B - A;
            double tile = distance / AB.GetLength();

            double x = tile * AB.X + A.X;
            double y = tile * AB.Y + A.Y;
            double z = tile * AB.Z + A.Z;

            return new XYZ(x, y, z);
        }
        public static (XYZ, XYZ) FindTwoPointsOnLineFromPoint(XYZ P, Line line, double distance)
        {
            // Lấy điểm đầu và điểm cuối của đoạn thẳng
            XYZ A = line.GetEndPoint(0);
            XYZ B = line.GetEndPoint(1);

            // Tính vectơ chỉ phương của đoạn thẳng (hướng từ A đến B)
            XYZ AB = B - A;
            double lineLength = AB.GetLength();
            XYZ unitDirection = AB.Normalize(); // Chỉ phương của đoạn thẳng

            // Tính khoảng cách từ P đến điểm đầu A
            XYZ AP = P - A;
            double t = AP.DotProduct(unitDirection); // Projection của P lên đoạn thẳng AB

            // Tính hai điểm cách đều từ P trên đoạn thẳng
            XYZ point1 = P + unitDirection * distance/2; // Điểm mở theo cùng phương
            XYZ point2 = P - unitDirection * distance/2; // Điểm mở theo chiều ngược lại

            return (point1, point2);
        }
        public static XYZ FindPointOnLineFromEndPoint(Line line, double distance)
        {

            XYZ A = line.GetEndPoint(0);
            XYZ B = line.GetEndPoint(1);
            XYZ vectorAB = A - B;

            double title = distance / vectorAB.GetLength();
            double x = title * vectorAB.X + B.X;
            double y = title * vectorAB.Y + B.Y;
            double z = title * vectorAB.Z + B.Z;



            return new XYZ(x, y, z);
        }
        public static FamilySymbol GetFamilySymbol(Document doc, string familyName, string typeName)
        {
            // Lọc tất cả các FamilySymbol trong tài liệu
            var listFamilySymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)) // Chỉ lấy các FamilySymbol (ElementType)
                .Cast<FamilySymbol>()
                .ToList();

            // Tìm FamilySymbol theo FamilyName và TypeName
            var symbol = listFamilySymbols
                .Find(x => x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase)
                           && x.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));

            return symbol; // Trả về FamilySymbol nếu tìm thấy

        }
        public static XYZ PointIntersectPlane(XYZ point, Plane plane)
        {




            XYZ sp = point;
            XYZ ep = point + new XYZ(0, 0, 1000);


            //vecto đơn vị
            XYZ normalize = (ep - sp).Normalize();

            double distance = (plane.Normal.DotProduct(plane.Origin) - plane.Normal.DotProduct(sp)) / plane.Normal.DotProduct(normalize);

            XYZ intersectPoint = sp + distance * normalize;

            return intersectPoint;
        }
        public static XYZ PointOnLine(XYZ point, Line pipeLine)
        {
            XYZ pointZ = point - new XYZ(0, 0, 1000);
            Line lineZ = Line.CreateBound(point, pointZ);
            pipeLine.Intersect(lineZ, out IntersectionResultArray array);
            if (array != null && array.Size == 1)
            {
                return array.get_Item(0).XYZPoint;
            }
            return null;
        }
        public static XYZ LineIntersectPlane(Line line, Plane plane)
        {


            XYZ normal = plane.Normal;
            XYZ origin = plane.Origin;

            XYZ sp = line.GetEndPoint(0);
            XYZ ep = line.GetEndPoint(1);

            //vecto đơn vị
            XYZ normalize = (ep - sp).Normalize();

            double distance = (normal.DotProduct(origin) - normal.DotProduct(sp)) / normal.DotProduct(normalize);

            XYZ intersectPoint = sp + distance * normalize;

            return intersectPoint;
        }

        public static List<Pipe> P_Get2pipemaxdistance(IList<Element> listPipes, XYZ pickPoint)
        {


            //tạo mặt phẳng
            Pipe pipe0 = listPipes[0] as Pipe;
            LocationCurve lc = pipe0.Location as LocationCurve;
            Line line0 = lc.Curve as Line;
            Plane plane = Plane.CreateByNormalAndOrigin(line0.Direction, pickPoint);

            var listXYZ = new List<XYZ>();
            foreach (Element element in listPipes)
            {
                Pipe pipe = element as Pipe;
                LocationCurve locationCurve = pipe.Location as LocationCurve;
                Line line = locationCurve.Curve as Line;

                Line newLine = CreateExtendLine(line, 500);
                XYZ intersectPoint = LineIntersectPlane(newLine, plane);
                listXYZ.Add(intersectPoint);
            }
            double maxdistance = P_GetMaxDistance(listPipes, pickPoint);
            XYZ points = null;
            XYZ pointe = null;
            foreach (XYZ point1 in listXYZ)
            {
                foreach (XYZ point2 in listXYZ)
                {
                    if (point1.DistanceTo(point2) == maxdistance)
                    {

                        points = point1;
                        pointe = point2;
                        break;
                    }


                }
            }
            List<Pipe> pipes = new List<Pipe>();
            Pipe pipengoaicung1 = P_GetPipe(listPipes, points);
            Pipe pipengoaicung2 = P_GetPipe(listPipes, pointe);
            pipes.Add(pipengoaicung1);

            return pipes;
        }
        public static List<Duct> D_Get2ductmaxdistance(IList<Element> listDucts, XYZ pickPoint)
        {


            //tạo mặt phẳng
            Duct duct0 = listDucts[0] as Duct;
            LocationCurve lc = duct0.Location as LocationCurve;
            Line line0 = lc.Curve as Line;
            Plane plane = Plane.CreateByNormalAndOrigin(line0.Direction, pickPoint);

            var listXYZ = new List<XYZ>();
            foreach (Element element in listDucts)
            {
                Duct duct = element as Duct;
                LocationCurve locationCurve = duct.Location as LocationCurve;
                Line line = locationCurve.Curve as Line;

                Line newLine = CreateExtendLine(line, 500);
                XYZ intersectPoint = LineIntersectPlane(newLine, plane);
                listXYZ.Add(intersectPoint);
            }
            double maxdistance = D_GetMaxDistance(listDucts, pickPoint);
            XYZ points = null;
            XYZ pointe = null;
            foreach (XYZ point1 in listXYZ)
            {
                foreach (XYZ point2 in listXYZ)
                {
                    if (point1.DistanceTo(point2) == maxdistance)
                    {

                        points = point1;
                        pointe = point2;
                        break;
                    }


                }
            }
            List<Duct> ducts = new List<Duct>();
            Duct ductngoaicung1 = D_GetDuct(listDucts, points);
            Duct ductngoaicung2 = D_GetDuct(listDucts, pointe);
            ducts.Add(ductngoaicung1);
            ducts.Add(ductngoaicung2);
            return ducts;
        }

        public static List<CableTray> C_Get2cabletraymaxdistance(IList<Element> listCabletrays, XYZ pickPoint)
        {


            //tạo mặt phẳng
            CableTray duct0 = listCabletrays[0] as CableTray;
            LocationCurve lc = duct0.Location as LocationCurve;
            Line line0 = lc.Curve as Line;
            Plane plane = Plane.CreateByNormalAndOrigin(line0.Direction, pickPoint);

            var listXYZ = new List<XYZ>();
            foreach (Element element in listCabletrays)
            {
                CableTray cb = element as CableTray;
                LocationCurve locationCurve = cb.Location as LocationCurve;
                Line line = locationCurve.Curve as Line;

                Line newLine = CreateExtendLine(line, 500);
                XYZ intersectPoint = LineIntersectPlane(newLine, plane);
                listXYZ.Add(intersectPoint);
            }
            double maxdistance = C_GetMaxDistance(listCabletrays, pickPoint);
            XYZ points = null;
            XYZ pointe = null;
            foreach (XYZ point1 in listXYZ)
            {
                foreach (XYZ point2 in listXYZ)
                {
                    if (point1.DistanceTo(point2) == maxdistance)
                    {

                        points = point1;
                        pointe = point2;
                        break;
                    }


                }
            }
            List<CableTray> cabletrays = new List<CableTray>();
            CableTray ductngoaicung1 = C_GetCableTray(listCabletrays, points);
            CableTray ductngoaicung2 = C_GetCableTray(listCabletrays, pointe);
            cabletrays.Add(ductngoaicung1);
            cabletrays.Add(ductngoaicung2);
            return cabletrays;
        }

        //mở rộng dường thẳng theo 2 hướng với 1 khoảng cách cố định
        public static Line CreateExtendLine(Line line, double distance)
        {
            XYZ sp = line.GetEndPoint(0);
            XYZ ep = line.GetEndPoint(1);

            //lấy vector normalize
            XYZ normalize = (ep - sp).Normalize();


            //tính toán tọa độ điểm mở rộng
            XYZ nsp = sp - normalize * distance;
            XYZ nep = ep + normalize * distance;

            Line newLine = Line.CreateBound(nsp, nep);
            return newLine;
        }
        public static Line P_GetNearestStartPP(Pipe pipe, XYZ pickPoint)
        {
            LocationCurve locationCurve = pipe.Location as LocationCurve;

            //Line pipeLine = locationCurve.Curve as Line;

            Line pipeLine = locationCurve.Curve as Line;
            XYZ p1 = pipeLine.GetEndPoint(0);
            XYZ p2 = pipeLine.GetEndPoint(1);
            XYZ p3 = pickPoint;
            double d1 = p3.DistanceTo(p1);
            double d2 = p3.DistanceTo(p2);
            if (d2 < d1)
            {
                Line pipeline1 = Line.CreateBound(p2, p1);
                return pipeline1;
            }

            return pipeLine;
        }

        public static Line D_GetNearestStartPP(Duct duct, XYZ pickPoint)
        {
            LocationCurve locationCurve = duct.Location as LocationCurve;

            // Lấy đường dẫn của Duct và ép kiểu sang Line
            Line ductLine = locationCurve.Curve as Line;

            // Lấy các điểm đầu và cuối của Line
            XYZ p1 = ductLine.GetEndPoint(0);
            XYZ p2 = ductLine.GetEndPoint(1);
            XYZ p3 = pickPoint;

            // Tính khoảng cách từ pickPoint đến p1 và p2
            double d1 = p3.DistanceTo(p1);
            double d2 = p3.DistanceTo(p2);

            // So sánh khoảng cách để xác định điểm gần hơn
            if (d2 < d1)
            {
                // Nếu p2 gần hơn, đảo chiều Line
                Line reversedLine = Line.CreateBound(p2, p1);
                return reversedLine;
            }

            // Nếu p1 gần hơn hoặc khoảng cách bằng nhau, trả về Line ban đầu
            return ductLine;
        }
        public static Line C_GetNearestStartPP(CableTray cabletray, XYZ pickPoint)
        {
            LocationCurve locationCurve = cabletray.Location as LocationCurve;

            // Lấy đường dẫn của Duct và ép kiểu sang Line
            Line cabletrayLine = locationCurve.Curve as Line;

            // Lấy các điểm đầu và cuối của Line
            XYZ p1 = cabletrayLine.GetEndPoint(0);
            XYZ p2 = cabletrayLine.GetEndPoint(1);
            XYZ p3 = pickPoint;

            // Tính khoảng cách từ pickPoint đến p1 và p2
            double d1 = p3.DistanceTo(p1);
            double d2 = p3.DistanceTo(p2);

            // So sánh khoảng cách để xác định điểm gần hơn
            if (d2 < d1)
            {
                // Nếu p2 gần hơn, đảo chiều Line
                Line reversedLine = Line.CreateBound(p2, p1);
                return reversedLine;
            }

            // Nếu p1 gần hơn hoặc khoảng cách bằng nhau, trả về Line ban đầu
            return cabletrayLine;
        }
        public static bool P_ChecksameSlope(Pipe pipe1, Pipe pipe2)
        {
            bool x = true;

            double slopeParam1 = Math.Round(pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE).AsDouble(), 2);
            double slopeParam2 = Math.Round(pipe2.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE).AsDouble(), 2);

            if (slopeParam1 == slopeParam2)
            {
                x = true;
            }

            else
            {
                x = false;
            }


            return x;

        }

        public static Line ExtendLine(Line originalLine, double thin1, double thin2)
        {
            // Lấy điểm đầu và điểm cuối của đoạn thẳng
            XYZ startPoint = originalLine.GetEndPoint(0);
            XYZ endPoint = originalLine.GetEndPoint(1);

            // Tính toán vector hướng
            XYZ direction1 = (-endPoint + startPoint).Normalize();
            XYZ direction2 = (endPoint - startPoint).Normalize();

            // Mở rộng điểm cuối
            XYZ newEndPoint1 = startPoint + direction1 * thin1;
            XYZ newEndPoint2 = endPoint + direction2 * thin2;

            // Tạo đoạn thẳng mới
            Line extendedLine = Line.CreateBound(newEndPoint1, newEndPoint2);

            return extendedLine;
        }

        public static XYZ GetMidpoint(Line extendLine)
        {
            XYZ point1 = extendLine.GetEndPoint(0);
            XYZ point2 = extendLine.GetEndPoint(1);
            double midX = (point1.X + point2.X) / 2;
            double midY = (point1.Y + point2.Y) / 2;
            double midZ = (point1.Z + point2.Z) / 2;

            return new XYZ(midX, midY, midZ);
        }
        public static FamilySymbol GetFamilySymbolCumU(Document doc, string familyName)
        {
            // Lọc tất cả các FamilySymbol trong tài liệu
            var listFamilySymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)) // Chỉ lấy các FamilySymbol (ElementType)
                .Cast<FamilySymbol>()
                .ToList();

            // Tìm FamilySymbol theo FamilyName và TypeName
            var symbol = listFamilySymbols
                .Find(x => x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase)
                          /* && x.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase)*/);

            return symbol; // Trả về FamilySymbol nếu tìm thấy

        }
        public static double P_GetMaxDistance(IList<Element> listPipes, XYZ pickPoint)
        {
            var listXYZ = P_GetListPoint(listPipes, pickPoint);
            var listDistance = new List<double>();

            for (int i = 0; i < listXYZ.Count; i++)
            {
                XYZ p1 = listXYZ[i];

                foreach (XYZ p2 in listXYZ)
                {
                    double distance = p2.DistanceTo(p1);
                    listDistance.Add(distance);
                }
            }

            return listDistance.Max();
        }
        public static double D_GetMaxDistance(IList<Element> listDucts, XYZ pickPoint)
        {
            var listXYZ = D_GetListPoint(listDucts, pickPoint);
            var listDistance = new List<double>();

            for (int i = 0; i < listXYZ.Count; i++)
            {
                XYZ p1 = listXYZ[i];

                foreach (XYZ p2 in listXYZ)
                {
                    double distance = p2.DistanceTo(p1);
                    listDistance.Add(distance);
                }
            }

            return listDistance.Max();
        }
        public static double C_GetMaxDistance(IList<Element> listCabletrays, XYZ pickPoint)
        {
            var listXYZ = C_GetListPoint(listCabletrays, pickPoint);
            var listDistance = new List<double>();

            for (int i = 0; i < listXYZ.Count; i++)
            {
                XYZ p1 = listXYZ[i];

                foreach (XYZ p2 in listXYZ)
                {
                    double distance = p2.DistanceTo(p1);
                    listDistance.Add(distance);
                }
            }

            return listDistance.Max();
        }
        public static List<XYZ> P_GetListPoint(IList<Element> listPipes, XYZ pickPoint)
        {
            //tạo mặt phẳng
            Pipe pipe0 = listPipes[0] as Pipe;
            LocationCurve lc = pipe0.Location as LocationCurve;
            Line line0 = lc.Curve as Line;
            Plane plane = Plane.CreateByNormalAndOrigin(line0.Direction, pickPoint);

            var listXYZ = new List<XYZ>();
            foreach (Element element in listPipes)
            {
                Pipe pipe = element as Pipe;
                LocationCurve locationCurve = pipe.Location as LocationCurve;
                Line line = locationCurve.Curve as Line;

                Line newLine = CreateExtendLine(line, 500);
                XYZ intersectPoint = LineIntersectPlane(newLine, plane);
                listXYZ.Add(intersectPoint);
            }
            return listXYZ;
        }
        public static List<XYZ> D_GetListPoint(IList<Element> listDucts, XYZ pickPoint)
        {
            //tạo mặt phẳng
            Duct duct0 = listDucts[0] as Duct;
            LocationCurve lc = duct0.Location as LocationCurve;
            Line line0 = lc.Curve as Line;
            Plane plane = Plane.CreateByNormalAndOrigin(line0.Direction, pickPoint);

            var listXYZ = new List<XYZ>();
            foreach (Element element in listDucts)
            {
                Duct duct = element as Duct;
                LocationCurve locationCurve = duct.Location as LocationCurve;
                Line line = locationCurve.Curve as Line;

                Line newLine = CreateExtendLine(line, 500);
                XYZ intersectPoint = LineIntersectPlane(newLine, plane);
                listXYZ.Add(intersectPoint);
            }
            return listXYZ;
        }
        public static List<XYZ> C_GetListPoint(IList<Element> listCabletrays, XYZ pickPoint)
        {
            //tạo mặt phẳng
            CableTray cabletray0 = listCabletrays[0] as CableTray;
            LocationCurve lc = cabletray0.Location as LocationCurve;
            Line line0 = lc.Curve as Line;
            Plane plane = Plane.CreateByNormalAndOrigin(line0.Direction, pickPoint);

            var listXYZ = new List<XYZ>();
            foreach (Element element in listCabletrays)
            {
                CableTray duct = element as CableTray;
                LocationCurve locationCurve = duct.Location as LocationCurve;
                Line line = locationCurve.Curve as Line;

                Line newLine = CreateExtendLine(line, 500);
                XYZ intersectPoint = LineIntersectPlane(newLine, plane);
                listXYZ.Add(intersectPoint);
            }
            return listXYZ;
        }

        //Ống nào đi qua điểm chỉ định
        public static Pipe P_GetPipe(IList<Element> listPipes, XYZ pickPoint)
        {
            foreach (Element element in listPipes)
            {
                Pipe pipe1 = element as Pipe;
                Line pipeLine = (pipe1.Location as LocationCurve).Curve as Line;

                double x = Math.Round(pickPoint.DistanceTo(pipeLine.GetEndPoint(0)) + pickPoint.DistanceTo(pipeLine.GetEndPoint(1)), 2);
                if (x == Math.Round(pipeLine.Length, 2))
                { return pipe1; }

            }
            return null;
        }
        public static Duct D_GetDuct(IList<Element> listDucts, XYZ pickPoint)
        {
            foreach (Element element in listDucts)
            {
                Duct duct1 = element as Duct;
                Line DuctLine = (duct1.Location as LocationCurve).Curve as Line;

                double x = Math.Round(pickPoint.DistanceTo(DuctLine.GetEndPoint(0)) + pickPoint.DistanceTo(DuctLine.GetEndPoint(1)), 2);
                if (x == Math.Round(DuctLine.Length, 2))
                { return duct1; }

            }
            return null;
        }
        public static CableTray C_GetCableTray(IList<Element> listCableTrays, XYZ pickPoint)
        {
            foreach (Element element in listCableTrays)
            {
                CableTray Cabletrays = element as CableTray;
                Line DuctLine = (Cabletrays.Location as LocationCurve).Curve as Line;

                double x = Math.Round(pickPoint.DistanceTo(DuctLine.GetEndPoint(0)) + pickPoint.DistanceTo(DuctLine.GetEndPoint(1)), 2);
                if (x == Math.Round(DuctLine.Length, 2))
                { return Cabletrays; }

            }
            return null;
        }

        public static BitmapImage Convert(Bitmap bimap)
        {
            MemoryStream memory = new MemoryStream();
            bimap.Save(memory, ImageFormat.Png);
            memory.Position = 0;
            BitmapImage bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = memory;
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
            bitmapImage.EndInit();
            return bitmapImage;
        }
        public static bool CheckCungTruc(Line line, Element element)
        {
            Line lineconnector = GetLineConnector(element);
            if (lineconnector != null)
            {
                double degree = GetAngle(line, lineconnector);
                if (degree == 0 || degree == 180)
                {
                    return true;
                }
            }
            return false;
        }
        public static Line GetLineConnector(Element element)
        {

            if (element is FamilyInstance familyInstance)
            {
                MEPModel mepModel = familyInstance.MEPModel;
                if (mepModel != null)
                {
                    ConnectorSet connectors = mepModel.ConnectorManager.Connectors;

                    // Chuyển đổi ConnectorSet thành mảng và sắp xếp theo tọa độ X trước
                    Connector[] connectorArray = connectors.Cast<Connector>()
                                                           .OrderBy(c => c.Origin.X) // Sắp xếp theo tọa độ X
                                                           .ThenBy(c => c.Origin.Y) // Sau đó sắp xếp theo Y
                                                           .ThenBy(c => c.Origin.Z) // Cuối cùng theo Z
                                                           .ToArray();

                    // Kiểm tra số lượng Connector trước khi truy cập
                    if (connectorArray.Length >= 2)
                    {
                        XYZ connector1 = connectorArray[0].Origin; // Connector đầu tiên
                        XYZ connector2 = connectorArray[1].Origin; // Connector thứ hai

                        Line lineconnector = Line.CreateBound(connector1, connector2);
                        return lineconnector;
                    }
                }
            }
            return null;
        }
        public static double GetAngle(Line line1, Line line2)
        {
            // Lấy vector chỉ phương của hai đường thẳng
            XYZ vector1 = line1.Direction; // Vector chỉ phương của đường thẳng 1
            XYZ vector2 = line2.Direction; // Vector chỉ phương của đường thẳng 2


            double angleInRadians = vector1.AngleTo(vector2);

            // Chuyển đổi sang độ (nếu cần)
            double angleInDegrees = angleInRadians * (180.0 / Math.PI);

            return angleInDegrees;
        }
        //Ngắt kết nối 1 dãy
        //public static void DisconnectElemnts(Document doc,IList<Element> elements)
        //{
        //    foreach (Element element in elements)
        //    {

        //            DisconnectTwoEnds(doc, element);


        //    }
        //}


        public static Hashtable DisconnectTwoEndsAndCache(Document doc, IList<Element> selectedElements)
        {
            Hashtable connectionMap = new Hashtable();
            HashSet<string> processedPairs = new HashSet<string>();
            HashSet<int> existingIds = new HashSet<int>();

            foreach (Element e in selectedElements)
                existingIds.Add(e.Id.IntegerValue);

            List<Element> toAdd = new List<Element>();

            foreach (Element element in selectedElements)
            {
                MEPModel mep = (element as FamilyInstance)?.MEPModel;
                if (mep == null) continue;

                foreach (Connector connector in mep.ConnectorManager.Connectors)
                {
                    foreach (Connector connected in connector.AllRefs)
                    {
                        Element other = connected.Owner;
                        if (other.Id == element.Id) continue;

                        string pairKey = $"{Math.Min(element.Id.IntegerValue, other.Id.IntegerValue)}_{Math.Max(element.Id.IntegerValue, other.Id.IntegerValue)}";
                        if (processedPairs.Contains(pairKey)) continue;

                        processedPairs.Add(pairKey);

                        string key = GetConnectorKey(connector);
                        if (!connectionMap.ContainsKey(key))
                            connectionMap[key] = connected;

                        // Thêm phần tử ngoài danh sách nếu chưa có
                        if (!existingIds.Contains(other.Id.IntegerValue))
                        {
                            toAdd.Add(other);
                            existingIds.Add(other.Id.IntegerValue);
                        }

                        try { connector.DisconnectFrom(connected); } catch { }
                    }
                }
            }

            // Thêm phần tử mới vào danh sách ban đầu
            foreach (Element e in toAdd)
                selectedElements.Add(e);

            return connectionMap;
        }

        private static string GetConnectorKey(Connector conn)
        {
            XYZ p = conn.Origin;
            return $"{conn.Owner.Id.IntegerValue}_{Math.Round(p.X, 6)}_{Math.Round(p.Y, 6)}_{Math.Round(p.Z, 6)}";
        }

        // Ngắt kết nối 2 đầu đối tượng
        public static void DisconnectTwoEnds(Document doc, Element element)
        {
           

            ConnectorSet connectors = GetConnectors(element);

           

                // Duyệt qua tất cả các Connector
                foreach (Connector connector in connectors)
                {
                    // Kiểm tra các kết nối hiện tại
                    ConnectorSet connectedConnectors = connector.AllRefs;
                    foreach (Connector connectedConnector in connectedConnectors)
                    {
                        if (connectedConnector.Owner.Id != element.Id)
                        {
                            // Ngắt kết nối
                            connector.DisconnectFrom(connectedConnector);
                            
                        }
                    }
                }

            

            //TaskDialog.Show("Success", "Đã ngắt kết nối hai đầu của đối tượng.");
        }


       
      
        public static ConnectorSet GetConnectors(Element element)
        {

            if (element == null) return null;

            // Trường hợp đối tượng là MEPCurve (Pipe, Duct, Conduit, etc.)
            if (element is MEPCurve mepCurve)
            {
                return mepCurve.ConnectorManager?.Connectors;
            }

            // Nếu đối tượng là FamilyInstance (ví dụ: Fittings)
            if (element is FamilyInstance familyInstance)
            {
                return familyInstance.MEPModel?.ConnectorManager?.Connectors;
            }

            return null; // Không lấy được connectors
        }
        public static Connector FindUnconnectedConnector(ConnectorSet connectors, XYZ targetPoint, double tolerance = 1e-3)
        {
            foreach (Connector connector in connectors)
            {
                if (!connector.IsConnected && connector.Origin.DistanceTo(targetPoint) < tolerance)
                {
                    return connector;
                }
            }

            return null;
        }



        public static void ConnectListElement(Document doc, IList<Element> elements)
        {
            foreach (Element element1 in elements)
            {
                foreach (Element element2 in elements)
                {
                    if (element1.Id == element2.Id)
                    { continue; }

                    // Lấy connectors của phần tử
                    ConnectorSet connectorsset1 = GetConnectors(element1);




                    foreach (Connector connector1 in connectorsset1)
                    {



                        // Kiểm tra nếu phần tử có connectors
                        ConnectorSet connectorsset2 = GetConnectors(element2);
                        if (connectorsset2 != null && connectorsset2.Size > 0)
                        {

                            foreach (Connector connector2 in connectorsset2)
                            {
                                try 
                                {
                                    if (connector1.Origin.DistanceTo(connector2.Origin) < 1 / 304.8)
                                    {


                                        connector1.ConnectTo(connector2);
                                    }
                                }

                                    catch
                                { }
                                   
                            }


                        }


                    }
                }
            }
        }



        public static void ConnectElementWithList(Document doc,Element ele, IList<Element> elements)
        { // Lấy connectors của phần tử
            ConnectorSet connectorssetele= GetConnectors(ele);
            foreach (Connector connectorele in connectorssetele)
            {
                foreach (Element element1 in elements)
                {
                    foreach (Element element2 in elements)
                    {
                        if (element1.Id == element2.Id)
                        { continue; }

                        // Lấy connectors của phần tử
                        ConnectorSet connectorsset1 = GetConnectors(element1);




                        foreach (Connector connector1 in connectorsset1)
                        {
                            try
                            {
                                if (connector1.Origin.DistanceTo(connectorele.Origin) < 1 / 304.8)
                                {


                                    connector1.ConnectTo(connectorele);
                                }
                            }
                            catch
                            { }

                            // Kiểm tra nếu phần tử có connectors
                            ConnectorSet connectorsset2 = GetConnectors(element2);
                            if (connectorsset2 != null && connectorsset2.Size > 0)
                            {

                                foreach (Connector connector2 in connectorsset2)
                                {
                                    try
                                    {
                                        if (connector1.Origin.DistanceTo(connector2.Origin) < 1 / 304.8)
                                        {


                                            connector1.ConnectTo(connector2);
                                        }
                                    }

                                    catch
                                    { }

                                }


                            }


                        }
                    }
                }
            }
        }

        public static Line GetDirectForAxis(Document doc, Element ele, IList<Element> elements)
        { 
            Pipe pipe=ele as Pipe;
            Line pipeLine = (pipe.Location as LocationCurve).Curve as Line;
            XYZ sp = (pipe.Location as LocationCurve).Curve.GetEndPoint(0);
            XYZ ep = (pipe.Location as LocationCurve).Curve.GetEndPoint(1);
            
            // Lấy connectors của phần tử
            ConnectorSet connectorssetele = GetConnectors(ele);
            foreach (Connector connectorele in connectorssetele)
            {
                foreach (Element element1 in elements)
                {
                   

                        // Lấy connectors của phần tử
                        ConnectorSet connectorsset1 = GetConnectors(element1);




                        foreach (Connector connector1 in connectorsset1)
                        {
                            
                                if (connector1.Origin.DistanceTo(connectorele.Origin) < 1 / 304.8)
                                {
                            if (sp.DistanceTo(connectorele.Origin) < 1 / 304.8)
                            {
                                pipeLine = Line.CreateBound(ep, sp);
                            } 
                                
                           

                                }
                            }
                            



                        }
                    }
            return pipeLine;
                }
        public static void DeleteElement(Document doc, IList<ElementId> ids, double length, out List<ElementId> result)
        {
            result = new List<ElementId>();

            foreach (ElementId id in ids)
            {
                LocationCurve locationCurve = null;

                Element element = doc.GetElement(id);

                if (element is Duct duct) locationCurve = duct.Location as LocationCurve;
                if (element is Pipe pipe) locationCurve = pipe.Location as LocationCurve;
                if (element is CableTray cableTray) locationCurve = cableTray.Location as LocationCurve;
                if (element is Conduit conduit) locationCurve = conduit.Location as LocationCurve;

                double curveLength = locationCurve.Curve.Length;
                double val = Math.Round(curveLength - length, 3);

                if (val == 0) doc.Delete(id);
                else result.Add(id);
            }
        }

        public static ElementId SliptCableTray(Document doc, ElementId element1id, XYZ sliptPoint)
        {
            Element element1 = doc.GetElement(element1id);
            CableTray trayOld = element1 as CableTray;

            LocationCurve locCurve = trayOld.Location as LocationCurve;
            if (locCurve == null)
            {
                TaskDialog.Show("Lỗi", "Không thể lấy đường Location của Cable Tray.");
                return null;
            }


            Line cableLine = locCurve.Curve as Line;
            XYZ startPoint = cableLine.GetEndPoint(0);
            XYZ endPoint = cableLine.GetEndPoint(1);
            // 🟢 BƯỚC 1: lấy connector ở endpoint cũ
            Connector endConnector = null;
            foreach (Connector c in trayOld.ConnectorManager.Connectors)
            {
                if (c.Origin.IsAlmostEqualTo(endPoint))
                {
                    endConnector = c;
                    break;
                }
            }
          
           
            double height = trayOld.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).AsDouble();
            double width = trayOld.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).AsDouble();

            CableTray cableTray1 = null;
            if (startPoint.DistanceTo(endPoint) == startPoint.DistanceTo(sliptPoint) + endPoint.DistanceTo(sliptPoint))
            {
                cableTray1 = CableTray.Create(doc, trayOld.GetTypeId(), sliptPoint, endPoint,
                    trayOld.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId());

                cableTray1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).Set(height);
                cableTray1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).Set(width);
                cableTray1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE)
                    .Set(trayOld.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString());

                (trayOld.Location as LocationCurve).Curve = Line.CreateBound(startPoint, sliptPoint);
            }

        

            if (endConnector == null)
            {
               
                return cableTray1?.Id;
            }

            // 🟢 BƯỚC 2: tìm connector trùng tọa độ trên cableTray1
            Connector newEndConnector = null;
            foreach (Connector c in cableTray1.ConnectorManager.Connectors)
            {
                // so sánh vị trí với endConnector
                if (c.Origin.IsAlmostEqualTo(endPoint))
                {
                    newEndConnector = c;
                    break;
                }
            }

            // 🟢 BƯỚC 3: tìm connector trên đối tượng khác từ endConnector
            Connector connectedConnector = null;
            Element connectedElement = null;
            if (endConnector.IsConnected)
            {
                foreach (Connector cref in endConnector.AllRefs)
                {
                    if (cref.Owner.Id != trayOld.Id) // bỏ qua chính trayOld
                    {
                        connectedConnector = cref;
                        connectedElement = cref.Owner;
                        break;
                    }
                }
            }

            // 👉 Debug thông tin trước khi ConnectTo
         
          

            // 🟢 BƯỚC 4: kết nối lại
            if (newEndConnector != null && connectedConnector != null)
            {
                try
                {
                    newEndConnector.ConnectTo(connectedConnector);
                }
                catch
                {
                    TaskDialog.Show("Reconnect", "Không thể ConnectTo.");
                }
            }

            return cableTray1?.Id;
        }













        public static ElementId SliptConduit(Document doc, ElementId element1id, XYZ sliptPoint)
        {
            Element element1 = doc.GetElement(element1id);
            Conduit conduitOld = element1 as Conduit;

            // Lấy đường tham chiếu (Location Curve) của Conduit
            LocationCurve locCurve = element1.Location as LocationCurve;
            if (locCurve == null)
            {
                TaskDialog.Show("Lỗi", "Không thể lấy đường Location của Conduit.");
                return null;
            }

            // Lấy điểm đầu và cuối của Conduit
            Line conduitLine = locCurve.Curve as Line;
            XYZ startPoint = conduitLine.GetEndPoint(0);
            XYZ endPoint = conduitLine.GetEndPoint(1);

            // 🟢 BƯỚC 1: lấy connector ở endpoint cũ
            Connector endConnector = null;
            foreach (Connector c in conduitOld.ConnectorManager.Connectors)
            {
                if (c.Origin.IsAlmostEqualTo(endPoint))
                {
                    endConnector = c;
                    break;
                }
            }

            // Tham số conduit
            double diameter = conduitOld.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).AsDouble();

            Conduit conduitNew = null;
            if (startPoint.DistanceTo(endPoint) == startPoint.DistanceTo(sliptPoint) + endPoint.DistanceTo(sliptPoint))
            {
                conduitNew = Conduit.Create(doc, conduitOld.GetTypeId(), sliptPoint, endPoint,
                    conduitOld.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId());

                conduitNew.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter);

                (conduitOld.Location as LocationCurve).Curve = Line.CreateBound(startPoint, sliptPoint);
            }

            if (endConnector == null)
            {
                return conduitNew?.Id;
            }
           
            // 🟢 BƯỚC 2: tìm connector trùng tọa độ trên conduitNew
            Connector newEndConnector = null;
            foreach (Connector c in conduitNew.ConnectorManager.Connectors)
            {
                if (c.Origin.IsAlmostEqualTo(endPoint))
                {
                    newEndConnector = c;
                    break;
                }
            }

            // 🟢 BƯỚC 3: tìm connector trên đối tượng khác từ endConnector
            Connector connectedConnector = null;
            Element connectedElement = null;
            if (endConnector.IsConnected)
            {
                foreach (Connector cref in endConnector.AllRefs)
                {
                    if (cref.Owner.Id != conduitOld.Id) // bỏ qua chính conduitOld
                    {
                        connectedConnector = cref;
                        connectedElement = cref.Owner;
                        break;
                    }
                }
            }

            // 🟢 BƯỚC 4: kết nối lại
            if (newEndConnector != null && connectedConnector != null)
            {
                try
                {
                    newEndConnector.ConnectTo(connectedConnector);
                }
                catch
                {
                    TaskDialog.Show("Reconnect", "Không thể ConnectTo.");
                }
            }

            return conduitNew?.Id;
        }



        public static void CreateElbowFiting(Document doc, Element e1, Element e2)
        {
            
            ConnectorSet cS1 = new ConnectorSet();
            ConnectorSet cS2 = new ConnectorSet();
            if (e1 is Duct duct1 && e2 is Duct duct2)
            {


                ConnectorManager cM1 = duct1.ConnectorManager;
                ConnectorManager cM2 = duct2.ConnectorManager;

                cS1 = cM1.Connectors;
                cS2 = cM2.Connectors;
            }
            if (e1 is Pipe pipe1 && e2 is Pipe pipe2)
            {


                ConnectorManager cM1 = pipe1.ConnectorManager;
                ConnectorManager cM2 = pipe2.ConnectorManager;

                cS1 = cM1.Connectors;
                cS2 = cM2.Connectors;
            }
            if (e1 is CableTray cableTray1 && e2 is CableTray cableTray2)
            {


                ConnectorManager cM1 = cableTray1.ConnectorManager;
                ConnectorManager cM2 = cableTray2.ConnectorManager;

                cS1 = cM1.Connectors;
                cS2 = cM2.Connectors;
            }
            if (e1 is Conduit conduit1 && e2 is Conduit conduit2)
            {


                ConnectorManager cM1 = conduit1.ConnectorManager;
                ConnectorManager cM2 = conduit2.ConnectorManager;

                cS1 = cM1.Connectors;
                cS2 = cM2.Connectors;
            }
            List<Connector> list = new List<Connector>();
            //tìm 2 connector trùng nhau
            foreach (Connector c1 in cS1)
            {
                foreach (Connector c2 in cS2)
                {
                    XYZ o1 = c1.Origin;
                    XYZ o2 = c2.Origin;
                    double kc = Math.Round(o1.DistanceTo(o2), 0);
                    if (kc == 0)
                    {
                        list.Add(c1);
                        list.Add(c2);
                        break;
                    }
                    
                }
            }

            try
            {
                doc.Create.NewElbowFitting(list[0], list[1]);
            }
            catch { }

        }



        public static void CreateElbowFitingsocua(Document doc, Element e1, Element e2)
        {

            ConnectorSet cS1 = new ConnectorSet();
            ConnectorSet cS2 = new ConnectorSet();
            if (e1 is Duct duct1 && e2 is Duct duct2)
            {


                ConnectorManager cM1 = duct1.ConnectorManager;
                ConnectorManager cM2 = duct2.ConnectorManager;

                cS1 = cM1.Connectors;
                cS2 = cM2.Connectors;
            }
            if (e1 is Pipe pipe1 && e2 is Pipe pipe2)
            {


                ConnectorManager cM1 = pipe1.ConnectorManager;
                ConnectorManager cM2 = pipe2.ConnectorManager;

                cS1 = cM1.Connectors;
                cS2 = cM2.Connectors;
            }
            if (e1 is CableTray cableTray1 && e2 is CableTray cableTray2)
            {


                ConnectorManager cM1 = cableTray1.ConnectorManager;
                ConnectorManager cM2 = cableTray2.ConnectorManager;

                cS1 = cM1.Connectors;
                cS2 = cM2.Connectors;
            }
            if (e1 is Conduit conduit1 && e2 is Conduit conduit2)
            {


                ConnectorManager cM1 = conduit1.ConnectorManager;
                ConnectorManager cM2 = conduit2.ConnectorManager;

                cS1 = cM1.Connectors;
                cS2 = cM2.Connectors;
            }
            List<Connector> list = new List<Connector>();
            //tìm 2 connector trùng nhau
            foreach (Connector c1 in cS1)
            {
                foreach (Connector c2 in cS2)
                {
                    XYZ o1 = c1.Origin;
                    XYZ o2 = c2.Origin;
                    double kc = Math.Round(o1.DistanceTo(o2), 3);
                    //MessageBox.Show(kc.ToString());
                    if (kc <1)
                    {
                        list.Add(c1);
                        list.Add(c2);
                        break;
                    }
                   
                }
            }

            try
            {
                doc.Create.NewElbowFitting(list[0], list[1]);
            }
            catch { }

        }

        public static bool ChekSolid(Element ele1, Element ele2)
        {
            ConnectorSet cS1 = new ConnectorSet();
            ConnectorSet cS2 = new ConnectorSet();
            Solid solid1 = GetMEPSolid(ele1);
            Solid solid2 = GetMEPSolid(ele2);

            if (solid1 == null || solid2 == null)
            {
                if (ele1 is Pipe pipe1 &&ele2 is Pipe pipe2)
                {
                    ConnectorManager cM1 = pipe1.ConnectorManager;
                    ConnectorManager cM2 = pipe2.ConnectorManager;

                    cS1 = cM1.Connectors;
                    cS2 = cM2.Connectors;
                    List<Connector> list = new List<Connector>();
                    foreach (Connector c1 in cS1)
                    {
                        foreach (Connector c2 in cS2)
                        {
                            XYZ o1 = c1.Origin;
                            XYZ o2 = c2.Origin;
                            double kc = Math.Round(o1.DistanceTo(o2), 3);
                          
                            if (kc == 0)
                            {
                                list.Add(c1);
                                list.Add(c2);
                                break;
                            }
                        }
                    }
                    if(list.Count == 2)
                    {
                        return true;
                    }
                   else { return false; }
                }
                if (ele1 is Conduit con1 &&ele2 is Conduit con2)
                {
                    ConnectorManager cM1 = con1.ConnectorManager;
                    ConnectorManager cM2 = con2.ConnectorManager;

                    cS1 = cM1.Connectors;
                    cS2 = cM2.Connectors;
                    List<Connector> list = new List<Connector>();
                    foreach (Connector c1 in cS1)
                    {
                        foreach (Connector c2 in cS2)
                        {
                            XYZ o1 = c1.Origin;
                            XYZ o2 = c2.Origin;
                            double kc = Math.Round(o1.DistanceTo(o2), 3);
                          
                            if (kc == 0)
                            {
                                list.Add(c1);
                                list.Add(c2);
                                break;
                            }
                        }
                    }
                    if(list.Count == 2)
                    {
                        return true;
                    }
                   else { return false; }
                }
                else { return false; }
               
            }
            else
            {
                Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Intersect);

                if (intersectSolid != null && intersectSolid.Volume > 0) return true;

                return false;
            }
          
        }
        public static Solid GetMEPSolid(Element element)
        {
            if (element is Duct||element is CableTray)
            {


                Options options = new Options()
                {
                    ComputeReferences = true,
                    DetailLevel = ViewDetailLevel.Medium,
                };

                GeometryElement geometryElement = element.get_Geometry(options);

                foreach (GeometryObject geoOb in geometryElement)
                {
                    if (geoOb is Solid solid) return solid;
                }

                return null;
            }
            else if (element is Pipe)
            {

               
                return null; // Nếu không có Solid hợp lệ
            }
            else
            {
                return null;    
            }
        }

        /* Check xem đối tượng sleeve đó có tồn tại trong môi trường chưa */ 
        public static bool Checkexist(Document doc, XYZ targetPoint )
        {


            // Lấy danh sách các FamilyInstance có FamilyName là "Sleeve"
            List<FamilyInstance> collector = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
               .Where(fi => fi.Symbol.Family.Name.ToLower().Contains("sleeve"))
                .ToList(); // Chuyển đổi thành List<FamilyInstance>

            // Kiểm tra xem có đối tượng nào trùng vị trí không
            bool exists = collector.Any(e =>
            {
                LocationPoint loc = e.Location as LocationPoint;
                if (loc != null)
                {
                    return loc.Point.DistanceTo(targetPoint) < 1; // Kiểm tra khoảng cách nhỏ để tránh sai số
                }
                return false;
            });

            return exists;



        }
        //public static IList<Element> Getfamilynameat2pointSPEP(Element duct)
        //{

        //    IList<string> familyname = new List<string>();
        //    List<Element> connectedElements = new List<Element>();

        //    // Lấy danh sách Connector của ống gió
        //    ConnectorSet connectors = duct.ConnectorManager.Connectors;
        //    LocationCurve locationCurve = duct.Location as LocationCurve;
        //    if (locationCurve == null)
        //    {
        //        return null;
        //    }

        //    // Xác định hai đầu ống gió
        //    XYZ startPoint = locationCurve.Curve.GetEndPoint(0);
        //    XYZ endPoint = locationCurve.Curve.GetEndPoint(1);

        //    // Danh sách tạm để chứa phần tử kết nối và khoảng cách đến hai đầu
        //    List<(Element element, double distanceToStart, double distanceToEnd)> elementsWithDistance = new List<(Element, double, double)>();

        //    foreach (Connector connector in connectors)
        //    {
        //        if (connector.IsConnected)
        //        {


        //            // Nếu connector không gần đầu hoặc cuối ống thì bỏ qua (continue)
        //            if (connector.Origin.DistanceTo(startPoint) > 0.001
        //                && connector.Origin.DistanceTo(endPoint) > 0.001) // Tức là không gần đầu hoặc đuôi ống
        //            {
        //                continue;
        //            }
        //            foreach (Connector refConnector in connector.AllRefs)
        //            {
        //                Element connectedElement = refConnector.Owner;

        //                // Đảm bảo phần tử kết nối hợp lệ và không trùng
        //                if (connectedElement != null && connectedElement.Id != duct.Id && !connectedElements.Contains(connectedElement))
        //                {
        //                    // Tính khoảng cách đến cả hai đầu ống
        //                    double distanceToStart = refConnector.Origin.DistanceTo(startPoint);
        //                    double distanceToEnd = refConnector.Origin.DistanceTo(endPoint);

        //                    // Thêm vào danh sách tạm
        //                    elementsWithDistance.Add((connectedElement, distanceToStart, distanceToEnd));
        //                }
        //            }
        //        }
        //    }

        //    // Sắp xếp danh sách sao cho phần tử gần đầu ống hơn đứng trước
        //    connectedElements = elementsWithDistance
        //        .OrderBy(e => e.distanceToStart)  // Sắp xếp theo khoảng cách đến đầu ống
        //        .ThenBy(e => e.distanceToEnd)  // Nếu cùng khoảng cách, sắp xếp tiếp theo khoảng cách đến đuôi
        //        .Select(e => e.element)
        //        .ToList();



        //    // Trả về danh sách chứa tên Family của các đối tượng
        //    return connectedElements;


        //}
        public static IList<Element> Getfamilynameat2pointSPEP(Element element)
        {
            IList<Element> connectedElements = new List<Element>();

            // Kiểm tra xem element có phải là đối tượng có ConnectorManager (Duct, Pipe, CableTray, FamilyInstance,...)
            ConnectorSet connectors = null;

            if (element is FamilyInstance familyInstance)
            {
                connectors = familyInstance.MEPModel.ConnectorManager.Connectors;
            }
            else if (element is Pipe pipe)
            {
                connectors = pipe.ConnectorManager.Connectors;
            }
            else if (element is Duct duct)
            {
                connectors = duct.ConnectorManager.Connectors;
            }
            else if (element is CableTray cableTray)
            {
                connectors = cableTray.ConnectorManager.Connectors;
            }
            else
            {
                TaskDialog.Show("Error", "Element is not a valid MEP component (Duct, Pipe, CableTray, or FamilyInstance).");
                return connectedElements;
            }

            if (connectors == null || connectors.Size == 0)
            {
                return connectedElements; // Nếu không có connectors, trả về danh sách rỗng
            }

            // Kiểm tra vị trí của element (LocationCurve hoặc LocationPoint)
            LocationCurve locationCurve = element.Location as LocationCurve;
            LocationPoint locationPoint = element.Location as LocationPoint;

            if (locationCurve != null)
            {
                // Đối với Duct, Pipe, CableTray (có LocationCurve)
                XYZ startPoint = locationCurve.Curve.GetEndPoint(0);
                XYZ endPoint = locationCurve.Curve.GetEndPoint(1);

                // Danh sách tạm để chứa phần tử kết nối và khoảng cách đến hai đầu
                List<(Element connectedElement, double distanceToStart, double distanceToEnd)> elementsWithDistance = new List<(Element, double, double)>();

                foreach (Connector connector in connectors)
                {
                    if (connector.IsConnected)
                    {
                        // Nếu connector không gần đầu hoặc cuối ống thì bỏ qua
                        if (connector.Origin.DistanceTo(startPoint) > 0.001 && connector.Origin.DistanceTo(endPoint) > 0.001)
                        {
                            continue;
                        }

                        foreach (Connector refConnector in connector.AllRefs)
                        {
                            Element connectedElement = refConnector.Owner;
                            if (connectedElement != null && connectedElement.Id != element.Id && !connectedElements.Contains(connectedElement))
                            {
                                double distanceToStart = refConnector.Origin.DistanceTo(startPoint);
                                double distanceToEnd = refConnector.Origin.DistanceTo(endPoint);
                                elementsWithDistance.Add((connectedElement, distanceToStart, distanceToEnd));
                            }
                        }
                    }
                }

                connectedElements = elementsWithDistance
                    .OrderBy(e => e.distanceToStart)
                    .ThenBy(e => e.distanceToEnd)
                    .Select(e => e.connectedElement)
                    .ToList();
            }
            else if (locationPoint != null)
            {
                // Đối với FamilyInstance (có LocationPoint)
                XYZ point = locationPoint.Point;

                foreach (Connector connector in connectors)
                {
                    if (connector.IsConnected)
                    {
                        foreach (Connector refConnector in connector.AllRefs)
                        {
                            Element connectedElement = refConnector.Owner;

                            // Đảm bảo phần tử kết nối hợp lệ và không trùng
                            if (connectedElement != null && connectedElement.Id != element.Id && !connectedElements.Contains(connectedElement))
                            {
                                // Tính khoảng cách đến vị trí của FamilyInstance
                                double distanceToPoint = refConnector.Origin.DistanceTo(point);

                                // Thêm vào danh sách kết nối
                                connectedElements.Add(connectedElement);
                            }
                        }
                    }
                }
            }
            else
            {
                TaskDialog.Show("Error", "Element does not have a valid location (neither LocationCurve nor LocationPoint).");
            }

            return connectedElements;


        }
        public static  void SlpitDuctFormStartPoint(Document doc, Duct originalDuct, double distance, double lastSegmentLength)
        {



            LocationCurve locationCurve = originalDuct.Location as LocationCurve;
            Line locationLine = locationCurve.Curve as Line;
            Line lineRemain = locationLine;
            List<ElementId> listId = new List<ElementId>();

            double totalLength = locationLine.Length;
            double num = Math.Floor((totalLength - lastSegmentLength) / distance); // Tính số đoạn cắt có chiều dài đều (số đoạn chia đều)


            // Cắt các đoạn có chiều dài đều (num-1 đoạn)
            for (int i = 0; i < num; i++)
            {
                try
                {
                    XYZ p = FindPointOnLineFromStartPoint(lineRemain, distance); // Tìm điểm cắt
                    ElementId id = MechanicalUtils.BreakCurve(doc, originalDuct.Id, p);
                    CreateUnionFiting(doc, originalDuct, doc.GetElement(id) as Duct);
                }
                catch
                {
                    // Xử lý ngoại lệ nếu cần
                }
            }

            // Tạo điểm cắt cho đoạn cuối sao cho chiều dài của nó là lastSegmentLength
            XYZ finalPoint = FindPointOnLineFromEndPoint(lineRemain, lastSegmentLength); // Tính điểm cắt cho đoạn cuối
            ElementId finalId = MechanicalUtils.BreakCurve(doc, originalDuct.Id, finalPoint);
            listId.Add(finalId); // Thêm phần tử cuối vào danh sách

            // Tạo các fitting giữa các đoạn
            CreateUnionFiting(doc, originalDuct, doc.GetElement(finalId) as Duct);
        }
        public static  void CreateUnionFiting(Document doc, Duct duct1, Duct duct2)
        {
            ConnectorManager Cm1 = duct1.ConnectorManager;
            ConnectorManager Cm2 = duct2.ConnectorManager;

            ConnectorSet cs1 = Cm1.Connectors;
            ConnectorSet cs2 = Cm2.Connectors;
            List<Connector> list = new List<Connector>();
            foreach (Connector c1 in cs1)
            {
                foreach (Connector c2 in cs2)
                {
                    XYZ o1 = c1.Origin;
                    XYZ o2 = c2.Origin;
                    double khoangcach = Math.Round(o1.DistanceTo(o2), 3);


                    if (khoangcach == 0)
                    {
                        list.Add(c1);
                        list.Add(c2);
                        break;
                    }




                }
            }
            doc.Create.NewUnionFitting(list[0], list[1]);

        }
        public static void SlpitDuctFormEndPoint(Document doc, Duct originDuct, double distance, double lastSegmentLength)
        {


            LocationCurve locationCurve = originDuct.Location as LocationCurve;
            Line locationLine = locationCurve.Curve as Line;
            double number = Math.Round(locationLine.Length / distance, 0);
            int total = int.Parse(number.ToString());

            Line line = locationLine;
            List<ElementId> listId = new List<ElementId>();
            listId.Add(originDuct.Id);

            for (int i = 0; i < total; i++)
            {
                try
                {
                    //ngắt ống gió
                    XYZ p = FindPointOnLineFromEndPoint(line, distance);
                    ElementId id = MechanicalUtils.BreakCurve(doc, originDuct.Id, p);
                    listId.Add(id);

                    //reset data
                    originDuct = doc.GetElement(id) as Duct;
                    LocationCurve lc = originDuct.Location as LocationCurve;
                    line = lc.Curve as Line;
                }
                catch { }
            }

            CreateUnionsFiting(doc, listId);
        }
        public static void CreateUnionsFiting(Document doc, List<ElementId> listIds)
        {
            Duct duct1 = doc.GetElement(listIds[0]) as Duct;

            for (int i = 1; i < listIds.Count; i++)
            {
                Duct duct_i = doc.GetElement(listIds[i]) as Duct;
                CreateUnionFiting(doc, duct1, duct_i);
                duct1 = duct_i;




            }

        }
        public static void ApplyInterferenceHighlight(Document doc, UIDocument uiDoc, ElementId elemId)
        {
            // Kiểm tra element
            Element element = doc.GetElement(elemId);
            if (element == null)
            {
                TaskDialog.Show("Error", "Element không tồn tại!");
                return;
            }

            // Highlight trên toàn bộ môi trường bằng Selection
            try
            {
                List<ElementId> selectedIds = new List<ElementId> { elemId };
                uiDoc.Selection.SetElementIds(selectedIds); // Highlight trên toàn bộ môi trường
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Không thể chọn element1: {ex.Message}");
                return;
            }
            //a

            // Áp dụng OverrideGraphicSettings cho các view chính (tùy chọn)
            // Lấy tất cả các view 3D và FloorPlan
            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && (v.ViewType == ViewType.ThreeD || v.ViewType == ViewType.FloorPlan))
                .ToList();

            // Kiểm tra xem element có hiển thị trong ít nhất một view không
            bool isVisibleInAnyView = false;
            foreach (View view in views)
            {
                BoundingBoxXYZ bbox = element.get_BoundingBox(view);
                if (bbox != null && bbox.Enabled)
                {
                    isVisibleInAnyView = true;
                    break;
                }
            }

            if (!isVisibleInAnyView)
            {
                TaskDialog.Show("Warning", $"Element '{element.Name}' (ID: {elemId}) không hiển thị trong bất kỳ view 3D hoặc FloorPlan nào. Kiểm tra Visibility/Graphics hoặc mặt cắt.");
                // Vẫn tiếp tục vì đã có Selection highlight
            }

            // Tạo OverrideGraphicSettings
            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            Autodesk.Revit.DB.Color highlightColor = new Color(255, 150, 0); // Màu cam

            // Lấy Solid Fill Pattern
            FillPatternElement solidFill = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(q => q.GetFillPattern().IsSolidFill);

            if (solidFill == null)
            {
                TaskDialog.Show("Warning", "Không tìm thấy Solid Fill Pattern trong tài liệu! Highlight tùy chỉnh sẽ không được áp dụng.");
                // Vẫn tiếp tục vì đã có Selection highlight
            }
            else
            {
                // Thiết lập đồ họa cho mặt bằng
                ogs.SetSurfaceForegroundPatternId(solidFill.Id);
                ogs.SetSurfaceForegroundPatternColor(highlightColor);
                ogs.SetSurfaceBackgroundPatternId(solidFill.Id);
                ogs.SetSurfaceBackgroundPatternColor(highlightColor);
                ogs.SetCutForegroundPatternId(solidFill.Id); // Quan trọng cho mặt bằng
                ogs.SetCutForegroundPatternColor(highlightColor);
                ogs.SetCutBackgroundPatternId(solidFill.Id);
                ogs.SetCutBackgroundPatternColor(highlightColor);

                // Thiết lập đường bao
                ogs.SetProjectionLineColor(highlightColor);
                ogs.SetProjectionLineWeight(5);
                ogs.SetCutLineColor(highlightColor);
                ogs.SetCutLineWeight(5);

                // Áp dụng override cho từng view
                foreach (View view in views)
                {
                    BoundingBoxXYZ bbox = element.get_BoundingBox(view);
                    if (bbox == null || !bbox.Enabled)
                    {
                        continue; // Bỏ qua view mà element không hiển thị
                    }

                    try
                    {
                        view.SetElementOverrides(elemId, ogs);
                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("Warning", $"Không thể áp dụng highlight trong view '{view.Name}': {ex.Message}");
                    }
                }
            }

            // Làm mới view hiện tại
            uiDoc.RefreshActiveView();
        }
        public static View3D OpenOrCreate3DView( UIDocument uidoc, Document doc)
        {


            // Lấy tất cả view 3D hiện có
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(View3D));
            View3D view3D = null;

            foreach (Element element in collector)
            {
                view3D = element as View3D;
                if (view3D != null && !view3D.IsTemplate)
                {
                    break;
                }
            }

            if (view3D != null)
            {
                // Kích hoạt view 3D
                uidoc.ActiveView = view3D;
            }
            else
            {
                TaskDialog.Show("Error", "View 3D không tồn tại, cần tạo mới.");
            }

            return view3D;

        }



        public static void CreateSectionBoxForElements(UIDocument uiDoc, Document doc,  ElementId e1, ElementId e2 )
        {
            // Kiểm tra view hiện tại có phải là View3D hay không
            //View currentView = uiDoc.ActiveView;
            //if (currentView.ViewType != ViewType.ThreeD)
            //{
            //    TaskDialog.Show("Error", "View hiện tại không phải là View3D! Vui lòng chuyển sang view 3D trước.");
            //    return;
            //}

            View3D view3D = CT.OpenOrCreate3DView(uiDoc,doc);

            // Tạo SectionBox bao quanh các elements
            BoundingBoxXYZ combinedBox = null;

            // Lấy bounding box của element đầu tiên
            if (e1 != null)
            {
                Element elem1 = doc.GetElement(e1);
                if (elem1 != null)
                {
                    BoundingBoxXYZ bbox1 = elem1.get_BoundingBox(null); // Lấy trong không gian mô hình
                    if (bbox1 != null)
                    {
                        combinedBox = new BoundingBoxXYZ
                        {
                            Min = bbox1.Min,
                            Max = bbox1.Max
                        };
                    }
                    else
                    {
                        TaskDialog.Show("Warning", $"Không thể lấy BoundingBox của element1 với ID {e1}.");
                    }
                }
                else
                {
                    TaskDialog.Show("Warning", $"Element với ID {e1} không tồn tại!");
                }
            }
            else
            {
                TaskDialog.Show("Error", "ElementId đầu tiên không được để trống!");
                return;
            }

            // Kết hợp với bounding box của element thứ hai (nếu có)
            if (e2 != null)
            {
                Element elem2 = doc.GetElement(e2);
                if (elem2 != null)
                {
                    BoundingBoxXYZ bbox2 = elem2.get_BoundingBox(null);
                    if (bbox2 != null)
                    {
                        if (combinedBox == null)
                        {
                            combinedBox = new BoundingBoxXYZ
                            {
                                Min = bbox2.Min,
                                Max = bbox2.Max
                            };
                        }
                        else
                        {
                            // Kết hợp hai bounding box
                            combinedBox.Min = new XYZ(
                                Math.Min(combinedBox.Min.X, bbox2.Min.X),
                                Math.Min(combinedBox.Min.Y, bbox2.Min.Y),
                                Math.Min(combinedBox.Min.Z, bbox2.Min.Z));
                            combinedBox.Max = new XYZ(
                                Math.Max(combinedBox.Max.X, bbox2.Max.X),
                                Math.Max(combinedBox.Max.Y, bbox2.Max.Y),
                                Math.Max(combinedBox.Max.Z, bbox2.Max.Z));
                        }
                    }
                    else
                    {
                        TaskDialog.Show("Warning", $"Không thể lấy BoundingBox của element1 với ID {e2}.");
                    }
                }
                else
                {
                    TaskDialog.Show("Warning", $"Element với ID {e2} không tồn tại!");
                }
            }

            // Nếu có bounding box hợp lệ, áp dụng SectionBox
            if (combinedBox != null)
            {
                // Mở rộng bounding box để tạo SectionBox
                double padding = 5.0; // Khoảng cách mở rộng (feet)
                combinedBox.Min = new XYZ(
                    combinedBox.Min.X - padding,
                    combinedBox.Min.Y - padding,
                    combinedBox.Min.Z - padding);
                combinedBox.Max = new XYZ(
                    combinedBox.Max.X + padding,
                    combinedBox.Max.Y + padding,
                    combinedBox.Max.Z + padding);

                // Áp dụng SectionBox cho view 3D
                using (Transaction t = new Transaction(doc, "Set Section Box"))
                {
                    t.Start();
                    view3D.SetSectionBox(combinedBox);
                    view3D.GetSectionBox().Enabled = true; // Đảm bảo SectionBox được bật
                    t.Commit();
                }

                //TaskDialog.Show("Success", "Đã tạo SectionBox thành công!");
            }
            else
            {
                TaskDialog.Show("Error", "Không thể tạo SectionBox vì không có BoundingBox hợp lệ.");
                return;
            }

            // Làm mới view
            uiDoc.RefreshActiveView();
        }
        public static void ClearHighlight(UIDocument uiDoc, Document doc )
        {
            // Lấy tất cả các element trong tài liệu
            var elements = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToElements();

            if (!elements.Any()) return;

            // Lấy tất cả các view (3D và FloorPlan) không phải template
            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                //a
                .Cast<View>()
                .Where(v => !v.IsTemplate && (v.ViewType == ViewType.ThreeD || v.ViewType == ViewType.FloorPlan))
                .ToList();

            if (!views.Any()) return;

            // Xóa override đồ họa của tất cả các element trong các view
            using (Transaction t = new Transaction(doc, "Clear All Highlights"))
            {
                t.Start();
                foreach (View view in views)
                {
                    foreach (Element element in elements)
                    {
                        try
                        {
                            // Xóa override đồ họa (trả về mặc định)
                            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                            view.SetElementOverrides(element.Id, ogs);
                        }
                        catch { }
                    }
                }
                t.Commit();
            }

            // Xóa selection
            try
            {
                uiDoc.Selection.SetElementIds(new List<ElementId>()); // Xóa selection
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Không thể xóa selection: {ex.Message}");
            }

            // Làm mới view
            uiDoc.RefreshActiveView();
        }
        public static Hashtable GetFamilyInstancesInRooms(Document doc)
        {
            // Tạo Hashtable để lưu kết quả: Key là chuỗi Room, Value là danh sách chuỗi chứa category, ID cặp FamilyInstance và khoảng cách
            Hashtable roomFamilyInstances = new Hashtable();

            // Lấy tất cả các Room trong document
            FilteredElementCollector roomCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            // Tạo bộ lọc danh mục cho Mechanical Equipment và Air Terminal
            ElementCategoryFilter mechEquipFilter = new ElementCategoryFilter(BuiltInCategory.OST_MechanicalEquipment);
            ElementCategoryFilter airTerminalFilter = new ElementCategoryFilter(BuiltInCategory.OST_DuctTerminal);
            LogicalOrFilter categoryFilter = new LogicalOrFilter(mechEquipFilter, airTerminalFilter);

            // Lấy tất cả FamilyInstance thuộc Mechanical Equipment hoặc Air Terminal
            FilteredElementCollector instanceCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .WherePasses(categoryFilter);

            // Duyệt qua từng Room
            foreach (Room room in roomCollector)
            {
                try
                {
                    // Kiểm tra xem Room có hợp lệ không
                    if (room == null || room.Location == null) continue;

                    // Tạo danh sách để lưu FamilyInstance trong Room này
                    List<FamilyInstance> instancesInRoom = new List<FamilyInstance>();
                    List<string> distanceInfoList = new List<string>();
                    IList<double> KCdentuong = new List<double>();

                    // Kiểm tra từng FamilyInstance
                    foreach (FamilyInstance fi in instanceCollector)
                    {
                        try
                        {
                            // Lấy vị trí của FamilyInstance
                            Location location = fi.Location;
                            if (location is LocationPoint locationPoint)
                            {
                                XYZ origin = locationPoint.Point;

                                // Kiểm tra xem điểm có nằm trong Room không
                                if (room.IsPointInRoom(origin))
                                {
                                    instancesInRoom.Add(fi);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            //System.Diagnostics.Debug.WriteLine($"Lỗi khi xử lý FamilyInstance {fi.Id}: {ex.Message}");
                            continue;
                        }
                    }

                    // Tính khoảng cách giữa các cặp FamilyInstance trong cùng Room
                    for (int i = 0; i < instancesInRoom.Count; i++)
                    {
                        FamilyInstance fi1 = instancesInRoom[i];
                        Location location1 = fi1.Location;
                        if (!(location1 is LocationPoint locationPoint1)) continue;
                        XYZ origin1 = locationPoint1.Point;

                        // Lấy Family Symbol và Family của fi1
                        Family family1 = fi1.Symbol.Family;
                        Document familyDoc1 = doc; // Truy cập trực tiếp trong document chính

                        // Lấy Reference Plane từ FamilyInstance
                        ReferencePlane targetRefPlaneLeftRight = new FilteredElementCollector(familyDoc1)
                            .OfClass(typeof(ReferencePlane))
                            .Cast<ReferencePlane>()
                            .FirstOrDefault(rp => rp.Name == "Center (Left/Right)");
                        ReferencePlane targetRefPlaneFrontBack = new FilteredElementCollector(familyDoc1)
                            .OfClass(typeof(ReferencePlane))
                            .Cast<ReferencePlane>()
                            .FirstOrDefault(rp => rp.Name == "Center (Front/Back)");

                        XYZ normalLeftRight = targetRefPlaneLeftRight?.Normal ?? XYZ.BasisX; // Mặc định nếu null
                        XYZ normalFrontBack = targetRefPlaneFrontBack?.Normal ?? XYZ.BasisY; // Mặc định nếu null
                        SpatialElementBoundaryOptions boundaryOptions = new SpatialElementBoundaryOptions
                        {
                            SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish // Lấy biên dạng tại vị trí hoàn thiện
                        };

                        IList<IList<BoundarySegment>> boundarySegments = room.GetBoundarySegments(boundaryOptions);
                        foreach (IList<BoundarySegment> segmentLoop in boundarySegments)
                        {
                            foreach (BoundarySegment segment in segmentLoop)
                            {
                                Curve curve = segment.GetCurve();
                                if (curve is Line line)
                                {
                                    XYZ startpoint = line.GetEndPoint(0);
                                    double distance = Math.Min(
                                        Math.Round(Math.Abs((startpoint - origin1).DotProduct(normalLeftRight) * 304.8), 0),
                                        Math.Round(Math.Abs((startpoint - origin1).DotProduct(normalFrontBack) * 304.8), 0));
                                    KCdentuong.Add(distance);
                                }
                            }
                        }
                        if (KCdentuong.Any() && KCdentuong.Min() < 3000)
                        {
                            string categoryName = fi1.Category?.Name ?? "Unknown Category";
                            string distanceInfo = $"{categoryName}: Id {fi1.Id} - Khoảng cách đến tường gần nhất: {KCdentuong.Min()}mm";
                            distanceInfoList.Add(distanceInfo);
                        }

                        for (int j = i + 1; j < instancesInRoom.Count; j++)
                        {
                            FamilyInstance fi2 = instancesInRoom[j];
                            Location location2 = fi2.Location;
                            if (!(location2 is LocationPoint locationPoint2)) continue;
                            XYZ origin2 = locationPoint2.Point;

                            // Tính khoảng cách
                            XYZ vectorDiff = origin2 - origin1; // Vector hiệu O_2 - O_1
                            double distance = Math.Max(
                                Math.Round(Math.Abs(vectorDiff.DotProduct(normalLeftRight) * 304.8), 0),
                                Math.Round(Math.Abs(vectorDiff.DotProduct(normalFrontBack) * 304.8), 0)
                            );

                            if (distance < 7000)
                            {
                                // Tạo chuỗi ghép: Category - fi1.Id - fi2.Id - Distance
                                string categoryName = fi1.Category?.Name ?? "Unknown Category";
                                string distanceInfo = $"{categoryName}: Id {fi1.Id} - {fi2.Category?.Name ?? "Unknown Category"}: Id {fi2.Id} - Distance: {distance}mm";
                                distanceInfoList.Add(distanceInfo);
                            }
                        }
                    }

                    // Thêm vào Hashtable với key là ID của Room
                    if (distanceInfoList.Count > 0)
                    {
                        string categoryName = room.Category != null ? room.Category.Name : "Unknown Category";
                        string key = $"{categoryName} - {room.Id}";
                        roomFamilyInstances.Add(key, distanceInfoList);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Lỗi khi xử lý Room {room.Id}: {ex.Message}");
                    continue;
                }
            }

            return roomFamilyInstances;
        }

        public static IList<Element> Align_CurrentSelection(
    Document doc,
    UIDocument uidoc,
    ICollection<ElementId> _SelectedIds,
    IList<Element> pickedElements,
    ISelectionFilter selFilter)
        {
            foreach (ElementId selectedId in (IEnumerable<ElementId>)_SelectedIds)
            {
                try
                {
                    Element element = doc.GetElement(selectedId);
                    if (element.Category != null)
                    {
                        if (element.Location != null)
                        {
                            if (element.Category.IsTagCategory)
                            {
                                if (element.Category.Id.IntegerValue != -2000280)
                                {
                                    if (element.Category.Id.IntegerValue != -2000480)
                                    {
                                        if (element.Category.Id.IntegerValue != -2005020)
                                        {
                                            if (element.Category.Id.IntegerValue != 2000485)
                                                pickedElements.Add(element);
                                        }
                                    }
                                }
                            }
                            else if (element.Category.Id.IntegerValue == -2000300)
                                pickedElements.Add(element);
                            else if (element.Location.GetType() == typeof(LocationPoint))
                                pickedElements.Add(element);
                            else if (element.Location.GetType() == typeof(LocationCurve))
                                pickedElements.Add(element);
                        }
                    }
                }
                catch (Exception ex)
                {
                    int num = (int)System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                }
            }
            if (pickedElements.Count <= 0)
            {
                try
                {
                    pickedElements = uidoc.Selection.PickElementsByRectangle(selFilter, "Select by rectangle");
                }
                catch
                {
                }
            }
            return pickedElements;
        }
        public static Plane Set_WorkPlane(Document doc)
        {
            View view = doc.ActiveView;

            Plane byNormalAndOrigin = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin);
            SketchPlane sketchPlane = SketchPlane.Create(doc, byNormalAndOrigin);

            if (view.ViewType != ViewType.FloorPlan &&
                view.ViewType != ViewType.CeilingPlan &&
                view.ViewType != ViewType.Section)
            {
                return null; // hoặc throw exception nếu bạn muốn bắt buộc 2D
            }

            return byNormalAndOrigin;
        }
        public static double SignedDistanceTo(Plane plane, XYZ p)
        {
            XYZ vec = p - plane.Origin; // dùng toán tử trừ trực tiếp
            return plane.Normal.DotProduct(vec); // khoảng cách có dấu
        }
        public static XYZ ProjectOnto(Plane plane, XYZ p)
        {
            double d = SignedDistanceTo(plane, p);
            return p - d * plane.Normal;
        }



    }
}

        

    


