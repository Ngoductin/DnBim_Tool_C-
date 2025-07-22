using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using DnBim_Tool;
using System.Windows;
using Autodesk.Revit.DB.Mechanical;

namespace DnBim_Tool
{
    public class DuctUltis
    {
        public static void D_ThanhUGiaDo_Auto(UIDocument uidoc, Document doc, Plane infinitePlane, double distanceStart, double distanceOffset, string familyName, string typeName)
        {

            
            try
            {
               

                while (true)
                {
                    IList<Element> selectedRefs = uidoc.Selection.PickElementsByRectangle(new DuctFilter(), "Chọn các ống gió cần đặt support");

                    if (selectedRefs.Count > 1)
                    {
                        
                        Reference reference1 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new DuctFilter(), "Hãy chọn ống làm chuẩn");

                        XYZ startPoint = reference1.GlobalPoint;


                        Duct ductstd = doc.GetElement(reference1) as Duct;
                        LocationCurve locationCurvestd = ductstd.Location as LocationCurve;
                        Line ductLinestd = CT.FindDirectionOfElement(reference1, locationCurvestd);


                        List<Duct> ducts = CT.D_Get2ductmaxdistance(selectedRefs, startPoint);
                        Duct duct1 = ducts[0];
                        Duct duct2 = ducts[1];
                        //MessageBox.Show("Tôi bị ngu");

                        // Ống thứ 1
                        LocationCurve locationduct1curve = duct1.Location as LocationCurve;
                        Line lineDuct1 = locationduct1curve.Curve as Line;
                        lineDuct1 = CT.D_GetNearestStartPP(duct1, startPoint);
                        double lowpoint1 = Math.Round((duct1.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble() * 0.5 + duct1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);

                        XYZ sp1 = lineDuct1.GetEndPoint(0);
                        XYZ ep1 = lineDuct1.GetEndPoint(1);

                        XYZ Lsp1 = new XYZ(sp1.X, sp1.Y, (sp1.Z - lowpoint1));
                        XYZ Lep1 = new XYZ(ep1.X, ep1.Y, (ep1.Z - lowpoint1));
                        Line newline1 = Line.CreateBound(Lsp1, Lep1);


                        // Ống thứ 2

                        LocationCurve locationpipe2curve = duct2.Location as LocationCurve;
                        Line linePipe2 = locationpipe2curve.Curve as Line;
                        linePipe2 = CT.D_GetNearestStartPP(duct2, startPoint);
                        double lowpoint2 = Math.Round((duct2.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble() * 0.5 + duct2.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        XYZ sp2 = linePipe2.GetEndPoint(0);
                        XYZ ep2 = linePipe2.GetEndPoint(1);

                        XYZ Lsp2 = new XYZ(sp2.X, sp2.Y, (sp2.Z - lowpoint2));
                        XYZ Lep2 = new XYZ(ep2.X, ep2.Y, (ep2.Z - lowpoint2));



                        if ((Math.Round(Lsp1.Z, 2)) == (Math.Round(Lsp2.Z, 2)))
                        {
                            Line newline2 = Line.CreateBound(Lsp2, Lep2);
                            XYZ directionsp = newline2.Direction;
                            Plane plane1 = Plane.CreateByNormalAndOrigin(directionsp, Lsp1);
                            XYZ p1 = CT.LineIntersectPlane(newline2, plane1);



                            Line distacelineSP = Line.CreateBound(Lsp1, p1);


                            double space1 = Math.Round((duct1.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble() * 0.5 + duct1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3); ;
                            double space2 = Math.Round((duct2.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble() * 0.5 + duct2.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3); ;

                            //MessageBox.Show($"{space1*304.8}\n {space2}");
                            Line extendLineSP = CT.ExtendLine(distacelineSP, space1, space2);
                            double distance = Math.Round((extendLineSP.Length), 2);

                           
                            //Line ong 1
                            XYZ s1 = extendLineSP.GetEndPoint(0);
                           
                            //Line ong 2
                            XYZ s2 = extendLineSP.GetEndPoint(1);

                            // Tính toán vector dịch chuyển
                            XYZ offsetVector = ductLinestd.Length * ductLinestd.Direction.Normalize();

                            // Tính điểm mới
                            XYZ e1 = s1 + offsetVector;
                            XYZ e2 = s2 + offsetVector;

                            // Tạo đoạn thẳng mới
                            Line line1 = Line.CreateBound(s1, e1);
                            Line line2 = Line.CreateBound(s2, e2);



                            Plane planestartpoint = Plane.CreateByNormalAndOrigin(lineDuct1.Direction, Lsp2);
                            // planestartpoint sẽ cắt các ống ngoài cùng theo điểm đầu của ống làm chuẩn
                            XYZ pointSOnDuctLine1 = CT.LineIntersectPlane(line1, planestartpoint);
                            XYZ pointSOnDuctLine2 = CT.LineIntersectPlane(line2, planestartpoint);

                            //MessageBox.Show($"{line1.Length*304.8}\n{line2.Length*304.80}");
                            Line lineConnect2sp = Line.CreateBound(pointSOnDuctLine1, pointSOnDuctLine2);
                            //Điểm đầu của đường trung tâm
                            XYZ Startp = CT.GetMidpoint(lineConnect2sp);

                            Plane planeEndpoint = Plane.CreateByNormalAndOrigin(lineDuct1.Direction, Lep2);
                            // planestartpoint sẽ cắt các ống ngoài cùng theo điểm đầu của ống làm chuẩn
                            XYZ pointEOnDuctLine1 = CT.LineIntersectPlane(line1, planeEndpoint);
                            XYZ pointEOnDuctLine2 = CT.LineIntersectPlane(line2, planeEndpoint);

                            Line lineConnect2ep = Line.CreateBound(pointEOnDuctLine1, pointEOnDuctLine2);
                            //Điểm đầu của đường trung tâm
                            XYZ EndP = CT.GetMidpoint(lineConnect2ep);

                            Line PlaceLine = Line.CreateBound(Startp, EndP);
                            XYZ Pointdistance = CT.FindPointOnLineFromStartPoint(PlaceLine, distanceStart);
                            Line PlaceSUP = Line.CreateBound(Pointdistance, EndP);
                            double number = (PlaceSUP.Length - distanceStart) / distanceOffset;
                            int total = (int)Math.Truncate(number) + 1;

                            //Thanh gia do
                            //double radianThanh = PlaceSUP.Direction.AngleTo(XYZ.BasisX);
                            double radianThanh = PlaceSUP.Direction.AngleTo(XYZ.BasisY);
                            double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);

                            

                            using (Transaction t = new Transaction(doc, " "))
                            {
                                for (int i = 0; i < total; i++)
                                {
                                    XYZ point = CT.FindPointOnLineFromStartPoint(PlaceSUP, i * distanceOffset);

                                    t.Start();
                                    Plane planeCumU = Plane.CreateByNormalAndOrigin(newline1.Direction, point);



                                    XYZ pointgiado = CT.LineIntersectPlane(PlaceSUP, planeCumU);

                                    FamilySymbol symbol = CT.GetFamilySymbol(doc, familyName, typeName);
                                    if (!symbol.IsActive) symbol.Activate();
                                    FamilyInstance support = doc.Create.NewFamilyInstance(pointgiado, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                    XYZ pointOnFace = CT.PointIntersectPlane(pointgiado, infinitePlane);
                                    XYZ pointOnLine = CT.PointOnLine(pointgiado, PlaceSUP);
                                    if (/*pointOnFace != null &&*/ pointOnLine != null)
                                    {
                                        support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(PlaceSUP.GetEndPoint(0).Z);
                                        support.LookupParameter("L_in").Set(distance);
                                        double height = pointOnFace.DistanceTo(pointOnLine);


                                        support.LookupParameter("l_ty").Set(100/304.8);
                                    }
                                    else
                                    {
                                        doc.Delete(support.Id);
                                    }

                                    //ROTARE
                                    XYZ pointZ = pointgiado + new XYZ(0, 0, 1);
                                    Line axis = Line.CreateBound(pointgiado, pointZ);

                                 

                                    if (degreeThanh == 0 || degreeThanh == 180) //ông song song trục Z
                                    {
                                       
                                        (support.Location as LocationPoint).Rotate(axis, Math.PI / 2);
                                        
                                    }
                                    if (degreeThanh != 0 && degreeThanh != 180 && degreeThanh != 90) // ống không song song trục X và Y
                                    {
                                        (support.Location as LocationPoint).Rotate(axis, Math.PI / 2);
                                      

                                    }

                                    t.Commit();
                                }

                            }
                        }

                        else
                        {
                            MessageBox.Show("Các ống gió chưa nằm cùng mặt phẳng", "Cảnh báo!");
                        }


                    }










                    else if (selectedRefs.Count == 1)
                    {
                       

                        Reference reference2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new DuctFilter(), "Hãy chọn ống làm chuẩn");

                        XYZ startPoint = reference2.GlobalPoint;

                        Duct duct1 = selectedRefs[0] as Duct;
                        double heightduct = Math.Round((duct1.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble() * 0.5 + duct1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        double widthduct = Math.Round((duct1.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble() + 2 * duct1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        LocationCurve locationCurvestd = duct1.Location as LocationCurve;
                        Line ductLine2 = CT.FindDirectionOfElement(reference2, locationCurvestd);
                        XYZ spnewline2 = new XYZ(ductLine2.GetEndPoint(0).X, ductLine2.GetEndPoint(0).Y, (ductLine2.GetEndPoint(0).Z - heightduct));
                        XYZ epnewline2 = new XYZ(ductLine2.GetEndPoint(1).X, ductLine2.GetEndPoint(1).Y, (ductLine2.GetEndPoint(1).Z - heightduct)); ;
                        Line newline2 = Line.CreateBound(spnewline2, epnewline2);
                        XYZ point0 = CT.FindPointOnLineFromStartPoint(newline2, distanceStart);
                        XYZ point1 = newline2.GetEndPoint(1);
                        Line PlaceSUP = Line.CreateBound(point0, point1);
                        //MessageBox.Show((PlaceSUP.Length * 304.8).ToString());
                        Line PlaceCumU = newline2;
                        XYZ director = newline2.GetEndPoint(1) - newline2.GetEndPoint(0);
                        double number = (PlaceSUP.Length) / distanceOffset;
                       
                        int total = (int)Math.Truncate(number) + 1;
                        //MessageBox.Show(number.ToString());




                        //Thanh gia do
                        double radianThanh = PlaceSUP.Direction.AngleTo(XYZ.BasisY);
                        double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);
                     

                        using (Transaction t = new Transaction(doc, " "))
                        {

                            XYZ point = reference2.GlobalPoint;

                            t.Start();
                            for (int i = 0; i < total; i++)
                            {
                                XYZ pointx = CT.FindPointOnLineFromStartPoint(PlaceSUP, i * distanceOffset);
                                Plane planeCumU = Plane.CreateByNormalAndOrigin(director, pointx);





                                //}
                                XYZ pointgiado = CT.LineIntersectPlane(PlaceSUP, planeCumU);

                                FamilySymbol symbol = CT.GetFamilySymbol(doc, familyName, typeName);
                                if (!symbol.IsActive) symbol.Activate();
                                FamilyInstance support = doc.Create.NewFamilyInstance(pointgiado, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                XYZ pointOnFace = CT.PointIntersectPlane(pointgiado, infinitePlane);
                                XYZ pointOnLine = CT.PointOnLine(pointgiado, PlaceSUP);
                                if (/*pointOnFace != null &&*/ pointOnLine != null)
                                {
                                    support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(pointgiado.Z);
                                    support.LookupParameter("L_in").Set(widthduct);
                                    double height = pointOnFace.DistanceTo(pointOnLine);


                                    support.LookupParameter("l_ty").Set(100/304.8);
                                }
                                else
                                {
                                    doc.Delete(support.Id);
                                }

                                //ROTARE
                                XYZ pointZ = pointgiado + new XYZ(0, 0, 1);
                                Line axis = Line.CreateBound(pointgiado, pointZ);

                                if (degreeThanh == 0 || degreeThanh == 180) //ông song song trục Z
                                {
                                    (support.Location as LocationPoint).Rotate(axis, Math.PI / 2);

                                }
                                if (degreeThanh != 0 && degreeThanh != 180 && degreeThanh != 90) // ống không song song trục X và Y
                                {
                                    (support.Location as LocationPoint).Rotate(axis, Math.PI / 2 - radianThanh);
                                    
                                }
                            }
                            t.Commit();





                        }
                    }
                    else
                    {
                        break;
                    }




                }
            }
            catch
            {

            }
        }

        public static void D_ThanhUGiaDo_Manual(UIDocument uidoc, Document doc, Plane infinitePlane, double distanceStart, double distanceOffset, string familyName, string typeName, bool Auto)
        {

            try
            {
               

            while (!Auto)
            {
                    IList<Element> selectedRefs = uidoc.Selection.PickElementsByRectangle(new DuctFilter(), "Select region to filter elements");

                    if (selectedRefs.Count > 1)
                {
                    Reference reference1 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new DuctFilter(), "Hãy chọn ống làm chuẩn");

                    XYZ startPoint = reference1.GlobalPoint;


                    Duct ductstd = doc.GetElement(reference1) as Duct;
                    LocationCurve locationCurvestd = ductstd.Location as LocationCurve;
                    Line ductLinestd = CT.FindDirectionOfElement(reference1, locationCurvestd);


                    List<Duct> ducts = CT.D_Get2ductmaxdistance(selectedRefs, startPoint);
                    Duct duct1 = ducts[0];
                    Duct duct2 = ducts[1];
                    //MessageBox.Show("Tôi bị ngu");

                    // Ống thứ 1
                    LocationCurve locationduct1curve = duct1.Location as LocationCurve;
                    Line lineDuct1 = locationduct1curve.Curve as Line;
                    lineDuct1 = CT.D_GetNearestStartPP(duct1, startPoint);
                    double lowpoint1 = Math.Round((duct1.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble() * 0.5 + duct1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);

                    XYZ sp1 = lineDuct1.GetEndPoint(0);
                    XYZ ep1 = lineDuct1.GetEndPoint(1);

                    XYZ Lsp1 = new XYZ(sp1.X, sp1.Y, (sp1.Z - lowpoint1));
                    XYZ Lep1 = new XYZ(ep1.X, ep1.Y, (ep1.Z - lowpoint1));
                    Line newline1 = Line.CreateBound(Lsp1, Lep1);


                    // Ống thứ 2

                    LocationCurve locationpipe2curve = duct2.Location as LocationCurve;
                    Line linePipe2 = locationpipe2curve.Curve as Line;
                    linePipe2 = CT.D_GetNearestStartPP(duct2, startPoint);
                    double lowpoint2 = Math.Round((duct2.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble() * 0.5 + duct2.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                    XYZ sp2 = linePipe2.GetEndPoint(0);
                    XYZ ep2 = linePipe2.GetEndPoint(1);

                    XYZ Lsp2 = new XYZ(sp2.X, sp2.Y, (sp2.Z - lowpoint2));
                    XYZ Lep2 = new XYZ(ep2.X, ep2.Y, (ep2.Z - lowpoint2));



                    if ((Math.Round(Lsp1.Z, 2)) == (Math.Round(Lsp2.Z, 2)))
                    {
                        Line newline2 = Line.CreateBound(Lsp2, Lep2);
                        XYZ directionsp = newline2.Direction;
                        Plane plane1 = Plane.CreateByNormalAndOrigin(directionsp, Lsp1);
                        XYZ p1 = CT.LineIntersectPlane(newline2, plane1);



                        Line distacelineSP = Line.CreateBound(Lsp1, p1);


                        double space1 = Math.Round((duct1.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble() * 0.5 + duct1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3); ;
                        double space2 = Math.Round((duct2.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble() * 0.5 + duct2.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3); ;


                        Line extendLineSP = CT.ExtendLine(distacelineSP, space1, space2);
                        double distance = Math.Round((extendLineSP.Length), 2);

                        Line PlaceLine;

                        //Line ong 1
                        XYZ s1 = extendLineSP.GetEndPoint(0);

                        //Line ong 2
                        XYZ s2 = extendLineSP.GetEndPoint(1);

                        // Tính toán vector dịch chuyển
                        XYZ offsetVector = ductLinestd.Length * ductLinestd.Direction.Normalize();

                        // Tính điểm mới
                        XYZ e1 = s1 + offsetVector;
                        XYZ e2 = s2 + offsetVector;

                        // Tạo đoạn thẳng mới
                        Line line1 = Line.CreateBound(s1, e1);
                        Line line2 = Line.CreateBound(s2, e2);



                        Plane planestartpoint = Plane.CreateByNormalAndOrigin(lineDuct1.Direction, Lsp2);
                        // planestartpoint sẽ cắt các ống ngoài cùng theo điểm đầu của ống làm chuẩn
                        XYZ pointSOnDuctLine1 = CT.LineIntersectPlane(line1, planestartpoint);
                        XYZ pointSOnDuctLine2 = CT.LineIntersectPlane(line2, planestartpoint);

                        //MessageBox.Show($"{line1.Length*304.8}\n{line2.Length*304.80}");
                        Line lineConnect2sp = Line.CreateBound(pointSOnDuctLine1, pointSOnDuctLine2);
                        //Điểm đầu của đường trung tâm
                        XYZ Startp = CT.GetMidpoint(lineConnect2sp);

                        Plane planeEndpoint = Plane.CreateByNormalAndOrigin(lineDuct1.Direction, Lep2);
                        // planestartpoint sẽ cắt các ống ngoài cùng theo điểm đầu của ống làm chuẩn
                        XYZ pointEOnDuctLine1 = CT.LineIntersectPlane(line1, planeEndpoint);
                        XYZ pointEOnDuctLine2 = CT.LineIntersectPlane(line2, planeEndpoint);

                        Line lineConnect2ep = Line.CreateBound(pointEOnDuctLine1, pointEOnDuctLine2);
                        //Điểm đầu của đường trung tâm
                        XYZ EndP = CT.GetMidpoint(lineConnect2ep);

                        PlaceLine = Line.CreateBound(Startp, EndP);
                        XYZ Pointdistance = CT.FindPointOnLineFromStartPoint(PlaceLine, 0);
                        Line PlaceSUP = Line.CreateBound(Pointdistance, EndP);
                        double number = (PlaceSUP.Length - 0) / distanceOffset;
                        int total = (int)Math.Truncate(number) + 1;

                        //Thanh gia do
                        double radianThanh = PlaceSUP.Direction.AngleTo(XYZ.BasisY);
                        double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);

                        using (Transaction t = new Transaction(doc, " "))
                        {

                            XYZ point = reference1.GlobalPoint;

                            t.Start();
                            Plane planeCumU = Plane.CreateByNormalAndOrigin(PlaceSUP.Direction, point);






                            //}
                            XYZ pointgiado = CT.LineIntersectPlane(PlaceSUP, planeCumU);

                            FamilySymbol symbol = CT.GetFamilySymbol(doc, familyName, typeName);
                            if (!symbol.IsActive) symbol.Activate();
                            FamilyInstance support = doc.Create.NewFamilyInstance(pointgiado, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                            XYZ pointOnFace = CT.PointIntersectPlane(pointgiado, infinitePlane);
                            XYZ pointOnLine = CT.PointOnLine(pointgiado, PlaceSUP);
                            if (/*pointOnFace != null &&*/ pointOnLine != null)
                            {
                                support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(PlaceSUP.GetEndPoint(0).Z);
                                support.LookupParameter("L_in").Set(distance);
                                double height = pointOnFace.DistanceTo(pointOnLine);


                                support.LookupParameter("l_ty").Set(100 / 304.8);
                            }
                            else
                            {
                                doc.Delete(support.Id);
                            }

                            //ROTARE
                            XYZ pointZ = pointgiado + new XYZ(0, 0, 1);
                            Line axis = Line.CreateBound(pointgiado, pointZ);

                            if (degreeThanh == 0 || degreeThanh == 180) //ông song song trục Z
                            {
                                (support.Location as LocationPoint).Rotate(axis, Math.PI / 2);
                            }
                            if (degreeThanh != 0 && degreeThanh != 180 && degreeThanh != 90) // ống không song song trục X và Y
                            {
                                (support.Location as LocationPoint).Rotate(axis, Math.PI / 2 - radianThanh);
                            }

                            t.Commit();
                        }

                    }

                    else
                    {
                        MessageBox.Show("Các ống gió chưa nằm cùng mặt phẳng", "Cảnh báo!");
                    }


                }

                
                else if (selectedRefs.Count == 1)
                {
                    Reference reference2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new DuctFilter(), "Hãy chọn ống làm chuẩn");

                    XYZ startPoint = reference2.GlobalPoint;

                    Duct duct1 = selectedRefs[0] as Duct;
                    double heigthduct = Math.Round((duct1.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble() * 0.5 + duct1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                    double widthduct= Math.Round((duct1.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble()  + duct1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        LocationCurve locationCurvestd = duct1.Location as LocationCurve;
                    Line ductLine2 = CT.FindDirectionOfElement(reference2, locationCurvestd);
                    XYZ spnewline2 = new XYZ(ductLine2.GetEndPoint(0).X, ductLine2.GetEndPoint(0).Y, (ductLine2.GetEndPoint(0).Z - heigthduct));
                    XYZ epnewline2 = new XYZ(ductLine2.GetEndPoint(1).X, ductLine2.GetEndPoint(1).Y, (ductLine2.GetEndPoint(1).Z - heigthduct)); ;
                    Line newline2 = Line.CreateBound(spnewline2, epnewline2);
                    XYZ point0 = CT.FindPointOnLineFromStartPoint(newline2, 0);
                    XYZ point1 = newline2.GetEndPoint(1);
                    Line PlaceSUP = Line.CreateBound(point0, point1);
                    Line PlaceCumU = newline2;
                    XYZ director = newline2.GetEndPoint(1) - newline2.GetEndPoint(0);
                    double number = (PlaceSUP.Length ) / distanceOffset;
                    int total = (int)Math.Truncate(number) + 1;




                    //Thanh gia do
                    double radianThanh = PlaceSUP.Direction.AngleTo(XYZ.BasisY);
                    double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);

                    using (Transaction t = new Transaction(doc, " "))
                    {

                        XYZ point = reference2.GlobalPoint;

                        t.Start();
                        Plane planeCumU = Plane.CreateByNormalAndOrigin(PlaceSUP.Direction, point);






                        //}
                        XYZ pointgiado = CT.LineIntersectPlane(PlaceSUP, planeCumU);

                        FamilySymbol symbol = CT.GetFamilySymbol(doc, familyName, typeName);
                        if (!symbol.IsActive) symbol.Activate();
                        FamilyInstance support = doc.Create.NewFamilyInstance(pointgiado, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                        XYZ pointOnFace = CT.PointIntersectPlane(pointgiado, infinitePlane);
                        XYZ pointOnLine = CT.PointOnLine(pointgiado, PlaceSUP);
                        if (/*pointOnFace != null &&*/ pointOnLine != null)
                        {
                            support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(PlaceSUP.GetEndPoint(0).Z);
                            support.LookupParameter("L_in").Set(widthduct);
                            double height = pointOnFace.DistanceTo(pointOnLine);


                            support.LookupParameter("l_ty").Set(100 / 304.8);
                        }
                        else
                        {
                            doc.Delete(support.Id);
                        }

                        //ROTARE
                        XYZ pointZ = pointgiado + new XYZ(0, 0, 1);
                        Line axis = Line.CreateBound(pointgiado, pointZ);

                        if (degreeThanh == 0 || degreeThanh == 180) //ông song song trục Z
                        {
                            (support.Location as LocationPoint).Rotate(axis, Math.PI / 2);
                        }
                        if (degreeThanh != 0 && degreeThanh != 180 && degreeThanh != 90) // ống không song song trục X và Y
                        {
                            (support.Location as LocationPoint).Rotate(axis, Math.PI / 2 - radianThanh);
                        }
                        //aaa
                        t.Commit();
                    }
                }

                else
                {
                    break;
                }
            }





            }
            catch
            {

            }

        }
    }

}
