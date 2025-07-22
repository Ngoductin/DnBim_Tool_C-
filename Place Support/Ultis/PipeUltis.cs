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
using System.Net;

namespace DnBim_Tool
{
    public class PipeUltis
    {

        public static void P_Cum_Auto(UIDocument uidoc, Document doc, Plane infinitePlane, double distanceStart, double distanceOffset, string familyName, string typeName)
        {
            try
            {
                while (true)
                {

                    //pick ống nước
                    Reference reference = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element/*, new PipeFilter()*/, "Hãy chọn ống nước");
                    Pipe pipe = doc.GetElement(reference) as Pipe;

                    //lấy thông tin của ống nước 
                    LocationCurve locationCurve = pipe.Location as LocationCurve;

                    //Line pipeLine = locationCurve.Curve as Line;

                    Line pipeLine = CT.FindDirectionOfElement(reference, locationCurve);

                    double diameter = (pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() + 2 * pipe.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble());
                    ElementId levelId = pipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    Level level = doc.GetElement(levelId) as Level;


                    double number = (pipeLine.Length - distanceStart) / distanceOffset;
                    int total = (int)Math.Truncate(number) + 1;

                    //tính toán các thông tin
                    XYZ point0 = CT.FindPointOnLineFromStartPoint(pipeLine, distanceStart);
                    XYZ point1 = pipeLine.GetEndPoint(1);
                    Line newLine = Line.CreateBound(point0, point1);
                    double radian = newLine.Direction.AngleTo(XYZ.BasisX);
                    double degree = Math.Round(radian * 180 / Math.PI, 2);


                    using (Transaction t = new Transaction(doc, " "))
                    {

                        t.Start();

                        FamilySymbol symbol = CT.GetFamilySymbol(doc, familyName, typeName);


                        if (!symbol.IsActive) symbol.Activate();

                        for (int i = 0; i < total; i++)
                        {
                            XYZ point = CT.FindPointOnLineFromStartPoint(newLine, i * distanceOffset);
                            FamilyInstance support = doc.Create.NewFamilyInstance(point, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);



                            XYZ pointOnFace = CT.PointIntersectPlane(point, infinitePlane);

                            XYZ pointOnLine = CT.PointOnLine(point, pipeLine);



                            if (/*pointOnFace != null &&*/ pointOnLine != null)
                            {
                                support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(pointOnLine.Z);
                                support.LookupParameter("Nominal Diameter").Set(diameter);

                                double height = pointOnFace.DistanceTo(pointOnLine);
                                //double height = P_CT.CheckElementIntersection(element,pointOnLine);


                                support.LookupParameter("Rod Length").Set(100/304.8);
                            }
                            else
                                doc.Delete(support.Id);



                            //    ROTARE
                            XYZ pointZ = point + new XYZ(0, 0, 1);
                            Line axis = Line.CreateBound(point, pointZ);

                            if (degree == 0 || degree == 180) //ông song song trục Z
                            {
                                (support.Location as LocationPoint).Rotate(axis, Math.PI / 2);
                            }
                            if (degree != 0 && degree != 180 && degree != 90) // ống không song song trục X và Y
                            {
                                (support.Location as LocationPoint).Rotate(axis, Math.PI / 2 - radian);
                            }

                        }


                        t.Commit();
                    }
                }

            }
            catch
            {

            }
        }



        public static void P_Cum_Manual(UIDocument uidoc, Document doc, Plane infinitePlane, double distanceStart, double distanceOffset, string familyName, string typeName, bool Auto)
        {



            try
            {
                while (!Auto)
                {

                    //pick ống nước

                    Reference reference = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new PipeFilter(), "Chọn điểm đặt ");
                    Pipe pipe = doc.GetElement(reference) as Pipe;


                    //lấy thông tin của ống nước 
                    LocationCurve locationCurve = pipe.Location as LocationCurve;
                    Line pipeLine = locationCurve.Curve as Line;
                    Line locationLine = locationCurve.Curve as Line;
                    XYZ direction = locationLine.Direction;

                    XYZ pick1 = reference.GlobalPoint;
                    Plane plane1 = Plane.CreateByNormalAndOrigin(direction, pick1);
                    XYZ gd1 = CT.LineIntersectPlane(locationLine, plane1);

                    double diameter = (pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() + 2 * pipe.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble());
                    ElementId levelId = pipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    Level level = doc.GetElement(levelId) as Level;


                    //double number = (pipeLine.Length - distanceStart) / distanceOffset;
                    //int total = (int)Math.Truncate(number) + 1;

                    //tính toán các thông tin
                    XYZ point0 = CT.FindPointOnLineFromStartPoint(pipeLine, distanceStart);
                    XYZ point1 = pipeLine.GetEndPoint(1);
                    Line newLine = Line.CreateBound(point0, point1);
                    double radian = newLine.Direction.AngleTo(XYZ.BasisX);
                    double degree = Math.Round(radian * 180 / Math.PI, 2);


                    using (Transaction t = new Transaction(doc, " "))
                    {

                        t.Start();

                        FamilySymbol symbol = CT.GetFamilySymbol(doc, familyName, typeName);
                        if (!symbol.IsActive) symbol.Activate();


                        FamilyInstance support = doc.Create.NewFamilyInstance(gd1, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);



                        XYZ pointOnFace = CT.PointIntersectPlane(gd1, infinitePlane);
                        XYZ pointOnLine = CT.PointOnLine(gd1, pipeLine);



                        if (pointOnFace != null && pointOnLine != null)
                        {
                            support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(pointOnLine.Z);
                            support.LookupParameter("Nominal Diameter").Set(diameter);

                            double height = pointOnFace.DistanceTo(pointOnLine);
                            support.LookupParameter("Rod Length").Set(100/304.8);
                        }
                        else
                            doc.Delete(support.Id);


                        //ROTARE
                        XYZ pointZ = gd1 + new XYZ(0, 0, 1);
                        Line axis = Line.CreateBound(gd1, pointZ);

                        if (degree == 0 || degree == 180) //ông song song trục Z
                        {
                            (support.Location as LocationPoint).Rotate(axis, Math.PI / 2);
                        }
                        if (degree != 0 && degree != 180 && degree != 90) // ống không song song trục X và Y
                        {
                            (support.Location as LocationPoint).Rotate(axis, Math.PI / 2 - radian);
                        }

                        //}


                        t.Commit();
                    }
                }

            }



            catch
            {

            }


        }

        public static void P_ThanhUGiaDo_CumU_Auto(UIDocument uidoc, Document doc, Plane infinitePlane, double distanceStart, double distanceOffset, string familyName, string typeName)
        { 

            try
            {

                IList<Element> selectedRefs = uidoc.Selection.PickElementsByRectangle(new PipeFilter(), "Select region to filter elements");
                

                //a

                while (true)
                {



                    if (selectedRefs.Count > 1)
                    {
                        Reference reference1 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new PipeFilter(), "Hãy chọn ống nước làm chuẩn");

                        XYZ startPoint = reference1.GlobalPoint;


                        Pipe pipestd = doc.GetElement(reference1) as Pipe;
                        LocationCurve locationCurvestd = pipestd.Location as LocationCurve;
                        Line pipeLinestd = CT.FindDirectionOfElement(reference1, locationCurvestd);


                        List<Pipe> pipes = CT.P_Get2pipemaxdistance(selectedRefs, startPoint);
                        Pipe pipe1 = pipes[0];
                        Pipe pipe2 = pipes[1];


                        // Ống thứ 1
                        LocationCurve locationpipe1curve = pipe1.Location as LocationCurve;
                        Line linePipe1 = locationpipe1curve.Curve as Line;
                        linePipe1 = CT.P_GetNearestStartPP(pipe1, startPoint);
                        double x = pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
                        double lowpoint1 = Math.Round((pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() * 0.5 
                            + pipe1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        XYZ sp1 = linePipe1.GetEndPoint(0);
                        XYZ ep1 = linePipe1.GetEndPoint(1);

                        XYZ Lsp1 = new XYZ(sp1.X, sp1.Y, (sp1.Z - lowpoint1));
                        XYZ Lep1 = new XYZ(ep1.X, ep1.Y, (ep1.Z - lowpoint1));
                        Line newline1 = Line.CreateBound(Lsp1, Lep1);


                        // Ống thứ 2

                        LocationCurve locationpipe2curve = pipe2.Location as LocationCurve;
                        Line linePipe2 = locationpipe2curve.Curve as Line;
                        linePipe2 = CT.P_GetNearestStartPP(pipe2, startPoint);
                        double y = pipe2.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();

                        
                        double lowpoint2 = Math.Round((pipe2.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() * 0.5 
                            + pipe2.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        XYZ sp2 = linePipe2.GetEndPoint(0);
                        XYZ ep2 = linePipe2.GetEndPoint(1);

                        XYZ Lsp2 = new XYZ(sp2.X, sp2.Y, (sp2.Z - lowpoint2));
                        XYZ Lep2 = new XYZ(ep2.X, ep2.Y, (ep2.Z - lowpoint2));

                        bool checkslope = CT.P_ChecksameSlope(pipe1, pipe2);

                        if ((Math.Round(Lsp1.Z, 2)) == (Math.Round(Lsp1.Z, 2)) && checkslope == true)
                        {
                            Line newline2 = Line.CreateBound(Lsp2, Lep2);
                            XYZ directionsp = newline2.Direction;
                            Plane plane1 = Plane.CreateByNormalAndOrigin(directionsp, Lsp1);
                            XYZ p1 = CT.LineIntersectPlane(newline2, plane1);

                            //a

                            Line distacelineSP = Line.CreateBound(Lsp1, p1);


                            double space1 = lowpoint1;
                            double space2 = lowpoint2;


                            Line extendLineSP = CT.ExtendLine(distacelineSP, space1, space2);
                            //Line extendLineSP = CT.ExtendLine(distacelineSP, 0, 0);
                            double distance = Math.Round((extendLineSP.Length), 5);
                            //MessageBox.Show((distance*304.8).ToString());
                            

                            //Line ong 1
                            XYZ s1 = extendLineSP.GetEndPoint(0);
                           
                            //Line ong 2
                            XYZ s2 = extendLineSP.GetEndPoint(1);

                            // Tính toán vector dịch chuyển
                            XYZ offsetVector = pipeLinestd.Length * pipeLinestd.Direction.Normalize();

                            // Tính điểm mới
                            XYZ e1 = s1 + offsetVector;
                            XYZ e2 = s2 + offsetVector;

                            // Tạo đoạn thẳng mới
                            Line line1 = Line.CreateBound(s1, e1);
                            Line line2 = Line.CreateBound(s2, e2);



                            Plane planestartpoint = Plane.CreateByNormalAndOrigin(linePipe1.Direction, Lsp2);
                            // planestartpoint sẽ cắt các ống ngoài cùng theo điểm đầu của ống làm chuẩn
                            XYZ pointSOnPipeLine1 = CT.LineIntersectPlane(line1, planestartpoint);
                            XYZ pointSOnPipeLine2 = CT.LineIntersectPlane(line2, planestartpoint);

                            //MessageBox.Show($"{line1.Length*304.8}\n{line2.Length*304.80}");
                            Line lineConnect2sp = Line.CreateBound(pointSOnPipeLine1, pointSOnPipeLine2);
                            //Điểm đầu của đường trung tâm
                            XYZ Startp = CT.GetMidpoint(lineConnect2sp);

                            Plane planeEndpoint = Plane.CreateByNormalAndOrigin(linePipe1.Direction, Lep2);
                            // planestartpoint sẽ cắt các ống ngoài cùng theo điểm đầu của ống làm chuẩn
                            XYZ pointEOnPipeLine1 = CT.LineIntersectPlane(line1, planeEndpoint);
                            XYZ pointEOnPipeLine2 = CT.LineIntersectPlane(line2, planeEndpoint);

                            Line lineConnect2ep = Line.CreateBound(pointEOnPipeLine1, pointEOnPipeLine2);
                            //Điểm cuối của đường trung tâm
                            XYZ EndP = CT.GetMidpoint(lineConnect2ep);

                            Line PlaceLine = Line.CreateBound(Startp, EndP);
                            XYZ Pointdistance = CT.FindPointOnLineFromStartPoint(PlaceLine, distanceStart);
                            Line PlaceSUP = Line.CreateBound(Pointdistance, EndP);
                            double number = (PlaceSUP.Length - distanceStart) / distanceOffset;
                            int total = (int)Math.Truncate(number) + 1;

                            //Thanh gia do
                            double radianThanh = PlaceSUP.Direction.AngleTo(XYZ.BasisX);
                            double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);
                            //Cùm
                            double radianCumU = PlaceSUP.Direction.AngleTo(XYZ.BasisX);
                            double degreeCumU = Math.Round(radianCumU * 180 / Math.PI, 2);
                            //MessageBox.Show(degree.ToString());
                            using (Transaction t = new Transaction(doc, " "))
                            {
                                for (int i = 0; i < total; i++)
                                {
                                    XYZ point = CT.FindPointOnLineFromStartPoint(PlaceSUP, i * distanceOffset);
                                    //XYZ pointCumU=
                                    t.Start();
                                    Plane planeCumU = Plane.CreateByNormalAndOrigin(newline1.Direction, point);

                                   
                                    foreach (Element element in selectedRefs)
                                    {
                                        Pipe pipe = element as Pipe;

                                        double D = Math.Round((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() + 2 * pipe.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                                        Line pipeLine = (pipe.Location as LocationCurve).Curve as Line;
                                        XYZ gd = CT.LineIntersectPlane(pipeLine, planeCumU);
                                        FamilySymbol phukien = CT.GetFamilySymbolCumU(doc, "U_Cum_Hanger");
                                        if (!phukien.IsActive) phukien.Activate();
                                        FamilyInstance CumU = doc.Create.NewFamilyInstance(gd, phukien, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                        CumU.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(gd.Z);
                                        CumU.LookupParameter("ThanhU_Dout").Set(D);



                                        //ROTARE
                                        XYZ pointZCumU = gd + new XYZ(0, 0, 1);
                                        Line axisCumU = Line.CreateBound(gd, pointZCumU);

                                        if (degreeCumU == 0 || degreeCumU == 180) //ông song song trục Z
                                        {
                                            (CumU.Location as LocationPoint).Rotate(axisCumU, Math.PI / 2);
                                        }
                                        if (degreeCumU != 0 && degreeCumU != 180 && degreeCumU != 90) // ống không song song trục X và Y
                                        {
                                            (CumU.Location as LocationPoint).Rotate(axisCumU, Math.PI / 2 - radianCumU);
                                        }




                                    }
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


                                        support.LookupParameter("l_ty").Set(height);
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
                        }

                        else
                        {
                            MessageBox.Show("Các ống chưa nằm cùng mặt phẳng", "Cảnh báo!");
                        }


                    }










                    else if (selectedRefs.Count == 1)
                    {

                        Reference reference2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new PipeFilter(), "Hãy chọn ống nước làm chuẩn");

                        XYZ startPoint = reference2.GlobalPoint;

                        Pipe pipe1 = selectedRefs[0] as Pipe;
                        double distance = Math.Round((pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() + 2 * pipe1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        LocationCurve locationCurvestd = pipe1.Location as LocationCurve;
                        Line pipeLine2 = CT.FindDirectionOfElement(reference2, locationCurvestd);
                        XYZ spnewline2 = new XYZ(pipeLine2.GetEndPoint(0).X, pipeLine2.GetEndPoint(0).Y, (pipeLine2.GetEndPoint(0).Z - distance / 2));
                        XYZ epnewline2 = new XYZ(pipeLine2.GetEndPoint(1).X, pipeLine2.GetEndPoint(1).Y, (pipeLine2.GetEndPoint(1).Z - distance / 2)); ;
                        Line newline2 = Line.CreateBound(spnewline2, epnewline2);
                        XYZ point0 = CT.FindPointOnLineFromStartPoint(newline2, distanceStart);
                        XYZ point1 = newline2.GetEndPoint(1);
                        Line PlaceSUP = Line.CreateBound(point0, point1);
                        Line PlaceCumU = newline2;
                        XYZ director = newline2.GetEndPoint(1) - newline2.GetEndPoint(0);
                        double number = (PlaceSUP.Length - distanceStart) / distanceOffset;
                        int total = (int)Math.Truncate(number) + 1;




                        //Thanh gia do
                        double radianThanh = PlaceSUP.Direction.AngleTo(XYZ.BasisX);
                        double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);
                        //Cùm
                        double radianCumU = PlaceSUP.Direction.AngleTo(XYZ.BasisX);
                        double degreeCumU = Math.Round(radianCumU * 180 / Math.PI, 2);
                        using (Transaction t = new Transaction(doc, " "))
                        {

                            XYZ point = reference2.GlobalPoint;

                            t.Start();
                            for (int i = 0; i < total; i++)
                            {
                                XYZ pointx = CT.FindPointOnLineFromStartPoint(PlaceSUP, i * distanceOffset);
                                Plane planeCumU = Plane.CreateByNormalAndOrigin(director, pointx);

                                foreach (Element element in selectedRefs)
                                {
                                    Pipe pipe = element as Pipe;

                                    double D = Math.Round((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() + 2 * pipe.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                                    Line pipeLine = (pipe.Location as LocationCurve).Curve as Line;
                                    XYZ gd = CT.LineIntersectPlane(pipeLine, planeCumU);
                                    FamilySymbol phukien = CT.GetFamilySymbolCumU(doc, "U_Cum_Hanger");
                                    if (!phukien.IsActive) phukien.Activate();
                                    FamilyInstance CumU = doc.Create.NewFamilyInstance(gd, phukien, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                    CumU.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(gd.Z);
                                    CumU.LookupParameter("ThanhU_Dout").Set(D);

                                    //ROTARE
                                    XYZ pointZCumU = gd + new XYZ(0, 0, 1);
                                    Line axisCumU = Line.CreateBound(gd, pointZCumU);

                                    if (degreeCumU == 0 || degreeCumU == 180) //ông song song trục Z
                                    {
                                        (CumU.Location as LocationPoint).Rotate(axisCumU, Math.PI / 2);
                                    }
                                    if (degreeCumU != 0 && degreeCumU != 180 && degreeCumU != 90) // ống không song song trục X và Y
                                    {
                                        (CumU.Location as LocationPoint).Rotate(axisCumU, Math.PI / 2 - radianCumU);
                                    }



                                }
                                XYZ pointgiado = CT.LineIntersectPlane(PlaceSUP, planeCumU);

                                FamilySymbol symbol = CT.GetFamilySymbol(doc, familyName, typeName);
                                if (!symbol.IsActive) symbol.Activate();
                                FamilyInstance support = doc.Create.NewFamilyInstance(pointgiado, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                XYZ pointOnFace = CT.PointIntersectPlane(pointgiado, infinitePlane);
                                XYZ pointOnLine = CT.PointOnLine(pointgiado, PlaceSUP);
                                if (/*pointOnFace != null &&*/ pointOnLine != null)
                                {
                                    support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(pointgiado.Z);
                                    support.LookupParameter("L_in").Set(distance);
                                    double height = pointOnFace.DistanceTo(pointOnLine);


                                    support.LookupParameter("l_ty").Set(height);
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
            catch { }
        }



        public static void P_ThanhUGiaDo_CumU_Manual(UIDocument uidoc, Document doc, Plane infinitePlane, double distanceStart, double distanceOffset, string familyName, string typeName, bool Auto)
        {


            IList<Element> selectedRefs = uidoc.Selection.PickElementsByRectangle(new PipeFilter(), "Select region to filter elements");
            IList<double> pipeId = new List<double>();


            try
            {
                while (!Auto)
            {


                if (selectedRefs.Count > 1)
                    {
                        Reference reference1 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new PipeFilter(), "Hãy chọn ống nước làm chuẩn");

                        XYZ startPoint = reference1.GlobalPoint;


                        Pipe pipestd = doc.GetElement(reference1) as Pipe;
                        LocationCurve locationCurvestd = pipestd.Location as LocationCurve;
                        Line pipeLinestd = CT.FindDirectionOfElement(reference1, locationCurvestd);


                        List<Pipe> pipes = CT.P_Get2pipemaxdistance(selectedRefs, startPoint);
                        Pipe pipe1 = pipes[0];
                        Pipe pipe2 = pipes[1];


                        // Ống thứ 1
                        LocationCurve locationpipe1curve = pipe1.Location as LocationCurve;
                        Line linePipe1 = locationpipe1curve.Curve as Line;
                        linePipe1 = CT.P_GetNearestStartPP(pipe1, startPoint);
                        double lowpoint1 = Math.Round((pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() * 0.5 + pipe1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        XYZ sp1 = linePipe1.GetEndPoint(0);
                        XYZ ep1 = linePipe1.GetEndPoint(1);

                        XYZ Lsp1 = new XYZ(sp1.X, sp1.Y, (sp1.Z - lowpoint1));
                        XYZ Lep1 = new XYZ(ep1.X, ep1.Y, (ep1.Z - lowpoint1));
                        Line newline1 = Line.CreateBound(Lsp1, Lep1);


                        // Ống thứ 2

                        LocationCurve locationpipe2curve = pipe2.Location as LocationCurve;
                        Line linePipe2 = locationpipe2curve.Curve as Line;
                        linePipe2 = CT.P_GetNearestStartPP(pipe2, startPoint);
                        double lowpoint2 = Math.Round((pipe2.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() * 0.5 + pipe2.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        XYZ sp2 = linePipe2.GetEndPoint(0);
                        XYZ ep2 = linePipe2.GetEndPoint(1);

                        XYZ Lsp2 = new XYZ(sp2.X, sp2.Y, (sp2.Z - lowpoint2));
                        XYZ Lep2 = new XYZ(ep2.X, ep2.Y, (ep2.Z - lowpoint2));

                        bool checkslope = CT.P_ChecksameSlope(pipe1, pipe2);

                        if ((Math.Round(Lsp1.Z, 2)) == (Math.Round(Lsp1.Z, 2)) && checkslope == true)
                        {
                            Line newline2 = Line.CreateBound(Lsp2, Lep2);
                            XYZ directionsp = newline2.Direction;
                            Plane plane1 = Plane.CreateByNormalAndOrigin(directionsp, Lsp1);
                            XYZ p1 = CT.LineIntersectPlane(newline2, plane1);



                            Line distacelineSP = Line.CreateBound(Lsp1, p1);


                            double space1 = lowpoint1;
                            double space2 = lowpoint2;


                            Line extendLineSP = CT.ExtendLine(distacelineSP, space1, space2);
                            double distance = Math.Round((extendLineSP.Length), 2);


                            //Line ong 1
                            XYZ s1 = extendLineSP.GetEndPoint(0);

                            //Line ong 2
                            XYZ s2 = extendLineSP.GetEndPoint(1);

                            // Tính toán vector dịch chuyển
                            XYZ offsetVector = pipeLinestd.Length * pipeLinestd.Direction.Normalize();

                            // Tính điểm mới
                            XYZ e1 = s1 + offsetVector;
                            XYZ e2 = s2 + offsetVector;

                            // Tạo đoạn thẳng mới
                            Line line1 = Line.CreateBound(s1, e1);
                            Line line2 = Line.CreateBound(s2, e2);



                            Plane planestartpoint = Plane.CreateByNormalAndOrigin(linePipe1.Direction, Lsp2);
                            // planestartpoint sẽ cắt các ống ngoài cùng theo điểm đầu của ống làm chuẩn
                            XYZ pointSOnPipeLine1 = CT.LineIntersectPlane(line1, planestartpoint);
                            XYZ pointSOnPipeLine2 = CT.LineIntersectPlane(line2, planestartpoint);


                            Line lineConnect2sp = Line.CreateBound(pointSOnPipeLine1, pointSOnPipeLine2);
                            //Điểm đầu của đường trung tâm
                            XYZ Startp = CT.GetMidpoint(lineConnect2sp);

                            Plane planeEndpoint = Plane.CreateByNormalAndOrigin(linePipe1.Direction, Lep2);
                            // planestartpoint sẽ cắt các ống ngoài cùng theo điểm đầu của ống làm chuẩn
                            XYZ pointEOnPipeLine1 = CT.LineIntersectPlane(line1, planeEndpoint);
                            XYZ pointEOnPipeLine2 = CT.LineIntersectPlane(line2, planeEndpoint);

                            Line lineConnect2ep = Line.CreateBound(pointEOnPipeLine1, pointEOnPipeLine2);
                            //Điểm đầu của đường trung tâm
                            XYZ EndP = CT.GetMidpoint(lineConnect2ep);

                            Line PlaceSUP = Line.CreateBound(Startp, EndP);


                            //Thanh gia do
                            double radianThanh = PlaceSUP.Direction.AngleTo(XYZ.BasisX);
                            double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);
                            //Cùm
                            double radianCumU = PlaceSUP.Direction.AngleTo(XYZ.BasisX);
                            double degreeCumU = Math.Round(radianCumU * 180 / Math.PI, 2);
                            //MessageBox.Show(degree.ToString());
                            using (Transaction t = new Transaction(doc, " "))
                            {

                                XYZ point = reference1.GlobalPoint;

                                t.Start();
                                Plane planeCumU = Plane.CreateByNormalAndOrigin(PlaceSUP.Direction, point);

                                foreach (Element element in selectedRefs)
                                {
                                    Pipe pipe = element as Pipe;

                                    double D = Math.Round((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() + 2 * pipe.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                                    Line pipeLine = (pipe.Location as LocationCurve).Curve as Line;
                                    XYZ gd = CT.LineIntersectPlane(pipeLine, planeCumU);
                                    FamilySymbol phukien = CT.GetFamilySymbolCumU(doc, "U_Cum_Hanger");
                                    if (!phukien.IsActive) phukien.Activate();
                                    FamilyInstance CumU = doc.Create.NewFamilyInstance(gd, phukien, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                    CumU.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(gd.Z);
                                    CumU.LookupParameter("ThanhU_Dout").Set(D);



                                    //ROTARE
                                    XYZ pointZCumU = gd + new XYZ(0, 0, 1);
                                    Line axisCumU = Line.CreateBound(gd, pointZCumU);

                                    if (degreeCumU == 0 || degreeCumU == 180) //ông song song trục Z
                                    {
                                        (CumU.Location as LocationPoint).Rotate(axisCumU, Math.PI / 2);
                                    }
                                    if (degreeCumU != 0 && degreeCumU != 180 && degreeCumU != 90) // ống không song song trục X và Y
                                    {
                                        (CumU.Location as LocationPoint).Rotate(axisCumU, Math.PI / 2 - radianCumU);
                                    }




                                }
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


                                    support.LookupParameter("l_ty").Set(height);
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
                            MessageBox.Show("Các ống chưa nằm cùng mặt phẳng", "Cảnh báo!");
                        }


                    }

                    else if (selectedRefs.Count == 1)
                    {

                        Reference reference2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new PipeFilter(), "Hãy chọn ống nước làm chuẩn");

                        XYZ startPoint = reference2.GlobalPoint;

                        Pipe pipe1 = selectedRefs[0] as Pipe;
                        double distance = Math.Round((pipe1.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() + 2 * pipe1.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                        LocationCurve locationCurvestd = pipe1.Location as LocationCurve;
                        Line pipeLine2 = CT.FindDirectionOfElement(reference2, locationCurvestd);
                        XYZ spnewline2 = new XYZ(pipeLine2.GetEndPoint(0).X, pipeLine2.GetEndPoint(0).Y, (pipeLine2.GetEndPoint(0).Z - distance / 2));
                        XYZ epnewline2 = new XYZ(pipeLine2.GetEndPoint(1).X, pipeLine2.GetEndPoint(1).Y, (pipeLine2.GetEndPoint(1).Z - distance / 2)); ;
                        Line newline2 = Line.CreateBound(spnewline2, epnewline2);
                        XYZ point0 = CT.FindPointOnLineFromStartPoint(newline2, 0);
                        XYZ point1 = newline2.GetEndPoint(1);
                        Line PlaceSUP = Line.CreateBound(point0, point1);
                        Line PlaceCumU = newline2;
                        XYZ director = newline2.GetEndPoint(1) - newline2.GetEndPoint(0);
                        double number = (PlaceSUP.Length - distanceStart) / distanceOffset;
                        int total = (int)Math.Truncate(number) + 1;




                        //Thanh gia do
                        double radianThanh = PlaceSUP.Direction.AngleTo(XYZ.BasisX);
                        double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);
                        //Cùm
                        double radianCumU = PlaceSUP.Direction.AngleTo(XYZ.BasisX);
                        double degreeCumU = Math.Round(radianCumU * 180 / Math.PI, 2);
                        using (Transaction t = new Transaction(doc, " "))
                        {

                            XYZ point = reference2.GlobalPoint;

                            t.Start();
                            Plane planeCumU = Plane.CreateByNormalAndOrigin(director, point);

                            foreach (Element element in selectedRefs)
                            {
                                Pipe pipe = element as Pipe;

                                double D = Math.Round((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble() + 2 * pipe.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble()), 3);
                                Line pipeLine = (pipe.Location as LocationCurve).Curve as Line;
                                XYZ gd = CT.LineIntersectPlane(pipeLine, planeCumU);
                                FamilySymbol phukien = CT.GetFamilySymbolCumU(doc, "U_Cum_Hanger");
                                if (!phukien.IsActive) phukien.Activate();
                                FamilyInstance CumU = doc.Create.NewFamilyInstance(gd, phukien, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                CumU.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(gd.Z);
                                CumU.LookupParameter("ThanhU_Dout").Set(D);

                                //ROTARE
                                XYZ pointZCumU = gd + new XYZ(0, 0, 1);
                                Line axisCumU = Line.CreateBound(gd, pointZCumU);

                                if (degreeCumU == 0 || degreeCumU == 180) //ông song song trục Z
                                {
                                    (CumU.Location as LocationPoint).Rotate(axisCumU, Math.PI / 2);
                                }
                                if (degreeCumU != 0 && degreeCumU != 180 && degreeCumU != 90) // ống không song song trục X và Y
                                {
                                    (CumU.Location as LocationPoint).Rotate(axisCumU, Math.PI / 2 - radianCumU);
                                }



                            }
                            XYZ pointgiado = CT.LineIntersectPlane(PlaceSUP, planeCumU);

                            FamilySymbol symbol = CT.GetFamilySymbol(doc, familyName, typeName);
                            if (!symbol.IsActive) symbol.Activate();
                            FamilyInstance support = doc.Create.NewFamilyInstance(pointgiado, symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                            XYZ pointOnFace = CT.PointIntersectPlane(pointgiado, infinitePlane);
                            XYZ pointOnLine = CT.PointOnLine(pointgiado, PlaceSUP);
                            if (/*pointOnFace != null &&*/ pointOnLine != null)
                            {
                                support.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(pointgiado.Z);
                                support.LookupParameter("L_in").Set(distance);
                                double height = pointOnFace.DistanceTo(pointOnLine);


                                support.LookupParameter("l_ty").Set(height);
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
