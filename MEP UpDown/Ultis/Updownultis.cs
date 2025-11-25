using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB;
using DnBim_Tool;
using System.Windows;
using System.Windows.Media.Media3D;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using System.Xml.Linq;


namespace DnBim_Tool
{
    public static class Updownultis
    {

        public static void DuctCut45(Document doc, Reference r1, Reference r2, double offset, bool isTop)
        {
            
            View view =doc.ActiveView;
            Line locationLine=null;
            XYZ pick1 = r1.GlobalPoint;
            XYZ pick2 = r2.GlobalPoint;
            //a
            Element element1 = doc.GetElement(r1);
           
            if (element1 is Duct duct)
            {
              
                LocationCurve locationCurve = duct.Location as LocationCurve;
                 locationLine = locationCurve.Curve as Line;
            }
            if (element1 is Pipe pipe)
            {
                LocationCurve locationCurve = pipe.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
                
            }
            if (element1 is CableTray cableTray)
            {
                LocationCurve locationCurve = element1.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
              


            }
            if (element1 is Conduit conduit)
            {
                LocationCurve locationCurve = conduit.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
               
            }
            XYZ direction = locationLine.Direction;

            Plane plane1 = Plane.CreateByNormalAndOrigin(direction, pick1);
            Plane plane2 = Plane.CreateByNormalAndOrigin(direction, pick2);

            XYZ gd1 = CT.LineIntersectPlane(locationLine, plane1);
            XYZ gd2 = CT.LineIntersectPlane(locationLine, plane2);

            double dis = offset / Math.Tan(CT.ToRadian(45));

            Line line = Line.CreateBound(gd1, gd2);

            XYZ p1 = CT.FindPointOnLineFromStartPoint(line, -dis);
            XYZ p2 = CT.FindPointOnLineFromEndPoint(line, -dis);

            double length = p1.DistanceTo(p2);


            IList<ElementId> ids = new List<ElementId>();

           

            using (Transaction t = new Transaction(doc, " "))
            {
                t.Start();
                ElementId id1 = null;
                if (element1 is Duct duct1)
                {
                    
                    id1 = MechanicalUtils.BreakCurve(doc, element1.Id, p1);
                    ids.Add(id1);
                }
                if (element1 is Pipe pipe1)
                {
                    id1 = PlumbingUtils.BreakCurve(doc, element1.Id, p1);
                    ids.Add(id1);
                   
                }
                if (element1 is CableTray cableTra1)
                {
                   id1=CT.SliptCableTray(doc, element1.Id, p1);
                    ids.Add(id1);
                }
                if (element1 is Conduit conduit1)
                {
                    id1 = CT.SliptConduit(doc, element1.Id, p1);
                    ids.Add(id1);
                    
                }
                //a
                ElementId id2 = null;
                ElementId id3 = null;

                try
                {
                    if (element1 is Autodesk.Revit.DB.Mechanical.Duct)
                    {
                       
                        id2 = MechanicalUtils.BreakCurve(doc, element1.Id, p2);
                        
                    }
                    else if (element1 is Autodesk.Revit.DB.Plumbing.Pipe )
                    {
                        id2 = PlumbingUtils.BreakCurve(doc, element1.Id, p2);
                        
                    }
                    else if (element1 is Autodesk.Revit.DB.Electrical.CableTray)
                    {
                        id2 = CT.SliptCableTray(doc, element1.Id, p2);
                    }
                    else if (element1 is Autodesk.Revit.DB.Electrical.Conduit)
                    {
                        id2 = CT.SliptConduit(doc, element1.Id, p2);
                    }
                }
                catch
                {
                    if (element1 is Duct)
                    {
                    
                        id3 = MechanicalUtils.BreakCurve(doc, id1, p2);
                    }
                    if (element1 is Pipe)
                    {
                        id3 = PlumbingUtils.BreakCurve(doc, id1, p2);
                    }
                    if (element1 is CableTray)
                    {
                        id3 = CT.SliptCableTray(doc, id1, p2);
                    }
                    if (element1 is Conduit)
                    {
                        id3 = CT.SliptConduit(doc, id1, p2);
                    }
                }



                if (id2 != null) ids.Add(id2);
                if (id3 != null) ids.Add(id3);
                ids.Add(element1.Id);
               
                CT.DeleteElement(doc, ids, length, out List<ElementId> newIds);
                XYZ ngd1;
                XYZ ngd2;
                if (isTop)
                {
                    ngd1 = new XYZ(gd1.X, gd1.Y, gd1.Z + offset);
                    ngd2 = new XYZ(gd2.X, gd2.Y, gd2.Z + offset);
                }
                else
                {
                    ngd1 = new XYZ(gd1.X, gd1.Y, gd1.Z - offset);
                    ngd2 = new XYZ(gd2.X, gd2.Y, gd2.Z - offset);
                }
                Element ductTop = null;
                Element ductLeft = null;
                Element ductRight = null;
                ElementId ductTypeId = null;
                ElementId levelId = null;
                ElementId systemId = null;
                double width = 0;
                double height = 0;
                double diameter = 0;
                
                if (element1 is Duct)
                {
                    ductTypeId = element1.GetTypeId();
                    levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    systemId = element1.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
                    height = element1.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble();
                    width = element1.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble();

                    ductTop = Duct.Create(doc, systemId, ductTypeId, levelId, ngd1, ngd2);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);



                    ductLeft = Duct.Create(doc, systemId, ductTypeId, levelId, p1, ngd1);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);

                    ductRight = Duct.Create(doc, systemId, ductTypeId, levelId, p2, ngd2);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);

                  
                }



                if (element1 is Pipe)
                {
                    systemId = element1.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
                    ductTypeId = element1.GetTypeId();
                    levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    diameter = element1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();
                    ductTop = Pipe.Create(doc, systemId, ductTypeId, levelId, ngd1, ngd2);
                    ductTop.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);




                    ductLeft = Pipe.Create(doc, systemId, ductTypeId, levelId, p1, ngd1);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);

                    ductRight = Pipe.Create(doc, systemId, ductTypeId, levelId, p2, ngd2);
                    ductRight.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);
                }
                if (element1 is CableTray)
                {
                 
                    ductTypeId = element1.GetTypeId();
                    levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    height = element1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).AsDouble();
                    width = element1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).AsDouble();
                    ductTop = CableTray.Create(doc, ductTypeId, ngd1, ngd2,levelId) ;
                    ductTop.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).Set(height);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).Set(width);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE)
                        .Set(element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString());



                    ductLeft = CableTray.Create(doc, ductTypeId, p1, ngd1, levelId);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).Set(height);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).Set(width);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE)
                        .Set(element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString());

                    ductRight = CableTray.Create(doc, ductTypeId, p2, ngd2, levelId);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).Set(height);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).Set(width);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE)
                        .Set(element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString());
                }
                if (element1 is Conduit)
                {
                    ductTypeId = element1.GetTypeId();
                    ElementId systemTypeId = element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsElementId();
                    ductTypeId = element1.GetTypeId();
                    levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    diameter = element1.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).AsDouble();
                    ductTop = Conduit.Create(doc, ductTypeId, ngd1, ngd2, levelId);
                    ductTop = Conduit.Create(doc, ductTypeId, ngd1, ngd2, levelId);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter);




                    ductLeft = Conduit.Create(doc, ductTypeId, p1, ngd1, levelId);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter);
                  

                    ductRight = Conduit.Create(doc, ductTypeId, p2, ngd2, levelId);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter);
                }
                CT.CreateElbowFiting(doc, ductTop, ductLeft);
                CT.CreateElbowFiting(doc, ductTop, ductRight);

                Element d1 = null;
                Element d2 = null;
                if (element1 is Duct)
                {
                    d1 = doc.GetElement(newIds[0]) as Duct;
                    d2 = doc.GetElement(newIds[1]) as Duct;
                }

                else if (element1 is Pipe)
                {
                    d1 = doc.GetElement(newIds[0]) as Pipe;
                    d2 = doc.GetElement(newIds[1]) as Pipe;
                }
                else if (element1 is CableTray)
                {
                   
                    d1= doc.GetElement(newIds[0]) as CableTray;
                    d2= doc.GetElement(newIds[1]) as CableTray;
                }
                else if (element1 is Conduit)
                {

                    d1 = doc.GetElement(newIds[0]) as Conduit;
                    d2 = doc.GetElement(newIds[1]) as Conduit;
                }
                bool isIntersection = CT.ChekSolid(d1, ductLeft);
                //MessageBox.Show(ductLeft.Id.ToString());
                //MessageBox.Show(d2.Id.ToString());
                if (isIntersection)
                {
                   
                    
                    CT.CreateElbowFitingsocua(doc, d1, ductLeft);
                    CT.CreateElbowFiting(doc, d2, ductRight);
                }
                else
                {
                    
                    CT.CreateElbowFiting(doc, d2, ductLeft);
                    //MessageBox.Show("Toibingu");
                    CT.CreateElbowFiting(doc, d1, ductRight);

                }


                t.Commit();
            }

        }
        public static void DuctCut90(Document doc, Reference r1, Reference r2, double offset, bool isTop)
        {




            XYZ pick1 = r1.GlobalPoint;
            XYZ pick2 = r2.GlobalPoint;

            //Duct duct = doc.GetElement(r1) as Duct;
            Element element1 = doc.GetElement(r1);

            Line locationLine = null;
            if (element1 is Duct duct)
            {
                LocationCurve locationCurve = duct.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
            }
            if (element1 is Pipe pipe)
            {
                LocationCurve locationCurve = pipe.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
            }
            if (element1 is CableTray)
            {
                LocationCurve locationCurve = element1.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;



            }
            if (element1 is Conduit)
            {
                LocationCurve locationCurve = element1.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;



            }
            XYZ x = locationLine.GetEndPoint(0);
            XYZ y = locationLine.GetEndPoint(1);
            XYZ direction = locationLine.Direction;

            Plane plane1 = Plane.CreateByNormalAndOrigin(direction, pick1);
            Plane plane2 = Plane.CreateByNormalAndOrigin(direction, pick2);

            XYZ gd1 = CT.LineIntersectPlane(locationLine, plane1);
            XYZ gd2 = CT.LineIntersectPlane(locationLine, plane2);


            //double dis = offset / Math.Tan(CT.ToRadian(45));
            double dis = 0;

            Line line = Line.CreateBound(gd1, gd2);

            XYZ p1 = CT.FindPointOnLineFromStartPoint(line, -dis);
            XYZ p2 = CT.FindPointOnLineFromEndPoint(line, -dis);

            double length = p1.DistanceTo(p2);


            IList<ElementId> ids = new List<ElementId>();


            using (Transaction t = new Transaction(doc, " "))
            {
                t.Start();



                ElementId id1 = null;
                if (element1 is Duct)
                {
                    id1 = MechanicalUtils.BreakCurve(doc, element1.Id, p1);
                    ids.Add(id1);
                }
                if (element1 is Pipe)
                {
                    id1 = PlumbingUtils.BreakCurve(doc, element1.Id, p1);
                    ids.Add(id1);
                }
                if (element1 is CableTray)
                {
                    id1 = CT.SliptCableTray(doc, element1.Id, p1);
                    ids.Add(id1);

                }
                if (element1 is Conduit)
                {
                    id1 = CT.SliptConduit(doc, element1.Id, p1);
                    ids.Add(id1);
                }
                //a
                ElementId id2 = null;
                ElementId id3 = null;

                try
                {
                    if (element1 is Duct)
                    {
                        id2 = MechanicalUtils.BreakCurve(doc, element1.Id, p2);
                    }
                    if (element1 is Pipe)
                    {
                        id2 = PlumbingUtils.BreakCurve(doc, element1.Id, p2);
                    }
                    if (element1 is CableTray)
                    {
                        id2 = CT.SliptCableTray(doc, element1.Id, p2);
                    }
                    if (element1 is Conduit)
                    {
                        id2 = CT.SliptConduit(doc, element1.Id, p2);
                    }
                }
                catch
                {
                    if (element1 is Duct)
                    {
                        id3 = MechanicalUtils.BreakCurve(doc, id1, p2);
                    }
                    if (element1 is Pipe)
                    {
                        id3 = PlumbingUtils.BreakCurve(doc, id1, p2);
                    }
                    if (element1 is CableTray)
                    {
                        id3 = CT.SliptCableTray(doc, id1, p2);
                    }
                    if (element1 is Conduit)
                    {
                        id3 = CT.SliptConduit(doc, id1, p2);
                    }
                }



                if (id2 != null) ids.Add(id2);
                if (id3 != null) ids.Add(id3);
                ids.Add(element1.Id);

                CT.DeleteElement(doc, ids, length, out List<ElementId> newIds);
                XYZ ngd1;
                XYZ ngd2;
                if (isTop)
                {
                    ngd1 = new XYZ(gd1.X, gd1.Y, gd1.Z + offset);
                    ngd2 = new XYZ(gd2.X, gd2.Y, gd2.Z + offset);
                }
                else
                {
                    ngd1 = new XYZ(gd1.X, gd1.Y, gd1.Z - offset);
                    ngd2 = new XYZ(gd2.X, gd2.Y, gd2.Z - offset);
                }
                Element ductTop = null;
                Element ductLeft = null;
                Element ductRight = null;
                ElementId ductTypeId = null;
                ElementId levelId = null;
                ElementId systemId = null;
                double width = 0;
                double height = 0;
                double diameter = 0;
                if (element1 is Duct)
                {
                    ductTypeId = element1.GetTypeId();
                    levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    systemId = element1.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();
                    height = element1.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble();
                    width = element1.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble();

                    ductTop = Duct.Create(doc, systemId, ductTypeId, levelId, ngd1, ngd2);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);



                    ductLeft = Duct.Create(doc, systemId, ductTypeId, levelId, p1, ngd1);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);

                    ductRight = Duct.Create(doc, systemId, ductTypeId, levelId, p2, ngd2);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);
                }



                if (element1 is Pipe)
                {
                    systemId = element1.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
                    ductTypeId = element1.GetTypeId();
                    levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    diameter = element1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();
                    ductTop = Pipe.Create(doc, systemId, ductTypeId, levelId, ngd1, ngd2);
                    ductTop.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);




                    ductLeft = Pipe.Create(doc, systemId, ductTypeId, levelId, p1, ngd1);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);

                    ductRight = Pipe.Create(doc, systemId, ductTypeId, levelId, p2, ngd2);
                    ductRight.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);
                }
                if (element1 is CableTray)
                {

                    ductTypeId = element1.GetTypeId();
                    levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    height = element1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).AsDouble();
                    width = element1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).AsDouble();
                    ductTop = CableTray.Create(doc, ductTypeId, ngd1, ngd2, levelId);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).Set(height);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).Set(width);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE)
                        .Set(element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString());



                    ductLeft = CableTray.Create(doc, ductTypeId, p1, ngd1, levelId);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).Set(height);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).Set(width);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE)
                        .Set(element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString());

                    ductRight = CableTray.Create(doc, ductTypeId, p2, ngd2, levelId);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).Set(height);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).Set(width);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE)
                        .Set(element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString());


                }
                if (element1 is Conduit)
                {
                    ductTypeId = element1.GetTypeId();
                    ElementId systemTypeId = element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsElementId();
                    ductTypeId = element1.GetTypeId();
                    levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    diameter = element1.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).AsDouble();
                    ductTop = Conduit.Create(doc, ductTypeId, ngd1, ngd2, levelId);
                    ductTop.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter);




                    ductLeft = Conduit.Create(doc, ductTypeId, p1, ngd1, levelId);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter);


                    ductRight = Conduit.Create(doc, ductTypeId, p2, ngd2, levelId);
                    ductRight.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM).Set(diameter);
                }

                Element d1 = null;
                Element d2 = null;
                if (element1 is Duct)
                {
                    d1 = doc.GetElement(newIds[0]) as Duct;
                    d2 = doc.GetElement(newIds[1]) as Duct;
                }

                else if (element1 is Pipe)
                {
                    d1 = doc.GetElement(newIds[0]) as Pipe;
                    d2 = doc.GetElement(newIds[1]) as Pipe;
                }
                else if (element1 is CableTray)
                {

                    d1 = doc.GetElement(newIds[0]) as CableTray;
                    d2 = doc.GetElement(newIds[1]) as CableTray;
                }
                else if (element1 is Conduit)
                {
                    d1 = doc.GetElement(newIds[0]) as Conduit;
                    d2 = doc.GetElement(newIds[1]) as Conduit;
                }

                if (element1 is Duct)
                {
                    Line lineducttop = (ductTop.Location as LocationCurve).Curve as Line;
                    //Thanh gia do
                    double radianThanh = lineducttop.Direction.AngleTo(XYZ.BasisX);
                    double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);
                    //MessageBox.Show(degreeThanh.ToString());
                    if (degreeThanh == 0 || degreeThanh == 180) //ông song song trục X
                    {

                        ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), Math.PI / 2);
                        ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), Math.PI / 2);
                    }
                    else
                    {
                        if (degreeThanh == 90)
                        {
                            ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), (/*Math.PI / 2-*/radianThanh));
                            ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), /*Math.PI / 2 -*/ radianThanh);
                        }
                        ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), (Math.PI / 2 + radianThanh));
                        ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), (Math.PI / 2 + radianThanh));

                    }
                }
                //a
                if (element1 is CableTray)
                {
                    Line lineducttop = (ductTop.Location as LocationCurve).Curve as Line;
                    //Thanh gia do
                    //double radianThanh = locationLine.Direction.AngleTo(XYZ.BasisY);
                    //double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);

                    //if (degreeThanh==0)
                    //{
                    //    if (p1.DistanceTo(locationLine.GetEndPoint(0)) < p2.DistanceTo(locationLine.GetEndPoint(0)))
                    //    {
                    //        ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), Math.PI);

                    //    }
                    //    if (p1.DistanceTo(locationLine.GetEndPoint(0)) > p2.DistanceTo(locationLine.GetEndPoint(0)))
                    //    {
                    //        ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), Math.PI);

                    //    }
                    //}
                    //if (degreeThanh == 180)
                    //{
                    //    if (p1.DistanceTo(locationLine.GetEndPoint(0)) < p2.DistanceTo(locationLine.GetEndPoint(0)))
                    //    {

                    //        ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), Math.PI);
                    //    }
                    //    if (p1.DistanceTo(locationLine.GetEndPoint(0)) > p2.DistanceTo(locationLine.GetEndPoint(0)))
                    //    {
                    //        ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), Math.PI);
                    //    }






                    //}
                    //if (degreeThanh != 0 && degreeThanh != 180)
                    //{

                    //    ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), (Math.PI  - radianThanh));
                    //    //MessageBox.Show("Toibingu");
                    //    ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), (2*Math.PI  -radianThanh));
                    //}


                    double radianThanh = lineducttop.Direction.AngleTo(XYZ.BasisX);
                    double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 0);
                    //MessageBox.Show(degreeThanh.ToString());
                    if (degreeThanh == 0 || degreeThanh == 180) //ông song song trục X
                    {

                        ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), Math.PI / 2);
                        ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), Math.PI / 2);
                    }
                    else
                    {

                        if (degreeThanh == 90)
                        {
                            ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), (/*Math.PI / 2-*/radianThanh));
                            ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), /*Math.PI / 2 -*/ radianThanh);
                        }

                        else
                        {
                            if (x.DistanceTo(gd1) > y.DistanceTo(gd1))
                            {
                                MessageBox.Show("1");
                                ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), (Math.PI / 2 - radianThanh));
                                ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), (3 * Math.PI / 2 - radianThanh));
                            }
                            else
                            {
                                MessageBox.Show("2");
                                //a
                                ElementTransformUtils.RotateElement(doc, ductLeft.Id, ((ductLeft.Location as LocationCurve).Curve as Line), (-radianThanh +Math.PI/2));
                                //ElementTransformUtils.RotateElement(doc, ductRight.Id, ((ductRight.Location as LocationCurve).Curve as Line), (-radianThanh + Math.PI / 2));
                            }




                        }
                    }
                    
                    bool isIntersection = CT.ChekSolid(d1, ductLeft);

                    CT.CreateElbowFiting(doc, ductTop, ductLeft);
                    CT.CreateElbowFiting(doc, ductTop, ductRight);

                    if (isIntersection)
                    {
                        CT.CreateElbowFiting(doc, d1, ductLeft);
                        CT.CreateElbowFiting(doc, d2, ductRight);
                        //a
                    }
                    else
                    {


                        CT.CreateElbowFiting(doc, d2, ductLeft);
                        CT.CreateElbowFiting(doc, d1, ductRight);


                    }



                    t.Commit();
                }

            }
        }
        //a



        public static void DuctMove90(Document doc, Reference r1, Reference r2, double offset, bool isTop)
        {
            XYZ pick1 = r1.GlobalPoint;
            XYZ pick2 = r2.GlobalPoint;

           
            Element element1 = doc.GetElement(r1);

            Line locationLine = null;
            if (element1 is Duct duct)
            {
                LocationCurve locationCurve = duct.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
            }
            if (element1 is Pipe pipe)
            {
                LocationCurve locationCurve = pipe.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
            }
            if (element1 is CableTray)
            {
                LocationCurve locationCurve = element1.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;



            }
            XYZ direction = locationLine.Direction;

            Plane plane1 = Plane.CreateByNormalAndOrigin(direction, pick1);
            Plane plane2 = Plane.CreateByNormalAndOrigin(direction, pick2);
            XYZ gd1 = CT.LineIntersectPlane(locationLine, plane1);
            XYZ gd2 = CT.LineIntersectPlane(locationLine, plane2);
            double dis = 0;

            Line line = Line.CreateBound(gd1, gd2);

            XYZ p1 = CT.FindPointOnLineFromStartPoint(line, -dis);





            using (Transaction t = new Transaction(doc, " "))
            {
                t.Start();
                ElementId element2Id = null;
                if (element1 is Duct)
                {
                    element2Id = MechanicalUtils.BreakCurve(doc, element1.Id, p1);
                   
                }
                if (element1 is Pipe)
                {
                    element2Id = PlumbingUtils.BreakCurve(doc, element1.Id, p1);
                    
                }
                if (element1 is CableTray)
                {
                    element2Id = CT.SliptCableTray(doc, element1.Id, p1);
                    
                }
                Line locationLineductId1 = null;
                Element element2 = doc.GetElement(element2Id) ;
                if (element2 is Duct duct2)
                {
                    LocationCurve locationCurveductId1 = duct2.Location as LocationCurve;
                    locationLineductId1 = locationCurveductId1.Curve as Line;
                }
                if (element2 is Pipe pipe2)
                {
                    LocationCurve locationCurveductId1 = pipe2.Location as LocationCurve;
                    locationLineductId1 = locationCurveductId1.Curve as Line;
                }
                if (element2 is CableTray)
                {
                    LocationCurve locationCurveductId1 = element2.Location as LocationCurve;
                    locationLineductId1 = locationCurveductId1.Curve as Line;

                }




                XYZ moveVector = null;

                if (isTop)
                {
                    XYZ translationVector = new XYZ(0, 0, offset);
                    moveVector = translationVector;
                }
                else
                {
                    XYZ translationVector = new XYZ(0, 0, -offset);
                    moveVector = translationVector;
                }
                double length = Math.Round((gd2.DistanceTo(locationLine.GetEndPoint(0))) + (gd2.DistanceTo(locationLine.GetEndPoint(1))), 0);

                Element ductonggiochinh = null;
                Element ductonggiophu = null;
                if (length != Math.Round(locationLine.Length, 0))
                {
                    ElementTransformUtils.MoveElement(doc, element2Id, moveVector);
                    ductonggiochinh = element1;
                    ductonggiophu = element2;
                }
                else
                {
                    ElementTransformUtils.MoveElement(doc, element1.Id, moveVector);
                    ductonggiochinh = element2;
                    ductonggiophu = element1;
                }

                LocationCurve locationCurveduconggiophu = ductonggiophu.Location as LocationCurve;
                Line locationLineduconggiophu = locationCurveduconggiophu.Curve as Line;

                XYZ ngd1 = new XYZ(p1.X, p1.Y, (p1.Z + offset));
                Line linex = Line.CreateBound(gd1, ngd1);


                double d = Math.Round(ngd1.DistanceTo(locationLineduconggiophu.GetEndPoint(1)), 0);

                Element ductLeft = null;
                if (element1 is Duct)
                {
                    ElementId ductTypeId = ductonggiophu.GetTypeId();
                    ElementId levelId = ductonggiophu.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    ElementId systemId = ductonggiophu.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();


                    double height = ductonggiophu.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble();
                    double width = ductonggiophu.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble();




                    if (d != 0)
                    {
                        XYZ d1 = CT.FindPointOnLineFromStartPoint(locationLineduconggiophu, offset);
                        ductLeft = Duct.Create(doc, systemId, ductTypeId, levelId, d1, ngd1);
                    }
                    else
                    {
                        XYZ d1 = CT.FindPointOnLineFromEndPoint(locationLineduconggiophu, offset);
                        ductLeft = Duct.Create(doc, systemId, ductTypeId, levelId, d1, ngd1);
                    }
                  

                    ductLeft.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);
                }
                if(element1 is Pipe)
                {
                   ElementId systemId = element1.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
                  ElementId  ductTypeId = element1.GetTypeId();
                  ElementId  levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                   double diameter = element1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();




                    if (d != 0)
                    {
                        XYZ d1 = CT.FindPointOnLineFromStartPoint(locationLineduconggiophu, offset);
                        ductLeft = Pipe.Create(doc, systemId, ductTypeId, levelId, d1, ngd1);
                    }
                    else
                    {
                        XYZ d1 = CT.FindPointOnLineFromEndPoint(locationLineduconggiophu, offset);
                        ductLeft = Pipe.Create(doc, systemId, ductTypeId, levelId, d1, ngd1);
                    }
                    ductLeft.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);
                }
                if (element1 is CableTray)
                {
                    ElementId ductTypeId = ductonggiophu.GetTypeId();
                    ElementId levelId = ductonggiophu.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();


                    double height = element1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).AsDouble();
                    double width = element1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).AsDouble();


                    XYZ d1 = new XYZ();

                    if (d != 0)
                    {
                        d1 = CT.FindPointOnLineFromStartPoint(locationLineduconggiophu, offset);

                    }
                    else
                    {
                        d1 = CT.FindPointOnLineFromEndPoint(locationLineduconggiophu, offset);

                    }
                    ductLeft = CableTray.Create(doc, ductTypeId, d1, ngd1, levelId);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE)
                   .Set(element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString());

                    ductLeft.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).Set(height);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).Set(width);
                }


                    // Tính vector hướng hiện tại của ống gió dọc theo đường a

                    XYZ directionA = null;
                if (d != 0)
                { directionA = (ngd1 - locationLineduconggiophu.GetEndPoint(1)).Normalize(); }
                else
                { directionA = (ngd1 - locationLineduconggiophu.GetEndPoint(0)).Normalize(); }




                // Tính vector hướng mục tiêu dọc theo đường b
                XYZ directionB = (gd1 - ngd1).Normalize();

                // Bước 4: Tính toán trục và góc quay
                XYZ rotationAxis = directionA.CrossProduct(directionB).Normalize(); // Trục quay
                double angle = directionA.AngleTo(directionB); // Góc quay



                // Thực hiện phép quay ống gió quanh điểm X với trục quay và góc quay đã tính toán
                ElementTransformUtils.RotateElement(doc, ductLeft.Id, Line.CreateUnbound(ngd1, rotationAxis), -angle);

                CT.CreateElbowFiting(doc, ductonggiochinh, ductLeft);
                CT.CreateElbowFiting(doc, ductonggiophu, ductLeft);




                t.Commit();
            }

        }
        public static void DuctMove45(Document doc, Reference r1, Reference r2, double offset, bool isTop)
        {

            XYZ pick1 = r1.GlobalPoint;
            XYZ pick2 = r2.GlobalPoint;


            Element element1 = doc.GetElement(r1);

            Line locationLine = null;
            if (element1 is Duct duct)
            {
                LocationCurve locationCurve = duct.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
            }
            if (element1 is Pipe pipe)
            {
                LocationCurve locationCurve = pipe.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
            }
            if (element1 is CableTray)
            {
                LocationCurve locationCurve = element1.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;



            }
            XYZ direction = locationLine.Direction;

            Plane plane1 = Plane.CreateByNormalAndOrigin(direction, pick1);
            Plane plane2 = Plane.CreateByNormalAndOrigin(direction, pick2);
            XYZ gd1 = CT.LineIntersectPlane(locationLine, plane1);
            XYZ gd2 = CT.LineIntersectPlane(locationLine, plane2);
            double dis = offset ;

            Line line = Line.CreateBound(gd1, gd2);

            XYZ p1 = CT.FindPointOnLineFromStartPoint(line, dis);





            using (Transaction t = new Transaction(doc, " "))
            {
                t.Start();
                ElementId element2Id = null;
                if (element1 is Duct)
                {
                    element2Id = MechanicalUtils.BreakCurve(doc, element1.Id, gd1);

                }
                if (element1 is Pipe)
                {
                    element2Id = PlumbingUtils.BreakCurve(doc, element1.Id, gd1);

                }
                if (element1 is CableTray)
                {
                    element2Id = CT.SliptCableTray(doc, element1.Id, gd1);

                }
                Line locationLineductId1 = null;
                Element element2 = doc.GetElement(element2Id);
                if (element2 is Duct duct2)
                {
                    LocationCurve locationCurveductId1 = duct2.Location as LocationCurve;
                    locationLineductId1 = locationCurveductId1.Curve as Line;
                }
                if (element2 is Pipe pipe2)
                {
                    LocationCurve locationCurveductId1 = pipe2.Location as LocationCurve;
                    locationLineductId1 = locationCurveductId1.Curve as Line;
                }
                if (element2 is CableTray)
                {
                    LocationCurve locationCurveductId1 = element2.Location as LocationCurve;
                    locationLineductId1 = locationCurveductId1.Curve as Line;

                }




                XYZ moveVector = null;
                XYZ ngd1;
                XYZ ngd2;

                if (isTop)
                {
                    XYZ translationVector = new XYZ(0, 0, offset);
                    ngd1 = new XYZ(p1.X, p1.Y, p1.Z + offset);
                    moveVector = translationVector;
                }
                else
                {
                    XYZ translationVector = new XYZ(0, 0, -offset);
                    ngd1 = new XYZ(p1.X, p1.Y, p1.Z - offset);
                    moveVector = translationVector;
                }
                double length = Math.Round((gd2.DistanceTo(locationLine.GetEndPoint(0))) + (gd2.DistanceTo(locationLine.GetEndPoint(1))), 0);

                Element ductonggiochinh = null;
                Element ductonggiophu = null;
                if (length != Math.Round(locationLine.Length, 0))
                {
                    ElementTransformUtils.MoveElement(doc, element2Id, moveVector);
                    ductonggiochinh = element1;
                    ductonggiophu = element2;
                }
                else
                {
                    ElementTransformUtils.MoveElement(doc, element1.Id, moveVector);
                    ductonggiochinh = element2;
                    ductonggiophu = element1;
                }
                Line locationLineduconggiophu = (ductonggiophu.Location as LocationCurve).Curve as Line;
                double d = 0;
                if (gd1.DistanceTo(locationLineduconggiophu.GetEndPoint(1)) > gd1.DistanceTo(locationLineduconggiophu.GetEndPoint(0)))
                {
                    (ductonggiophu.Location as LocationCurve).Curve = Line.CreateBound(ngd1, locationLineduconggiophu.GetEndPoint(1));
                    locationLineduconggiophu = (ductonggiophu.Location as LocationCurve).Curve as Line;
                    d = 1;
                }
                else
                {
                    (ductonggiophu.Location as LocationCurve).Curve = Line.CreateBound(ngd1, locationLineduconggiophu.GetEndPoint(0));
                    locationLineduconggiophu = (ductonggiophu.Location as LocationCurve).Curve as Line;

                }


                Line linex = Line.CreateBound(gd1, ngd1);




                Element ductLeft = null;
                if (element1 is Duct)
                {
                    ElementId ductTypeId = ductonggiophu.GetTypeId();
                    ElementId levelId = ductonggiophu.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    ElementId systemId = ductonggiophu.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsElementId();


                    double height = ductonggiophu.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble();
                    double width = ductonggiophu.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble();


                  

                   
                    ductLeft = Duct.Create(doc, systemId, ductTypeId, levelId, ngd1, gd1);

                    ductLeft.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).Set(height);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).Set(width);
                }
                if (element1 is Pipe)
                {
                    ElementId systemId = element1.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId();
                    ElementId ductTypeId = element1.GetTypeId();
                    ElementId levelId = element1.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    double diameter = element1.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();




                  
                        ductLeft = Pipe.Create(doc, systemId, ductTypeId, levelId, ngd1,gd1);
                   
                    ductLeft.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).Set(diameter);
                }
                if (element1 is CableTray)
                {
                    ElementId ductTypeId = ductonggiophu.GetTypeId();
                    ElementId levelId = ductonggiophu.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();


                    double height = element1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).AsDouble();
                    double width = element1.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).AsDouble();


                    ductLeft = CableTray.Create(doc, ductTypeId,  ngd1, gd1, levelId);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE)
                   .Set(element1.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString());

                    ductLeft.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).Set(height);
                    ductLeft.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM).Set(width);
                }


              

                CT.CreateElbowFiting(doc, ductonggiochinh, ductLeft);
                CT.CreateElbowFiting(doc, ductonggiophu, ductLeft);




                t.Commit();
            }
        }
        private static void EnsureSameUpDirection(Document doc, CableTray trayRef, CableTray trayToFix)
        {
            // Lấy transform của tray gốc
            Line lineRef = (trayRef.Location as LocationCurve).Curve as Line;
            Transform trfRef = lineRef.ComputeDerivatives(0.5, true);
            XYZ upRef = trfRef.BasisZ.Normalize();

            // Lấy transform của tray cần kiểm tra
            Line lineFix = (trayToFix.Location as LocationCurve).Curve as Line;
            Transform trfFix = lineFix.ComputeDerivatives(0.5, true);
            XYZ upFix = trfFix.BasisZ.Normalize();

            // Nếu dot < 0, nghĩa là hai hướng up ngược nhau → xoay 180°
            if (upRef.DotProduct(upFix) < 0)
            {
                ElementTransformUtils.RotateElement(doc, trayToFix.Id, Line.CreateUnbound(lineFix.Origin, lineFix.Direction), Math.PI);
            }
        }
    }
}

