using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using DnBim_Tool;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;

namespace Dnbim_Tool
{//a
    [Transaction(TransactionMode.Manual)]
    public class ALIGN_SUPPORT_CMD : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Line locationLine = null;
            // Chọn 2 đối tượng cần kiểm tra
            Reference obj1 = uidoc.Selection.PickObject(ObjectType.Element, "Chọn đối tượng cố định");
            Element element1 = doc.GetElement(obj1);
            double duongkinh=0;

            if (element1 is Duct duct)
            {
                LocationCurve locationCurve = duct.Location as LocationCurve;

                locationLine = locationCurve.Curve as Line;
                string size= duct.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsValueString();
                //if ( size== "ø100")
                //{ duongkinh = 100 / 304.8; }
            }
            if (element1 is Pipe pipe)
            {
                LocationCurve locationCurve = pipe.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
                //double x = Math.Round((((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)).AsDouble()*304.8)), 0);
                //MessageBox.Show((x ).ToString());
                //if (Math.Round((((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)).AsDouble() * 304.8)), 0) == 34)
                //{ duongkinh = 52 / 304.8; }
                //if (Math.Round((((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)).AsDouble() * 304.8)), 0) == 21)
                //{ duongkinh = 37 / 304.8; }
                //if (Math.Round((((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)).AsDouble() * 304.8)), 0) == 48)
                //{ duongkinh = 65 / 304.8; }
                //if (Math.Round((((pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)).AsDouble() * 304.8)), 0) == 27)
                //{ duongkinh = 45 / 304.8; }
            }
            if (element1 is CableTray)
            {
                LocationCurve locationCurve = element1.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;

                //a

            }
            if (element1 is Conduit conduit)
            {
                LocationCurve locationCurve = element1.Location as LocationCurve;
                locationLine = locationCurve.Curve as Line;
                string size = conduit.get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).AsValueString();
                //if (size == "53 mmø")
                //{ duongkinh = 53 / 304.8; }

                
            }

            XYZ direction = locationLine.Direction;
            //IList<Element> selectedRefs = uidoc.Selection.PickElementsByRectangle(new Pipe_Accessories_Filter (), "Chọn đối tượng support");
            //Reference  obj2 = uidoc.Selection.PickObject(ObjectType.Element,new Pipe_Accessories_Filter(), "Chọn support");
            //foreach (Element element2 in selectedRefs)
            //{
            Reference obj2 = uidoc.Selection.PickObject(ObjectType.Element, "Chọn đối tượng support");
            Element element2 = doc.GetElement(obj2);

            if (element2 is FamilyInstance support)
                {
                    LocationPoint location = support.Location as LocationPoint;
                    XYZ pointgiado = new XYZ();
                    if (location != null)
                    {
                        pointgiado = location.Point; // Đây là điểm gốc (insertion point)
                    }


                    //Thanh gia do
                   
                    using (Transaction t = new Transaction(doc, " "))
                    {
                        t.Start();

                        //ROTARE
                        XYZ pointZ = pointgiado + new XYZ(0, 0, 1);
                        Line axis = Line.CreateBound(pointgiado, pointZ);

                        //if (degreeThanh == 0 || degreeThanh == 180) //ông song song trục Z
                        //{
                        //    (support.Location as LocationPoint).Rotate(axis, Math.PI / 2);
                        //}
                        //if (degreeThanh != 0 && degreeThanh != 180 && degreeThanh != 90) // ống không song song trục X và Y
                        //{
                        //    (support.Location as LocationPoint).Rotate(axis, Math.PI / 2 - radianThanh);
                        //}
                        
                        //XYZ directorFamily=support.FacingOrientation;
                        // Lấy Transform của FamilyInstance
                        Transform transform = support.GetTransform();

                        // Hướng front thường là trục Y cục bộ (BasisY) trong hệ tọa độ toàn cục
                        XYZ directorFamily = transform.BasisX;

                        double radianThanh = direction.AngleTo(XYZ.BasisX);
                        double degreeThanh = Math.Round(radianThanh * 180 / Math.PI, 2);
                        if (Math.Round(degreeThanh) == 0 || Math.Round(degreeThanh) == 180)
                        {
                            //(support.Location as LocationPoint).Rotate(axis, Math.PI / 2);
                        }
                        if (Math.Round(degreeThanh) != 0 || Math.Round(degreeThanh) != 180|| Math.Round(degreeThanh) != 90)
                        {
                            //(support.Location as LocationPoint).Rotate(axis, Math.PI / 2);
                        }
                        Plane plane1 = Plane.CreateByNormalAndOrigin(direction, pointgiado);
                        XYZ gd1 = CT.LineIntersectPlane(locationLine, plane1);
                        XYZ moveVector = gd1 - pointgiado;
                        ElementTransformUtils.MoveElement(doc, support.Id, moveVector);
                    //support.LookupParameter("DN").Set(duongkinh);

                    t.Commit();
                    }
                //}
                
            }
                return Result.Succeeded;    
        }
        }
}
