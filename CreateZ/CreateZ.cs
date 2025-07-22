using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using DnBim_Tool;

namespace Dnbim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class CreateZ : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //uidoc= Lấy từ file đang mở
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //doc= để lấy dữ liệu trong dự án đó
            Document doc = uidoc.Document;
            // view= để lấy thông tin từ cái view đang làm việc 
            View view = doc.ActiveView;


            Reference reference = uidoc.Selection.PickObject(ObjectType.Element );
            Element element = doc.GetElement(reference);
            IList<Element> Elementlist = CT.Getfamilynameat2pointSPEP(element);
            using (Transaction t = new Transaction(doc, "Connect Duct to Fittings"))
            {
                t.Start();
                // Lấy các connectors của ống gió
                ConnectorSet Eleconnector = CT.GetConnectors(element);
                if (Eleconnector == null || Eleconnector.Size == 0)
                {
                    TaskDialog.Show("Error", "No connectors found on the duct.");

                }
                foreach (Connector connector in Eleconnector)
                {
                    foreach (Element e1 in Elementlist)
                    {
                        if (element == null) continue;
                        ConnectorSet e1connector = CT.GetConnectors(e1);
                        foreach(Connector c1 in e1connector)
                        {
                            try
                            {
                                double x = ((c1.Origin).DistanceTo(connector.Origin) * 304.8);
                                if (x < 1 / 304.8)
                                {
                                    if (e1 is Duct duct)
                                    {
                                        XYZ p1 = ((duct.Location as LocationCurve).Curve as Line).GetEndPoint(0);
                                        XYZ p2 = ((duct.Location as LocationCurve).Curve as Line).GetEndPoint(1);
                                        if (p1.DistanceTo(c1.Origin) < 1 / 304.8)
                                        {
                                            SlpitDuctFormStartPoint(doc, duct, 300 / 304.8, 0);
                                        }
                                        else
                                        {
                                            SlpitDuctFormEndPoint(doc, duct, 300 / 304.8, 0);
                                        }
                                    }
                                }
                            }

                            catch
                            {

                            }
                          

                        } 
                      


                    }

                }

                t.Commit();
            }


                return Result.Succeeded;
        }
        public  void SlpitDuctFormStartPoint(Document doc, Duct originalDuct, double distance, double lastSegmentLength)
        {



            LocationCurve locationCurve = originalDuct.Location as LocationCurve;
            Line locationLine = locationCurve.Curve as Line;
            Line lineRemain = locationLine;
            List<ElementId> listId = new List<ElementId>();

            double totalLength = locationLine.Length;
            double num = Math.Floor((totalLength - lastSegmentLength) / distance); // Tính số đoạn cắt có chiều dài đều (số đoạn chia đều)


            // Cắt các đoạn có chiều dài đều (num-1 đoạn)
            for (int i = 0; i < 1; i++)
            {
                try
                {
                    XYZ p = CT.FindPointOnLineFromStartPoint(lineRemain, distance); // Tìm điểm cắt
                    ElementId id = MechanicalUtils.BreakCurve(doc, originalDuct.Id, p);
                   CT.CreateUnionFiting(doc, originalDuct, doc.GetElement(id) as Duct);
                }
                catch
                {
                    // Xử lý ngoại lệ nếu cần
                }
            }

            // Tạo điểm cắt cho đoạn cuối sao cho chiều dài của nó là lastSegmentLength
            XYZ finalPoint = CT.FindPointOnLineFromEndPoint(lineRemain, lastSegmentLength); // Tính điểm cắt cho đoạn cuối
            ElementId finalId = MechanicalUtils.BreakCurve(doc, originalDuct.Id, finalPoint);
            listId.Add(finalId); // Thêm phần tử cuối vào danh sách

            // Tạo các fitting giữa các đoạn
            CT.CreateUnionFiting(doc, originalDuct, doc.GetElement(finalId) as Duct);
        }
        public  void SlpitDuctFormEndPoint(Document doc, Duct originDuct, double distance, double lastSegmentLength)
        {


            LocationCurve locationCurve = originDuct.Location as LocationCurve;
            Line locationLine = locationCurve.Curve as Line;
            double number = Math.Round(locationLine.Length / distance, 0);
            int total = int.Parse(number.ToString());

            Line line = locationLine;
            List<ElementId> listId = new List<ElementId>();
            listId.Add(originDuct.Id);

            for (int i = 0; i < 1; i++)
            {
                try
                {
                    //ngắt ống gió
                    XYZ p = CT.FindPointOnLineFromEndPoint(line, distance);
                    ElementId id = MechanicalUtils.BreakCurve(doc, originDuct.Id, p);
                    listId.Add(id);

                    //reset data
                    originDuct = doc.GetElement(id) as Duct;
                    LocationCurve lc = originDuct.Location as LocationCurve;
                    line = lc.Curve as Line;
                }
                catch { }
            }

            CT.CreateUnionsFiting(doc, listId);
        }
    }
}
