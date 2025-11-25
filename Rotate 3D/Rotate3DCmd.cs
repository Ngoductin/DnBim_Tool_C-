


using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using DnBim_Tool.Rotate_3D;
//using Syncfusion.UI.Xaml.Grid;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class Rotate3DCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Đổi tên biến để tránh trùng với tham số "elements"
            IList<Element> selectedElements = new List<Element>();

            // Lấy các Element được chọn
            ICollection<ElementId> selectedElementIds = uidoc.Selection.GetElementIds();



            //try
            //{

                if (selectedElementIds.Count > 0)
                {
                    foreach (ElementId elementId in selectedElementIds)
                    {
                        Element element = doc.GetElement(elementId);
                        selectedElements.Add(element);
                    }
                }
                else
                {
                    // Các đối tượng bị xoay
                    selectedElements = uidoc.Selection.PickElementsByRectangle(new ElementbelongtoPipe(), "Chọn những đối tượng bị xoay");

                }

                #region Cmd



                while (true)
                {

                    if (selectedElements.Count == 0)
                    {
                        break;
                    }
                    var window = new Rotate3DView();
                    window.ShowDialog();
                    if (window.DialogResult == true)
                    {

                        //try
                        //{
                            // Chọn ra trục quay của các vật thể
                            Reference ref0 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new Pipe_ConduitFilter(), "Chọn ống trục");
                            Line pipeline0 = null;
                            if
                                (doc.GetElement(ref0) is Pipe)
                            {

                                // Lấy đối tượng Pipe từ Reference
                                Pipe pipe0 = doc.GetElement(ref0) as Pipe;
                                pipeline0 = (pipe0.Location as LocationCurve).Curve as Line;
                            }
                            if (doc.GetElement(ref0) is Conduit)
                            {

                                Conduit pipe0 = doc.GetElement(ref0) as Conduit;
                                pipeline0 = (pipe0.Location as LocationCurve).Curve as Line;
                            }

                            //Line pipeline0 = CT.GetDirectForAxis(doc, doc.GetElement(ref0), selectedElements);

                            // Kiểm tra nếu các đối tượng cùng trục
                            bool checkdongtruc = CT.CheckCungTruc(pipeline0, selectedElements[0]);

                            // Góc quay
                            double degree = Math.Round((Math.PI - (window.Angle * Math.PI / 180)), 5);
                            if (selectedElements[0].Category.Id.IntegerValue
                               == (int)BuiltInCategory.OST_PipeAccessory
                               && checkdongtruc == false
                               && selectedElements.Count == 1
                          )
                            {
                                using (Transaction trans = new Transaction(doc, "Rotate Elements"))
                                {
                                    Line lineconnector = CT.GetLineConnector(selectedElements[0]);
                                    trans.Start();

                                    foreach (Element element in selectedElements)
                                    {
                                        // Kiểm tra loại Location (LocationPoint hoặc LocationCurve)
                                        var locationPoint = element.Location as LocationPoint;

                                        if (locationPoint != null) // Nếu là LocationPoint (ví dụ: Pipe Fitting, Accessory)
                                        {

                                            locationPoint.Rotate(lineconnector, Math.PI - degree);
                                        }
                                    }
                                    trans.Commit();
                                }
                            }
                            else if (
                            selectedElements.Count != 0
                           )
                            {

                                using (Transaction trans = new Transaction(doc, "Rotate Elements"))
                                {
                                    trans.Start();
                                    Hashtable connectionMap = CT.DisconnectTwoEndsAndCache(doc, selectedElements);
                                    //HighlightConnectionMap(uidoc, doc, connectionMap);
                                  
                            foreach (Element element in selectedElements)
                                    { //// Gỡ kết nối các đối tượng

                                        // Kiểm tra loại Location (LocationPoint hoặc LocationCurve)
                                        var locationPoint = element.Location as LocationPoint;
                                        var locationCurve = element.Location as LocationCurve;

                                        if (locationPoint != null) // Nếu là LocationPoint (ví dụ: Pipe Fitting, Accessory)
                                        {
                                            locationPoint.Rotate(pipeline0, Math.PI - degree);
                                        }
                                        else if (locationCurve != null) // Nếu là LocationCurve (ví dụ: Pipe)
                                        {
                                            locationCurve.Rotate(pipeline0, Math.PI - degree);
                                        }

                                    }

                                    CT.ConnectElementWithList(doc, doc.GetElement(ref0), selectedElements);


                                    trans.Commit();
                                }///a
                            }


                        //}
                        //catch { }
                    }
                    else
                    {
                        break;
                    }

                }



                #endregion

                #region Test

                //Reference r1 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                //Element element = doc.GetElement(r1);

                //Reference r2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                //Element element2 = doc.GetElement(r2);
                ////a
                ////CT.ReconnectElementsInList(doc, selectedElements);
                ////foreach (Element element in selectedElements)
                ////{
                ////MessageBox.Show($"{element}");
                ////MessageBox.Show($"{element2}");

                //CT.GetTwoConnectors(doc, element, element2);//a
                ////}

                #endregion

            //}
            //catch
            //{ }
            return Result.Succeeded;
        }
        public static void HighlightEachPair(UIDocument uidoc, Hashtable connectionMap)
        {
            Document doc = uidoc.Document;

            foreach (DictionaryEntry entry in connectionMap)
            {
                Connector conn1 = GetConnectorFromKey(doc, entry.Key.ToString());
                Connector conn2 = entry.Value as Connector;

                if (conn1 != null && conn2 != null)
                {
                    ElementId id1 = conn1.Owner.Id;
                    ElementId id2 = conn2.Owner.Id;

                    uidoc.Selection.SetElementIds(new List<ElementId> { id1, id2 });

                    TaskDialog.Show("Kết nối", $"Đã từng kết nối:\n - ID {id1}\n - ID {id2}");
                }
            }
        }
        private static Connector GetConnectorFromKey(Document doc, string key)
        {
            string[] parts = key.Split('_');
            if (parts.Length != 4) return null;

            int ownerId = int.Parse(parts[0]);
            double x = double.Parse(parts[1]);
            double y = double.Parse(parts[2]);
            double z = double.Parse(parts[3]);

            Element owner = doc.GetElement(new ElementId(ownerId));
            if (owner is FamilyInstance fi && fi.MEPModel != null)
            {
                foreach (Connector c in fi.MEPModel.ConnectorManager.Connectors)
                {
                    if (c.Origin.DistanceTo(new XYZ(x, y, z)) < 1e-4)
                        return c;
                }
            }
            return null;
        }
        public static void HighlightConnectionMap(UIDocument uidoc, Document doc, Hashtable connectionMap)
        {
            foreach (DictionaryEntry entry in connectionMap)
            {
                // Lấy connector đầu (từ key)
                Connector conn1 = GetConnectorFromKey(doc, entry.Key.ToString());

                // Lấy connector đã từng kết nối
                Connector conn2 = entry.Value as Connector;

                if (conn1 != null && conn2 != null)
                {
                    ElementId id1 = conn1.Owner.Id;
                    ElementId id2 = conn2.Owner.Id;

                    // Highlight cả hai đối tượng
                    uidoc.Selection.SetElementIds(new List<ElementId> { id1, id2 });

                    // Hiển thị thông tin từng cặp
                    TaskDialog.Show("Kết nối cũ", $"Connector từ Element {id1}\nđã từng kết nối với Element {id2}");
                }
            }
        }
    }
}
