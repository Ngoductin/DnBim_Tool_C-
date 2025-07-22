using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using DnBim_Tool.Rotate_3D;
//using Syncfusion.UI.Xaml.Grid;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class CalculateLength : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            // Đổi tên biến để tránh trùng với tham số "elements"
            IList<Element> selectedElements = new List<Element>();

            // Lấy các Element được chọn
            ICollection<ElementId> selectedElementIds = uidoc.Selection.GetElementIds();
            // Sử dụng một danh sách để lưu trữ các Duct trong vòng đầu tiên (đa luồng)
            List<Duct> allDucts = new List<Duct>();
            // Khởi tạo biến để lưu trữ tổng độ dài
            //double totalLength = 0;
            IList<Pipe> allPipes = new List<Pipe>();
            IList<Conduit> allConduits= new List<Conduit>();
            IList<CableTray> allCableTrays= new List<CableTray>();

            try
            {

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
                    //MessageBox.Show("Tôi bị ngu");
                }




               

                    if (selectedElements.Count == 0)
                    {
                        
                    }
                // Chạy đa luồng để thu thập các đối tượng Duct, Pipe, Conduit, Cable Tray
                Parallel.ForEach(selectedElements.Cast<Element>(), (element) =>
                {
                    // Kiểm tra và xử lý nếu là Duct, Pipe, Conduit hoặc Cable Tray
                    if (element is Duct duct)
                    {
                        lock (allDucts)
                        {
                            allDucts.Add(duct); // Thêm vào danh sách nếu là Duct
                        }
                    }
                    else if (element is Pipe pipe)
                    {
                        lock (allPipes)
                        {
                            allPipes.Add(pipe); // Thêm vào danh sách nếu là Pipe
                        }
                    }
                    else if (element is Conduit conduit)
                    {
                        lock (allConduits)
                        {
                            allConduits.Add(conduit); // Thêm vào danh sách nếu là Conduit
                        }
                    }
                    else if (element is CableTray cableTray)
                    {
                        lock (allCableTrays)
                        {
                            allCableTrays.Add(cableTray); // Thêm vào danh sách nếu là Cable Tray
                        }
                    }
                });

                // Dùng HashSet để lọc các đối tượng trùng lặp theo ID
                HashSet<int> processedIds = new HashSet<int>();
                double totalLength = 0;

                // Duyệt qua danh sách Duct, Pipe, Conduit, Cable Tray và tính tổng độ dài
                foreach (var duct in allDucts)
                {
                    if (!processedIds.Contains(duct.Id.IntegerValue))
                    {
                        processedIds.Add(duct.Id.IntegerValue); // Thêm ID vào HashSet để tránh trùng lặp
                                                                // Lấy parameter "Length" của Duct
                        Parameter lengthParam = duct.LookupParameter("Length");

                        if (lengthParam != null && lengthParam.HasValue)
                        {
                            totalLength += lengthParam.AsDouble(); // Cộng dồn độ dài
                        }
                    }
                }

                foreach (var pipe in allPipes)
                {
                    if (!processedIds.Contains(pipe.Id.IntegerValue))
                    {
                        processedIds.Add(pipe.Id.IntegerValue); // Thêm ID vào HashSet để tránh trùng lặp
                                                                // Lấy parameter "Length" của Pipe
                        Parameter lengthParam = pipe.LookupParameter("Length");

                        if (lengthParam != null && lengthParam.HasValue)
                        {
                            totalLength += lengthParam.AsDouble(); // Cộng dồn độ dài
                        }
                    }
                }

                foreach (var conduit in allConduits)
                {
                    if (!processedIds.Contains(conduit.Id.IntegerValue))
                    {
                        processedIds.Add(conduit.Id.IntegerValue); // Thêm ID vào HashSet để tránh trùng lặp
                                                                   // Lấy parameter "Length" của Conduit
                        Parameter lengthParam = conduit.LookupParameter("Length");

                        if (lengthParam != null && lengthParam.HasValue)
                        {
                            totalLength += lengthParam.AsDouble(); // Cộng dồn độ dài
                        }
                    }
                }

                foreach (var cableTray in allCableTrays)
                {
                    if (!processedIds.Contains(cableTray.Id.IntegerValue))
                    {
                        processedIds.Add(cableTray.Id.IntegerValue); // Thêm ID vào HashSet để tránh trùng lặp
                                                                     // Lấy parameter "Length" của Cable Tray
                        Parameter lengthParam = cableTray.LookupParameter("Length");

                        if (lengthParam != null && lengthParam.HasValue)
                        {
                            totalLength += lengthParam.AsDouble(); // Cộng dồn độ dài
                        }
                    }
                }
                MessageBox.Show("Độ dài tổng đường ống là " + (totalLength * 0.3048).ToString("F2") + " m");
            }
            catch { }
            return Result.Succeeded;
        }
        //a
    }
}