using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class CREATDETAILGAS : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // --- Bước 1: Chọn vùng chứa ống và fitting ---
                IList<Element> selectedElements = uidoc.Selection.PickElementsByRectangle("Chọn vùng chứa ống và fitting")
                    .Where(e =>
                        e is Pipe ||
                        (e is FamilyInstance fi && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting))
                    .ToList();

                if (selectedElements.Count == 0)
                {
                    TaskDialog.Show("Thông báo", "Không có đối tượng nào được chọn!");
                    return Result.Cancelled;
                }

                // --- Bước 2: Chọn điểm đặt Family Detail ---
                //XYZ pickPoint = uidoc.Selection.PickPoint("Chọn điểm để đặt Family Detail");

                double angleDeg = 0;
                string familyName = "";
                double lengthMM = 0;
                string sizeongdong = "";
                double maxDetail = DetailUtils.TIMGIATRIDETAILMAX(doc, "CHI TIẾT DETAIL");
                double nextValue = maxDetail + 1;
                //MessageBox.Show(nextValue.ToString());

                // --- Bước 3: Lấy thông tin góc từ fitting ---
                foreach (Element e in selectedElements)
                {
                    if (e is FamilyInstance fi &&
                        fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeFitting)
                    {
                        // Tìm parameter "Angle" (hoặc "Bend Angle")
                        Parameter angleParam =
                            fi.LookupParameter("Angle") ??
                            fi.LookupParameter("ANGLE") ??
                            fi.Symbol.LookupParameter("Angle") ??
                            fi.Symbol.LookupParameter("ANGLE") ??
                            fi.LookupParameter("Bend Angle") ??
                            fi.Symbol.LookupParameter("Bend Angle");
                       

                        if (angleParam != null)
                        {
                            double angleRad = angleParam.AsDouble();
                            angleDeg = angleRad * (180 / Math.PI);
                           
                        }
                    }
                    if( e is Pipe)
                    {
                        if (e is Pipe)
                        {
                            Parameter lenParam = e.LookupParameter("Length");
                            if (lenParam != null && lenParam.StorageType == StorageType.Double)
                            {
                                double lengthFeet = lenParam.AsDouble(); // giá trị gốc trong feet


                                lengthMM = DetailUtils.RoundFeetToNearest5mm(lengthFeet);

                                //TaskDialog.Show("Chiều dài ống", $"Length = {lengthMM:F0} mm");
                            }
                            Parameter kichthuoc = e.LookupParameter("Đường kính óng gas (hơi)");
                           sizeongdong = kichthuoc.AsString();
                            
                        }
                    }
                }

                if (angleDeg == 0)
                {
                    TaskDialog.Show("Thông báo", "Không tìm thấy parameter 'Angle' trong fitting nào trong vùng chọn.");
                    return Result.Cancelled;
                }

                // --- Bước 4: Xác định family cần chèn ---
                if (angleDeg<55)
                    familyName = "2 CO 45 ỐNG ĐỒNG";
                else if (angleDeg>55)
                    familyName = "2 CO 90 ỐNG ĐỒNG";
                else
                {
                    TaskDialog.Show("Thông báo", $"Không có family tương ứng với góc {angleDeg:F1}°");
                    return Result.Cancelled;
                }

            if (angleDeg > 55)
                {
                    familyName = "2 CO 90 ỐNG ĐỒNG";
                }
                   
            else
                {
                    familyName = "2 CO 45 ỐNG ĐỒNG";
                }

                // 🔹 Tìm FamilySymbol thuộc family đó (Generic Annotation)
                FamilySymbol symbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(s =>
                        s.Family.Name.Equals(familyName, System.StringComparison.OrdinalIgnoreCase) &&
                        s.Family.FamilyCategory != null &&
                        s.Family.FamilyCategory.Name == "Generic Annotations");

                // 🔹 Nếu không tìm thấy thì báo lỗi
                if (symbol == null)
                {
                    string allSymbols = string.Join("\n",
                        new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .Where(s => s.Family.FamilyCategory.Name == "Generic Annotations")
                        .Select(s => $"{s.Family.Name} - {s.Name}"));

                    TaskDialog.Show("Không tìm thấy",
                        $"Không tìm thấy Family '{familyName}' trong project.\n\nCác Generic Annotation hiện có:\n{allSymbols}");
                    return Result.Failed;
                }

                // 🔹 Chọn điểm để đặt family
                XYZ point = uidoc.Selection.PickPoint("Chọn vị trí để đặt 2 CO 90 ỐNG ĐỒNG");

                using (Transaction t = new Transaction(doc, "Chèn 2 CO 90 ỐNG ĐỒNG"))
                {
                    t.Start();
                    foreach (Element e in selectedElements)
                    {
                        Parameter detailNumber = e.LookupParameter("CHI TIẾT DETAIL");
                        if (detailNumber != null)
                        {
                            if (detailNumber.StorageType == StorageType.String)
                            {
                                // 🔹 Convert nextValue (double/int) sang chuỗi trước khi gán
                                detailNumber.Set(nextValue.ToString());
                            }
                            else if (detailNumber.StorageType == StorageType.Double)
                            {
                                // 🔹 Nếu là số, có thể gán trực tiếp
                                detailNumber.Set(nextValue);
                            }
                            else if (detailNumber.StorageType == StorageType.Integer)
                            {
                                detailNumber.Set((int)nextValue);
                            }
                        }
                    }

                    // 🔹 Đặt family vào View hiện tại (Generic Annotation = 2D)
                    // 🔹 Chèn family vào view (Generic Annotation là 2D)
                    FamilyInstance fiPlaced = doc.Create.NewFamilyInstance(point, symbol, uidoc.ActiveView);

                    // 🔹 Tìm parameter KT2 và gán giá trị chiều dài ống (mm)
                    Parameter kt2Param = fiPlaced.LookupParameter("KT3");
                    if (kt2Param != null)
                    {
                        if (kt2Param.StorageType == StorageType.Double)
                        {
                            // Đổi đơn vị mm → feet trước khi set (vì Revit lưu nội bộ theo feet)
                            double valFeet = lengthMM;
                            kt2Param.Set(valFeet);
                        }
                      
                    }
                    Parameter KTbox = fiPlaced.LookupParameter("SIZE ỐNG ĐỒNG");
                    if (kt2Param != null)
                    {
                        if (KTbox.StorageType == StorageType.String)
                        {
                            // Đổi đơn vị mm → feet trước khi set (vì Revit lưu nội bộ theo feet)
                           KTbox.Set(sizeongdong);
                        }

                    }
                    Parameter STT = fiPlaced.LookupParameter("STT");
                    if (STT != null)
                    {
                        if (STT.StorageType == StorageType.String)
                        {
                            // Đổi đơn vị mm → feet trước khi set (vì Revit lưu nội bộ theo feet)
                            STT.Set(nextValue.ToString());
                        }

                    }


                    t.Commit();
                }

               
                return Result.Succeeded;



            }

            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi", ex.Message);
                return Result.Failed;
            }
        }

    }
}


