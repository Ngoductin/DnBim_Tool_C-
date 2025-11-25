using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Mechanical;
using Microsoft.Office.Interop.Excel;
using Line = Autodesk.Revit.DB.Line;
using DnBim_Tool;

namespace Dnbim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class REHOSTCMD : IExternalCommand
    {
        public Result Execute(ExternalCommandData cdata, ref string message, ElementSet elements)
        {
            UIDocument uidoc = cdata.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                //a
               
                    // --------- CHẾ ĐỘ A: PLACE NEW ----------
                    // 1) Chọn duct/pipe
                    Reference rDuct = uidoc.Selection.PickObject(ObjectType.Element, "Chọn ống/ống gió (duct/pipe)");
                    MEPCurve duct = doc.GetElement(rDuct) as MEPCurve;
                    if (duct == null) { TaskDialog.Show("Lỗi", "Đối tượng không phải MEPCurve."); return Result.Cancelled; }

                    // 2) Chọn một instance mẫu để lấy Symbol (hoặc bạn có thể tự tìm Symbol khác)
                    Reference rSample = uidoc.Selection.PickObject(ObjectType.Element, "Chọn 1 FamilyInstance face-based để LẤY SYMBOL");
                    FamilyInstance fiSample = doc.GetElement(rSample) as FamilyInstance;
                    if (fiSample == null) { TaskDialog.Show("Lỗi", "Đối tượng được chọn không phải FamilyInstance."); return Result.Cancelled; }
                    FamilySymbol sym = fiSample.Symbol;

                    // Kiểm tra face-based
                    bool isFaceBased = fiSample.HostFace != null;   // true nếu instance đang host theo face
                    if (!isFaceBased)
                    {
                        TaskDialog.Show("Lỗi", "Family được chọn không phải Face-Based (hoặc instance mẫu không bám mặt).");
                        return Result.Cancelled;
                    }

                    // 3) Chọn điểm gần vị trí muốn đặt
                    //XYZ nearPt = uidoc.Selection.PickPoint("Click 1 điểm gần vị trí đặt (sẽ được chiếu xuống đáy ống)");
                    XYZ nearPt = rDuct.GlobalPoint;

                    // 4) Tìm mặt dưới & tạo instance mới
                    Face bottomFace = GetBottomFace(doc, duct);
                    if (bottomFace == null) { TaskDialog.Show("Lỗi", "Không tìm được mặt dưới của ống."); return Result.Failed; }

                    var proj = bottomFace.Project(nearPt);
                    if (proj == null) { TaskDialog.Show("Lỗi", "Không project được điểm xuống mặt dưới."); return Result.Failed; }
                    XYZ placePt = proj.XYZPoint;
                    XYZ pick1 = rDuct.GlobalPoint;
                    Element element1 = doc.GetElement(rDuct);
                    Line locationLine = null;
                    if (element1 is Duct duct1)
                    {
                        LocationCurve locationCurve = duct.Location as LocationCurve;
                        locationLine = locationCurve.Curve as Line;
                    }
                    XYZ x = locationLine.GetEndPoint(0);
                    XYZ direction = locationLine.Direction;
                    Plane plane1 = Plane.CreateByNormalAndOrigin(direction, pick1);
                    XYZ gd1 = CT.LineIntersectPlane(locationLine, plane1);
                    placePt = new XYZ(gd1.X, gd1.Y, placePt.Z);

                    XYZ refDir = ComputeRefDirOnFace(bottomFace, duct, proj.UVPoint);

                    using (Transaction t = new Transaction(doc, "TIN – Place support on duct bottom"))
                    {
                        t.Start();
                        if (!sym.IsActive) sym.Activate();

                        Reference hostRef = bottomFace.Reference;
                        FamilyInstance newFi = doc.Create.NewFamilyInstance(hostRef, placePt, refDir, sym);

                        t.Commit();
                    }

                    TaskDialog.Show("OK", "Đã đặt family bám mặt dưới ống.");
                    return Result.Succeeded;
              }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (TargetInvocationException tex)
            {
                message = "Lỗi khi gọi SetHostFace(): " + (tex.InnerException?.Message ?? tex.Message);
                return Result.Failed;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ===== Helpers ===================================================================

        // Tìm mặt “dưới” theo trọng lực (−Z). Hoạt động cho duct chữ nhật & tròn.
        private static Face GetBottomFace(Document doc, Element hostElem)
        {
            var opt = new Options
            {
                ComputeReferences = true,
                IncludeNonVisibleObjects = true,
                DetailLevel = ViewDetailLevel.Fine
            };

            GeometryElement ge = hostElem.get_Geometry(opt);
            Face bestFace = null;
            double best = double.NegativeInfinity;

            foreach (GeometryObject go in ge)
            {
                GeometryElement ge2 = (go as GeometryInstance)?.GetInstanceGeometry() ?? go as GeometryElement;
                var solids = ge2?.OfType<Solid>().Where(s => s != null && s.Volume > 0) ?? Enumerable.Empty<Solid>();
                if (go is Solid s0 && s0.Volume > 0) solids = solids.Append(s0);

                foreach (var s in solids)
                {
                    foreach (Face f in s.Faces)
                    {
                        var bb = f.GetBoundingBox();
                        var mid = new UV((bb.Min.U + bb.Max.U) * 0.5, (bb.Min.V + bb.Max.V) * 0.5);
                        XYZ n = f.ComputeNormal(mid).Normalize();

                        double score = n.DotProduct(-XYZ.BasisZ); // càng hướng xuống càng tốt
                        if (score > best) { best = score; bestFace = f; }
                    }
                }
            }
            return bestFace;
        }

        // Tính refDir (vector nằm trong mặt) – dùng dọc theo ống × pháp tuyến
        private static XYZ ComputeRefDirOnFace(Face face, MEPCurve duct, UV uv)
        {
            XYZ n = face.ComputeNormal(uv).Normalize();

            XYZ ductDir = XYZ.BasisX; // fallback
            var lc = duct.Location as LocationCurve;
            if (lc?.Curve is Line ln) ductDir = ln.Direction.Normalize();

            XYZ refDir = ductDir.CrossProduct(n);
            if (refDir.IsZeroLength())
            {
                if (face is PlanarFace pf)
                    refDir = (pf.XVector.IsZeroLength() ? pf.YVector : pf.XVector).Normalize();
                else
                    refDir = n.CrossProduct(XYZ.BasisX).Normalize();
            }
            return refDir.Normalize();
        }

        private static void MoveInstanceTo(Document doc, FamilyInstance fi, XYZ target)
        {
            var lp = fi.Location as LocationPoint;
            if (lp == null) return;
            XYZ delta = target - lp.Point;
            if (delta.GetLength() > 1e-9)
                ElementTransformUtils.MoveElement(doc, fi.Id, delta);
        }
    }
    }
