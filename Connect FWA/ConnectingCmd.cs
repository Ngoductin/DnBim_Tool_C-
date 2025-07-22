using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class ConnectingCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                Reference ref1 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new filterFamilyInstance(), "Chọn thiết bị 1 (cố định)");
                Reference ref2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new filterFamilyInstance(), "Chọn thiết bị 2 (di chuyển)");

                FamilyInstance inst1 = doc.GetElement(ref1) as FamilyInstance;
                FamilyInstance inst2 = doc.GetElement(ref2) as FamilyInstance;

                Connector conn1 = null;
                Connector conn2 = null;
                Element mepElementToDelete = null;

                foreach (Connector c1 in GetConnectors(inst1))
                {
                    foreach (Connector refC1 in c1.AllRefs)
                    {
                        Element owner = refC1.Owner;

                        if (owner is MEPCurve || owner is CableTray) // MEPCurve gồm Pipe, Duct, Conduit
                        {
                            foreach (Connector c2 in GetConnectors(inst2))
                            {
                                foreach (Connector refC2 in c2.AllRefs)
                                {
                                    if (refC2.Owner.Id == owner.Id)
                                    {
                                        // Tìm thấy đối tượng trung gian chung
                                        mepElementToDelete = owner;
                                        conn1 = c1;
                                        conn2 = c2;
                                        break;
                                    }
                                }
                                if (mepElementToDelete != null) break;
                            }
                        }

                        if (mepElementToDelete != null) break;
                    }
                    if (mepElementToDelete != null) break;
                }

                if (mepElementToDelete == null || conn1 == null || conn2 == null)
                {
                    TaskDialog.Show("Kết quả", "Không tìm thấy đoạn nối chung giữa hai thiết bị.");
                    return Result.Failed;
                }

                using (Transaction trans = new Transaction(doc, "Kết nối lại thiết bị"))
                {
                    trans.Start();

                    // Xóa đoạn ống / ống gió / conduit / cable tray
                    doc.Delete(mepElementToDelete.Id);

                    // Di chuyển thiết bị 2 về gần thiết bị 1
                    XYZ moveVec = conn1.Origin - conn2.Origin;
                    ElementTransformUtils.MoveElement(doc, inst2.Id, moveVec);              
                        
                    // Kết nối trực tiếp 2 connector
                    conn1.ConnectTo(conn2);
                    //a
                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Lỗi", ex.Message);
                return Result.Failed;
            }
        }

        private List<Connector> GetConnectors(FamilyInstance fi)
        {
            List<Connector> connectors = new List<Connector>();

            MEPModel mep = fi.MEPModel;
            if (mep != null)
            {
                ConnectorSet set = mep.ConnectorManager.Connectors;
                foreach (Connector c in set)
                {
                    connectors.Add(c);
                }
            }

            return connectors;
        }
    }
}
