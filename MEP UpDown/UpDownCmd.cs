using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Dnbim_Tool.MEP_UpDown;
using DnBim_Tool;
using DnBim_Tool.Place_Support;


namespace DnBim_Tool
{
    // Giấy phép để truy cập revit
    [Transaction(TransactionMode.Manual)]
    public class UpDownCmd : IExternalCommand
    {
        private static ExternalEvent _externalEvent;
        private static DuctEvent _ductEvent;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view=doc.ActiveView;
            //a
            while (true)
            {
                //a
                var window = new MEPUpDownView();
                window.ShowDialog();
                //string NameTabCtrl = window.tentabcontrol;

                if (window.DialogResult == true)
                {
                   

                      //  try
                      //{

                        // Chọn đối tượng trước khi gọi ExternalEvent
                        Reference r1 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new DuctsPipesCableTraysSelectionFilter(), "Pick first point");
                        Reference r2 = null;
                        if (doc.GetElement(r1) is Duct)
                        {
                            r2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new DuctsSelectionFilter(), "Pick second point");
                        }
                        if (doc.GetElement(r1) is Pipe)
                        {
                            r2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new PipeFilter(), "Pick second point");
                        }
                        if (doc.GetElement(r1) is CableTray)
                        {
                            r2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new CableTrayFilter(), "Pick second point");
                        }
                        if (doc.GetElement(r1) is Conduit)
                        {
                            r2 = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, new ConduitFilter(), "Pick second point");
                        }


                        string option = window.LastSelectedOption;
                        string angle = window.LastAngle;
                        double offset = window.LastOffset / 304.8;
                        
                        switch (option)
                        {
                            case "Cut Up":
                                {
                                    if (angle == "45°") { Updownultis.DuctCut45(doc, r1, r2, offset, true); }

                                    else Updownultis.DuctCut90(doc, r1, r2, offset, true);

                                    break;
                                }

                            case "Cut Down":
                                {
                                    if (angle == "45°") Updownultis.DuctCut45(doc, r1, r2, offset, false);

                                    else Updownultis.DuctCut90(doc, r1, r2, offset, false);

                                    break;
                                }
                            case "Move Up":
                                {
                                    if (angle == "45°") Updownultis.DuctMove45(doc, r1, r2, offset, true);

                                    else Updownultis.DuctMove90(doc, r1, r2, offset, true);

                                    break;

                                }
                            case "Move Down":
                                {
                                    if (angle == "45°") Updownultis.DuctMove45(doc, r1, r2, offset, false);
                                    else Updownultis.DuctMove90(doc, r1, r2, offset, false);
                                    break;

                                }
                            default:
                                {
                                    break;
                                }

                        }
                    //}


                    //catch
                    //{ }
                }
                else
                {
break;
                }
            }
               
              

       












            return Result.Succeeded;
        }

    }
}
