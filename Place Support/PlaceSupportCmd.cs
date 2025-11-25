using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Visual;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.SqlServer.Server;
using DnBim_Tool.Place_Support;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using static Autodesk.Revit.DB.SpecTypeId;
using Reference = Autodesk.Revit.DB.Reference;
using Autodesk.Revit.DB.Structure;

//using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
//using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement.Status;


namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]


    public class PlaceSupportCmd : IExternalCommand
    {


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {



            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;


            Reference pickFace = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Hãy chọn một mặt từ file liên kết");
            XYZ p1 = pickFace.GlobalPoint;
            Plane infinitePlane =null;
            try
            {



                RevitLinkInstance linkInstance = doc.GetElement(pickFace.ElementId) as RevitLinkInstance;
                Document linkdoc = linkInstance.GetLinkDocument();
                Element ele = linkdoc.GetElement(pickFace.LinkedElementId);
                PlanarFace face = CT.FindPlanarFace(ele, p1);
                XYZ normal = face.FaceNormal; // Vector pháp tuyến của mặt phẳng
                                              // Tạo mặt phẳng vô hạn
                 infinitePlane = Plane.CreateByNormalAndOrigin(normal, p1);
            }
            catch { 
                Reference pickFace2 = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Chọn mặt phẳng thất bại vui lòng chọn điểm 2");
                Reference pickFace3 = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Chọn mặt phẳng thất bại vui lòng chọn điểm 3");

                XYZ p2 = pickFace2.GlobalPoint;
                XYZ p3 = pickFace3.GlobalPoint;
                infinitePlane = Plane.CreateByThreePoints(p1, p2, p3);
                //a
            }
            

            bool Auto;



            while (true)
            {
                try
                {
                    var window = new PlaceSupportView(doc);
                window.ShowDialog();
                string NameTabCtrl = window.tentabcontrol;
                if (window.DialogResult == true && NameTabCtrl == "Pipe")
                {

                    double distanceStart = window.Distance / 304.8;
                    double distanceOffset = window.Offset / 304.8;
                    string Mode = window.cbbOption.SelectedValue.ToString();

                    string familyName = window.cbbFamily.SelectedValue.ToString();
                    string typeName = window.cbbType.SelectedValue.ToString();

                    if (Mode == "Auto")
                    { Auto = true; }
                    else { Auto = false; }



                    switch (familyName)
                    {
                        case string name when name.IndexOf("Pipe_Normal_Hanger_R22", StringComparison.OrdinalIgnoreCase) >= 0:
                            {
                                if (Auto == true)
                                {
                                    PipeUltis.P_Cum_Auto(uidoc, doc, infinitePlane, distanceStart, distanceOffset, familyName, typeName);
                                }
                                else
                                {
                                    PipeUltis.P_Cum_Manual(uidoc, doc, infinitePlane, distanceStart, distanceOffset, familyName, typeName, Auto);
                                }
                                break; // Dừng hàm tại đây
                            }




                        default:
                            {
                                if (Auto == true)
                                {
                                    PipeUltis.P_ThanhUGiaDo_CumU_Auto(uidoc, doc, infinitePlane, distanceStart, distanceOffset, familyName, typeName);
                                }
                                else
                                {
                                    PipeUltis.P_ThanhUGiaDo_CumU_Manual(uidoc, doc, infinitePlane, distanceStart, distanceOffset, familyName, typeName, Auto);
                                }
                                break; // Dừng hàm tại đây
                            }
                    }


                }

                else if (window.DialogResult == true && NameTabCtrl == "Duct")
                {

                    double distanceStart = window.Distance / 304.8;
                    double distanceOffset = window.Offset / 304.8;
                    string Mode = window.cbbOption.SelectedValue.ToString();

                    string familyName = window.cbbFamily.SelectedValue.ToString();
                    string typeName = window.cbbType.SelectedValue.ToString();

                    if (Mode == "Auto")
                    { Auto = true; }
                    else { Auto = false; }


                    switch (familyName)
                    {
                        case string name when name.IndexOf("HT_Unistrut_Hanger_For_Elc_Hanger", StringComparison.OrdinalIgnoreCase) >= 0:
                            {
                                if (Auto == true)
                                {
                                      
                                    DuctUltis.D_ThanhUGiaDo_Auto(uidoc, doc, infinitePlane, distanceStart, distanceOffset, familyName, typeName);
                                }
                                else
                                {
                                    DuctUltis.D_ThanhUGiaDo_Manual(uidoc, doc, infinitePlane, distanceStart, distanceOffset, familyName, typeName, Auto);
                                }
                                break; // Dừng hàm tại đây
                            }




                        default:
                            {
                                break; // Dừng hàm tại đây
                            }
                    }
                }
                else if (window.DialogResult == true && NameTabCtrl == "Cable Tray")
                {
                    double distanceStart = window.Distance / 304.8;
                    double distanceOffset = window.Offset / 304.8;
                    string Mode = window.cbbOption.SelectedValue.ToString();

                    string familyName = window.cbbFamily.SelectedValue.ToString();
                    string typeName = window.cbbType.SelectedValue.ToString();

                    if (Mode == "Auto")
                    { Auto = true; }
                    else { Auto = false; }



                    switch (familyName)
                    {
                        case string name when name.IndexOf("HT_Unistrut_Hanger_For_Elc_Hanger", StringComparison.OrdinalIgnoreCase) >= 0:
                            {
                                if (Auto == true)
                                {
                                    CabletTrayUltis.C_ThanhUGiaDo_Auto(uidoc, doc, infinitePlane, distanceStart, distanceOffset, familyName, typeName);
                                }
                                else
                                {
                                    CabletTrayUltis.C_ThanhUGiaDo_Manual(uidoc, doc, infinitePlane, distanceStart, distanceOffset, familyName, typeName, Auto);
                                }
                                break; // Dừng hàm tại đây
                            }

                                //a


                        default:
                            {
                                break; // Dừng hàm tại đây
                            }
                    }
                }    
                else
                {
                    break;
                }
                }
                catch { }

            }





            return Result.Succeeded;

        }

    }
}




