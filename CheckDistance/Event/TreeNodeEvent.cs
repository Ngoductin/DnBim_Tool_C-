using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Dnbim_Tool.CheckDistance;
using Dnbim_Tool.MEP_UpDown;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DnBim_Tool;
using System.Windows;
using Autodesk.Revit.DB.ExtensibleStorage;
using System.Windows.Controls;

namespace Dnbim_Tool
{
  
       
    
    public class TreeNodeEvent: IExternalEventHandler

    {

        public CheckDistanceView window { get; set; }
        public ElementId Fi1Id { get; set; }
        public ElementId Fi2Id { get; set; }
      

        
        public string Data { get; set; }
        public void Execute(UIApplication app)
        {

            UIDocument uidoc = app.ActiveUIDocument;
            Document doc= uidoc.Document;








          CT.ClearHighlight(uidoc,doc);


            using (Transaction t = new Transaction(doc, "Highlight Element"))
            {
                t.Start();
                // Kiểm tra xem đã có element nào được highlight trước đó hay chưa
               

                if (Fi2Id != null)
                {
                    CT.ApplyInterferenceHighlight(doc, uidoc, Fi1Id);
                    CT.ApplyInterferenceHighlight(doc, uidoc, Fi2Id);


                }
                else
                {
                    CT.ApplyInterferenceHighlight(doc, uidoc, Fi1Id);
                }
                
              

                t.Commit();
            }



            // Cập nhật các element đã highlight
            window.highlightedFi1Id = Fi1Id;
            window.highlightedFi2Id = Fi2Id;




        }




        
        public string GetName()
        {
            return "My Event Handler";
        }

    }
}
