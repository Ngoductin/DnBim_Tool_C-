using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Dnbim_Tool.CheckDistance;
using DnBim_Tool;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Dnbim_Tool
{
    public class ShowEvent : IExternalEventHandler

    {

        public CheckDistanceView window { get; set; }
        public ElementId Fi1Id { get; set; }
        public ElementId Fi2Id { get; set; }



        public string Data { get; set; }
        public void Execute(UIApplication app)
        {

            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;



            CT.CreateSectionBoxForElements(uidoc, doc, Fi1Id, Fi2Id);

            //using (Transaction t = new Transaction(doc, "Highlight Element"))
            //{
            //    t.Start();
            //    if (Fi2Id != null)
            //    {
            //        CT.ApplyInterferenceHighlight(doc, uidoc, Fi1Id);
            //        CT.ApplyInterferenceHighlight(doc, uidoc, Fi2Id);


            //    }
            //    else
            //    {
            //        CT.ApplyInterferenceHighlight(doc, uidoc, Fi1Id);
            //    }



            //    t.Commit();
            //}






        }





        public string GetName()
        {
            return "My Event Handler";
        }

    }
}
