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
    public class CloseEvent : IExternalEventHandler

    {
        
        public CheckDistanceView window { get; set; }
        public void Execute(UIApplication app)
        {
            //A
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;
            CT.ClearHighlight(uidoc,doc);
            
        }
         public string GetName()
        {
            return "My Event Handler";
        }
    }
}

