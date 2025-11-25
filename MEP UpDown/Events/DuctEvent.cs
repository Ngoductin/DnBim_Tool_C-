using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Dnbim_Tool.MEP_UpDown;
using DnBim_Tool;

namespace DnBim_Tool
{
    public class DuctEvent : IExternalEventHandler
    {
        public MEPUpDownView window { get; set; }
        public Reference r1 { get; set; }
        public Reference r2 { get; set; }
        public Document doc { get; set; }
        public void Execute(UIApplication app)
        {





        }
                    
                    


                    
            
        
               
               
            
        

        public string GetName()
        {
            return "DuctEvent";
        }
    }
}
