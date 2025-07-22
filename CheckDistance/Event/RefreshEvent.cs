using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Dnbim_Tool.CheckDistance;
using DnBim_Tool;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Dnbim_Tool
{
    public class RefreshEvent : IExternalEventHandler

    {

        public CheckDistanceView window { get; set; }
        public void Execute(UIApplication app)
        {

            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;
            Hashtable roomFamilyInstances = CT.GetFamilyInstancesInRooms(doc);
            

            List<Hashtable> newData = new List<Hashtable> { roomFamilyInstances };

            // Gửi dữ liệu mới tới CheckDistanceView
            if (window != null)
            {
                window.Dispatcher.Invoke(() =>
                {
                    window.UpdateData(newData); // Gọi phương thức UpdateData để cập nhật dữ liệu
                });
            }
        }
        public string GetName()
        {
            return "My Event Handler";
        }
    }
}

