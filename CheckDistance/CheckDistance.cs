

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System;
using System.Linq;
using System.Collections.Generic;

using Dnbim_Tool.CheckDistance;

using System.Collections;


using Dnbim_Tool;
namespace DnBim_Tool
{
	[Transaction(TransactionMode.Manual)]
	public class InterferenceCheckWithSectionBox : IExternalCommand
	{
        public class DataContainer
        {
            public string Data { get; set; } // Dữ liệu bạn muốn truyền
                                             // Thêm các thuộc tính khác nếu cần
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Hashtable roomFamilyInstances = CT.GetFamilyInstancesInRooms(doc);
            List<Hashtable> data = new List<Hashtable>();
            data.Add(roomFamilyInstances);
            var window = new CheckDistanceView(uidoc,doc, data);
            //CT.OpenOrCreate3DView(doc, uidoc);
            window.Show();
            return Result.Succeeded;
            
		}

       
    }
}

