using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DnBim_Tool;
using Dnbim_Tool.CheckDistance;

namespace Dnbim_Tool
{
    public class CheckDistanceEventHandler : IExternalEventHandler
    {
        private ElementId fi1Id;
        private ElementId fi2Id;
        private Document doc;
        private UIDocument uiDoc;

        public CheckDistanceView window { get; set; }
        public string Data { get; set; }

        public void SetData(Document doc, UIDocument uiDoc, ElementId fi1Id, ElementId fi2Id = null)
        {
            this.doc = doc;
            this.uiDoc = uiDoc;
            this.fi1Id = fi1Id;
            this.fi2Id = fi2Id;
        }

        public void Execute(UIApplication app)
        {
            using (Transaction tx = new Transaction(doc, "Highlight and Section Box"))
            {
                tx.Start();

                // Highlight elements
                Element e1 = fi1Id != null ? doc.GetElement(fi1Id) : null;
                Element e2 = fi2Id != null ? doc.GetElement(fi2Id) : null;

                if (e1 != null)
                {
                    CT.ApplyInterferenceHighlight(doc, uiDoc, fi1Id);
                }
                if (e2 != null)
                {
                    CT.ApplyInterferenceHighlight(doc, uiDoc, fi2Id);
                }

                // Tạo section box nếu có cả e1 và e2 (tùy chọn, hiện đang comment)
                /*
                if (e1 != null && e2 != null)
                {
                    View3D view3D = GetOrCreate3DView(doc);
                    if (view3D != null)
                    {
                        CreateTightSectionBox(uiDoc, doc, view3D, e1, e2);
                        uiDoc.ActiveView = view3D;
                    }
                }
                */

                tx.Commit();
            }
        }

        public string GetName()
        {
            return "Check Distance Handler";
        }
    }
    }
