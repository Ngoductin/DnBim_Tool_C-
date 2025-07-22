//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text.RegularExpressions;
//using System.Windows;
//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.UI;
//using Autodesk.Revit.UI.Selection;
//using DnBim_Tool;

//namespace Dnbim_Tool
//{
//    [Transaction(TransactionMode.Manual)]
//    public class AlginNhaptop : IExternalCommand
//    {
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            UIApplication application = commandData.Application;
//            Document document = application.ActiveUIDocument.Document;
//            UIDocument activeUiDocument = application.ActiveUIDocument;
//            View activeView = document.ActiveView;
//            XYZ rightDirection = activeView.RightDirection;
//            XYZ upDirection = activeView.UpDirection;

//            ICollection<ElementId> elementIds = application.ActiveUIDocument.Selection.GetElementIds();
//            IList<Element> pickedElements = (IList<Element>)new List<Element>();
//            ISelectionFilter selFilter = (ISelectionFilter)new AlignFilter();
//            Transaction transaction = new Transaction(document);
//            MessageBox.Show(elementIds.Count.ToString());   
//            IList<Element> elementList = CT.Align_CurrentSelection(document, activeUiDocument, elementIds, pickedElements, selFilter);

//            if (elementList.Count <= 1)
//                return (Result)0;

//            transaction.Start("HG_Align_To_Top");
//            Plane plane = CT.Set_WorkPlane(document);
//            List<XYZ> xyzList = new List<XYZ>();
//            List<double> source = new List<double>();
//            foreach (Element element in (IEnumerable<Element>)elementList)
//            {
//                TextNote textNote = (TextNote)null;
//                IndependentTag independentTag = (IndependentTag)null;
//                if (element.GetType() == typeof(TextNote))
//                    textNote = element as TextNote;
//                else if (element.GetType() == typeof(IndependentTag))
//                    independentTag = element as IndependentTag;
//                if (element.Category.IsTagCategory)
//                {
//                    XYZ tagHeadPosition = independentTag.TagHeadPosition;
//                    XYZ xyz = CT.ProjectOnto(plane, tagHeadPosition);
//                    xyzList.Add(xyz);
//                    source.Add(xyz.X * upDirection.X + xyz.Y * upDirection.Y + xyz.Z * upDirection.Z);
//                }
//                else if (element.Category.Id.IntegerValue == -2000300)
//                {
//                    element.get_Parameter((BuiltInParameter)(-1006309));
//                    //element.get_Parameter(BuiltInParameter.TEXT_ALIGN_VERT);
//                    XYZ coord = ((TextElement)textNote).Coord;
//                    XYZ xyz = CT.ProjectOnto(plane, coord);
//                    xyzList.Add(xyz);
//                    source.Add(xyz.X * upDirection.X + xyz.Y * upDirection.Y + xyz.Z * upDirection.Z);
//                }
//                else if (element.Location.GetType() == typeof(LocationPoint))
//                {
//                    XYZ xyz = CT.ProjectOnto(plane, (element.Location as LocationPoint).Point);
//                    xyzList.Add(xyz);
//                    source.Add(xyz.X * upDirection.X + xyz.Y * upDirection.Y + xyz.Z * upDirection.Z);
//                }
//                else if (element.Location.GetType() == typeof(LocationCurve))
//                {
//                    try
//                    {
//                        Curve curve = (element.Location as LocationCurve).Curve;
//                        XYZ p = XYZ.op_Division(XYZ.op_Addition(curve.GetEndPoint(0), curve.GetEndPoint(1)), 2.0);
//                        XYZ xyz = Methods.ProjectOnto(plane, p);
//                        xyzList.Add(xyz);
//                        source.Add(xyz.X * upDirection.X + xyz.Y * upDirection.Y + xyz.Z * upDirection.Z);
//                    }
//                    catch
//                    {
//                    }
//                }
//            }
//            XYZ Pt = xyzList[source.IndexOf(source.Max())];
//            foreach (Element ele in (IEnumerable<Element>)elementList)
//                CT.Align_To_Hor_Pt(ele, activeView, Pt, document, plane, rightDirection, upDirection);
//            transaction.Commit();
//            return (Result)0;
//        }

//    }
//    }

//}

///// <summary>
///// Command để thực thi alignment
///// </summary>

