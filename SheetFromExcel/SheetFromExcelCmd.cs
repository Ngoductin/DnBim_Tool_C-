using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Dnbim_Tool;
using DnBim_Tool;
namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class SheetFromExcelCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            var window = new SheetFromExcelView(doc);
            window.ShowDialog();

            return Result.Succeeded;
        }
    }
}
