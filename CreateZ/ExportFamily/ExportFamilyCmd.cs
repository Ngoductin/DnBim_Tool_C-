using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace OPENSOURCE.ExportFamily
{
    [Transaction(TransactionMode.Manual)]
    public class ExportFamilyCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            //chọn folder để lưu các family
            string selectedPath = string.Empty;

            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                selectedPath = dialog.SelectedPath;
            }    

            if (string.IsNullOrEmpty(selectedPath) )
            {
                return Result.Failed;
            }

            //lấy danh sách các family có thể export
            var listFamily = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => f.IsEditable == true)
                .OrderBy(f => f.FamilyCategory.Name)
                .ToList();


            using(ProcessbarView bv = new ProcessbarView(listFamily.Count))
            {
                bv.Show();

                foreach (Family family in listFamily)
                {
                    //kiểm tra docukent của family có null hay không
                    Document familyDoc = doc.EditFamily(family);
                    if (familyDoc == null)
                    {
                        continue;
                    }

                    string familyName = family.Name;
                    string categoryName = family.FamilyCategory.Name;

                    //Tạo thư mục con theo tên Category
                    string categoryFolder = Path.Combine(selectedPath, categoryName);
                    if (!Directory.Exists(categoryFolder))
                    {
                        Directory.CreateDirectory(categoryFolder);
                    }

                    //Tạo đường dẫn đầy đủ cho file RFA
                    string outputPath = Path.Combine(categoryFolder, familyName + ".rfa");
                    ModelPath newPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(outputPath);

                    //export family
                    SaveAsOptions saveOptions = new SaveAsOptions();
                    saveOptions.OverwriteExistingFile = true;

                    familyDoc.SaveAs(newPath, saveOptions);
                    familyDoc.Close(false);

                    if (bv.Update()) break;
                }

            }


            return Result.Succeeded;
        }
    }
}
