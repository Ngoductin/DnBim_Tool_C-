using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Linq;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class ReplaceDimensionText : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            var dims = uidoc.Selection.GetElementIds()
                .Select(id => doc.GetElement(id))
                .OfType<Dimension>()
                .ToList();

            if (!dims.Any())
            {
                TaskDialog.Show("Thông báo", "Hãy chọn ít nhất một Dimension (góc hoặc dài).");
                return Result.Cancelled;
            }

            int countChanged = 0;

            using (Transaction t = new Transaction(doc, "Auto Replace Dimension Text"))
            {
                t.Start();

                foreach (var dim in dims)
                {
                    // Lấy chuỗi hiển thị hiện tại (ví dụ "45°" hoặc "90°")
                    string valueStr = dim.ValueString?.Trim();

                    if (string.IsNullOrEmpty(valueStr))
                        continue;

                    // Gỡ ký tự độ ra (nếu có)
                    valueStr = valueStr.Replace("°", "").Trim();

                    // Cố gắng parse sang double
                    if (double.TryParse(valueStr, out double angle))
                    {
                        string newText = null;

                        // Quy tắc thay thế
                        if (angle == 45)
                            newText = "120";
                        else if (angle == 90||angle==89)
                            newText = "245";

                        if (newText != null)
                        {
                            dim.ValueOverride = newText;
                            countChanged++;
                        }
                    }
                }

                t.Commit();
            }

            //TaskDialog.Show("Hoàn tất", $"Đã thay {countChanged} dimension theo quy tắc.");
            return Result.Succeeded;
        }
    }
}
