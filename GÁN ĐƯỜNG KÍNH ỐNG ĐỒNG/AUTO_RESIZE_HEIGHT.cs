using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    public class AUTO_RESIZE_HEIGHT : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            using (Transaction t = new Transaction(doc, "AUTO RESIZE HEIGHT"))
            {
                t.Start();
                View activeView = uidoc.ActiveView;
                ViewSchedule schedule = activeView as ViewSchedule;

                TableData table = schedule.GetTableData();
                TableSectionData header = table.GetSectionData(SectionType.Header);

                AutoFitRowHeight(header);

                t.Commit();
            }
            return Result.Succeeded;
        }
    
    public static void AutoFitRowHeight(TableSectionData section)
        {
            for (int r = 0; r < section.NumberOfRows; r++)
            {
                double maxHeight = 3.0; // chiều cao tối thiểu (mm)

                for (int c = 0; c < section.NumberOfColumns; c++)
                {
                    string text = section.GetCellText(r, c);
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    // Ước lượng chiều cao dựa vào số ký tự (15 ký tự / dòng)
                    double lines = Math.Ceiling((double)text.Length / 15.0);
                    double estimatedHeight = 3.5 + lines; // 3mm mỗi dòng

                    if (estimatedHeight > maxHeight)
                        maxHeight = estimatedHeight;
                }

                // Gán chiều cao (đổi sang internal units)
                section.SetRowHeight(r, UnitUtils.ConvertToInternalUnits(maxHeight, UnitTypeId.Millimeters));
            }
        }
    }
    }
