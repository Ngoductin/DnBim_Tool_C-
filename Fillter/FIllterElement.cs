using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DnBim_Tool
{

    public class ElementbelongtoPipe : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            if (element.Category == null) return false;

            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;

            return builtInCategory == BuiltInCategory.OST_PipeCurves ||
                   builtInCategory == BuiltInCategory.OST_Conduit   ||
                   builtInCategory == BuiltInCategory.OST_PipeAccessory 
                   ;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
    public class PipeFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            if (builtInCategory == BuiltInCategory.OST_PipeCurves)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
    public class Pipe_ConduitFilter : ISelectionFilter
    {
       
    public bool AllowElement(Element element)
        {
            if (element?.Category == null)
                return false;

            BuiltInCategory bic = (BuiltInCategory)element.Category.Id.IntegerValue;

            return bic == BuiltInCategory.OST_PipeCurves ||
                   bic == BuiltInCategory.OST_Conduit ||
                   bic == BuiltInCategory.OST_PipeAccessory;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    
}
    public class DuctFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            if (builtInCategory == BuiltInCategory.OST_DuctCurves)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
    public class CableTrayFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            if (builtInCategory == BuiltInCategory.OST_CableTray)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
    public class ConduitFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            if (builtInCategory == BuiltInCategory.OST_Conduit)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
    public class FloorFaceFilter : ISelectionFilter
    {
        // Phương thức kiểm tra xem đối tượng có phải là sàn hay không
        public bool AllowElement(Element element)
        {
            // Kiểm tra loại của đối tượng có phải là sàn không
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            if (builtInCategory == BuiltInCategory.OST_Floors)
            {
                return true; // Nếu đối tượng là sàn thì cho phép chọn
            }
            return false; // Nếu không phải là sàn, không cho phép chọn
        }

        // Phương thức kiểm tra tham chiếu mặt
        public bool AllowReference(Reference refer, XYZ point)
        {
            // Luôn trả về false vì chúng ta chỉ muốn chọn mặt, không chọn đối tượng.
            return false;
        }
    }
    public class filterFamilyInstance : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            if (element is FamilyInstance) // Kiểm tra nếu phần tử là FamilyInstance
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false; // Không cho phép lọc qua tham chiếu
        }
    }
    public class DuctsSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            if (element.Category.Name == "Ducts")
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
    public class DuctsPipesCableTraysSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            // Kiểm tra nếu đối tượng thuộc vào các danh mục Ducts, Pipes, hoặc Cable Trays
            if (element.Category.Name == "Ducts" || element.Category.Name == "Pipes" || element.Category.Name == "Cable Trays"|| element.Category.Name == "Conduits")
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }
    }
    public class Pipe_Accessories_Filter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            if (builtInCategory == BuiltInCategory.OST_PipeAccessory)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference refer, XYZ point)
        {
            return false;
        }

       
    }
    public class AlignFilter : ISelectionFilter
    {
        public bool AllowElement(Element ele)
        {
            return ele.Category.IsTagCategory ? ele.Category.Id.IntegerValue != -2000280 && ele.Category.Id.IntegerValue != -2000480 && ele.Category.Id.IntegerValue != -2005020 && ele.Category.Id.IntegerValue != 2000485 : ele.Category.Id.IntegerValue == -2000300 || ele.Location.GetType() == typeof(LocationPoint) || ele.Location.GetType() == typeof(LocationCurve);
        }

        public bool AllowReference(Reference refer, XYZ point) => false;
    }



}

