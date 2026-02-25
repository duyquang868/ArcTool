using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace ArcTool.Core.Utils
{
    // Lớp lọc chỉ cho phép chọn Linear Dimension
    public class LinearDimensionSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Dimension dim && dim.Curve is Line;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}