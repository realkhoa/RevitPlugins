using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DIMAIO
{
    public class DimSelectionFilter : ISelectionFilter
    {
        private Document _doc;
        public DimSelectionFilter(Document doc) { _doc = doc; }
        public bool AllowElement(Element elem) => true;
        public bool AllowReference(Reference reference, XYZ position) => true;
    }
}
