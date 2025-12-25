using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;


namespace IFC_Parameter_Automation
{
    /// <summary>
    /// this class is created to restrict the single selection to user
    /// only specific elements are allowed to be selected
    /// e.g. doors/walls/windows
    /// </summary>
    class SelectionFilter : ISelectionFilter
    {

        private Document _doc;

        public Document Doc
        {
            get { return _doc; }
            set { _doc = value; }
        }




        public bool AllowElement(Element elem)
        {
            if (elem.Category.Id.Value == (int)BuiltInCategory.OST_Windows)
            {
                    return true & null != elem;
            }
            else if (elem.Category.Id.Value == (int)BuiltInCategory.OST_Doors)
            {
                return true & null != elem;
            }
            else if (elem.Category.Id.Value == (int)BuiltInCategory.OST_Walls)
            {
                return true & null != elem;
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            //return false;
            RevitLinkInstance revitlinkinstance = _doc.GetElement(reference) as RevitLinkInstance;
            Document docLink = revitlinkinstance.GetLinkDocument();
            Element eLink = docLink.GetElement(reference.LinkedElementId);
            if (eLink.Category.Id.Value == (int)BuiltInCategory.OST_Windows)
            {
               
                return true && null != eLink;
                
                
            }
            else if (eLink.Category.Id.Value == (int)BuiltInCategory.OST_Doors)
            {
                return true & null != eLink;
            }
            else if (eLink.Category.Id.Value == (int)BuiltInCategory.OST_Walls)
            {
                return true & null != eLink;
            }
            return false;
        }
    }


}
