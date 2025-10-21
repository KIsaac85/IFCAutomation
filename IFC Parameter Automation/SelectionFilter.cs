using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;


namespace IFC_Parameter_Automation
{
    /// <summary>
    /// this class is created to restrict the single selection to user
    /// only specific elements are allowed to be selected
    /// the elements which may receive waterproofing 
    /// e.g. columns/walls/foundation
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
            if (elem.Category.Id.Value == (int)BuiltInCategory.OST_StructuralColumns
                || elem.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
            {
                FamilyInstance familyInstance = elem as FamilyInstance;
                if (familyInstance.StructuralMaterialType == StructuralMaterialType.Concrete)
                {
                    return true & null != elem; 
                }
            }
            else if (elem.Category.Id.Value == (int)BuiltInCategory.OST_Walls)
            {
                Wall wall = elem as Wall;
                if (wall.CurtainGrid==null&&wall.StructuralUsage==StructuralWallUsage.Bearing)
                {

                    return true & null != elem;
                }
            }
            else if (elem.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFoundation)
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
            Element eFootingLink = docLink.GetElement(reference.LinkedElementId);
            if (eFootingLink.Category.Id.Value == (int)BuiltInCategory.OST_StructuralColumns
                || eFootingLink.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
            {
                FamilyInstance familyInstance = eFootingLink as FamilyInstance;
                if (familyInstance.StructuralMaterialType == StructuralMaterialType.Concrete)
                {
                    return true && null != eFootingLink;
                }
                
            }
            else if (eFootingLink.Category.Id.Value == (int)BuiltInCategory.OST_Walls)
            {

                Wall wall = eFootingLink as Wall;
                if (wall.CurtainGrid == null && wall.StructuralUsage == StructuralWallUsage.Bearing)
                {

                    return true & null != eFootingLink;
                }

            }
            else if (eFootingLink.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFoundation)
            {
                return true & null != eFootingLink;
            }
            return false;
        }
    }


}
