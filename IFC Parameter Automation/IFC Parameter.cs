using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;



namespace IFC_Parameter_Automation
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class IFCParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the current document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            // Your code to automate IFC parameters goes here

            UserControl1 win = new UserControl1(uidoc);
            win.ShowDialog();
            return Result.Succeeded;
        }
    }


}
