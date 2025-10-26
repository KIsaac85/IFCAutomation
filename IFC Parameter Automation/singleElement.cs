using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IFC_Parameter_Automation
{
    public class singleSelectedElement
    {
        public long elememntID { get; set; }
        public string elementName { get; set; }
        public List<Parameter> IFCparameters { get; set; }

        public singleSelectedElement(Element element)
        {
            elememntID = element.Id.Value;
            elementName = element.Name;
            IFCparameters= element.Parameters.Cast<Parameter>().Where(p => p.Definition.GetGroupTypeId()== GroupTypeId.Ifc).ToList();
            StringBuilder sb = new StringBuilder();
            foreach (Parameter p in IFCparameters) {
                sb.AppendLine($"Parameter Name: {p.Definition.Name}, Value: {p.AsValueString()}");
            }
            TaskDialog.Show("IFC Parameters", sb.ToString());
            //TaskDialog.Show("Info", $"Element ID: {elememntID}\nElement Name: {elementName}\nIFC Parameters Count: {IFCparameters.Count}");
        }
    }
    
}
