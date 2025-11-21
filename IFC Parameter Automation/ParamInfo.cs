// ParamInfo.cs
using Autodesk.Revit.DB;

namespace IFC_Parameter_Automation
{
    public class ParamInfo
    {
        public string Code { get; set; }
        public string DataType { get; set; }
        public string Description { get; set; }
        public Category Category { get; set; }
    }
}
