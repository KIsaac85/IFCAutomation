using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;


namespace IFC_Parameter_Automation
{
    public class ExcelExport
    {
        public Document _doc { get; set; }
        private ExcelPackage package { get; set; }
        
        public void ExportParametersToExcel(Document doc)
        {
            ExcelPackage.License.SetNonCommercialPersonal("abc");

            package = new ExcelPackage();
            _doc = doc;
            FilteredElementCollector collector = new FilteredElementCollector(_doc);
            //Create Filter
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Windows);
            IList<Element> elements = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();
            IList<Parameter> IFCParameters = new List<Parameter>();
            StringBuilder sb = new StringBuilder();

            foreach (Element element in elements) {
                IList<Parameter> elementParams = element.Parameters.Cast<Parameter>().Where(p => p.Definition.ParameterGroup == BuiltInParameterGroup.PG_IFC).ToList();
                foreach (Parameter param in elementParams)
                {
                    IFCParameters.Add(param);
                    sb.AppendLine($"{element.Id},{element.Name},{param.Definition.Name},{param.AsString()}");
                }
            }

            //IList<Parameter> IFCParameters = ele.Parameters.Cast<Parameter>().Where(p => p.Definition.ParameterGroup == BuiltInParameterGroup.PG_IFC).ToList();
            TaskDialog.Show("Export", sb.ToString() );
            using (package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("IFC Parameters");
                worksheet.Cells[1, 1].Value = "Element ID";
                worksheet.Cells[1, 2].Value = "Element Name";
                worksheet.Cells[1, 3].Value = "Parameter Name";
                worksheet.Cells[1, 4].Value = "Parameter Value";
                int row = 2;
                foreach (Element element in elements)
                {
                    IList<Parameter> elementParams = element.Parameters.Cast<Parameter>().Where(p => p.Definition.ParameterGroup == BuiltInParameterGroup.PG_IFC).ToList();
                    foreach (Parameter param in elementParams)
                    {
                        worksheet.Cells[row, 1].Value = element.Id;
                        worksheet.Cells[row, 2].Value = element.Name;
                        worksheet.Cells[row, 3].Value = param.Definition.Name;
                        worksheet.Cells[row, 4].Value = param.AsString();
                        row++;
                    }
                }
                // Save the Excel file
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "IFC_Parameters.xlsx");
                File.WriteAllBytes(filePath, package.GetAsByteArray());
                TaskDialog.Show("Export", $"Excel file saved to {filePath}");
            }
        }
    }
}
