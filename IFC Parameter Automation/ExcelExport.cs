using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;


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
            

            foreach (Element element in elements) {
                IList<Parameter> elementParams = element.Parameters.Cast<Parameter>().Where(p => p.Definition.ParameterGroup == BuiltInParameterGroup.PG_IFC).ToList();
                foreach (Parameter param in elementParams)
                {
                    IFCParameters.Add(param);
                    
                }
            }

            //IList<Parameter> IFCParameters = ele.Parameters.Cast<Parameter>().Where(p => p.Definition.ParameterGroup == BuiltInParameterGroup.PG_IFC).ToList();
            
            using (package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(elements.FirstOrDefault().Category.Name);
                worksheet.Cells[1, 1].Value = "Element ID";
                worksheet.Cells[1, 2].Value = "Element Name";
                worksheet.Cells[1, 3].Value = "Parameter Name";
                worksheet.Cells[1, 4].Value = "Parameter Value";
                int row = 2;
                ElementId previousId = null;
                int mergeStartRow = row;

                foreach (Element element in elements)
                {
                    if (previousId != null && previousId != element.Id)
                    {
                        int mergeEndRow = row - 1;

                        if (mergeEndRow > mergeStartRow)
                        {
                            worksheet.Cells[mergeStartRow, 1, mergeEndRow, 1].Merge = true;
                            worksheet.Cells[mergeStartRow, 2, mergeEndRow, 2].Merge = true;
                        }

                        mergeStartRow = row;
                    }
                    IList<Parameter> elementParams = element.Parameters.Cast<Parameter>().Where(p => p.Definition.ParameterGroup == BuiltInParameterGroup.PG_IFC).ToList();
                    foreach (Parameter param in elementParams)
                    {
                        if (IsParameterEmpty(param))
                        {
                            worksheet.Cells[row, 1].Value = element.Id;

                            worksheet.Cells[row, 2].Value = element.Name;
                            worksheet.Cells[row, 3].Value = param.Definition.Name;
                            worksheet.Cells[row, 4].Value = param.AsString();
                            previousId = element.Id;
                            row++;
                        }

                    }
                    int lastEndRow = row - 1;
                    if (lastEndRow > mergeStartRow)
                    {
                        worksheet.Cells[mergeStartRow, 1, lastEndRow, 1].Merge = true;
                        worksheet.Cells[mergeStartRow, 2, lastEndRow, 2].Merge = true;
                    }

                    

                }
                worksheet.Cells.AutoFitColumns();
                worksheet.Cells.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                var range = worksheet.Cells[
                worksheet.Dimension.Start.Row,
                worksheet.Dimension.Start.Column,
                worksheet.Dimension.End.Row,
                worksheet.Dimension.End.Column
                ];
                range.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                SaveFileDialog saveFile = new SaveFileDialog
                {
                    FileName = "NewSheet", // Default file name
                    DefaultExt = ".xlsx", // Default file extension
                    Filter = "Excel Sheet (.xlsx)|*" // Filter files by extension
                };
                bool? result = saveFile.ShowDialog();
                string errormessage = null;
                do
                {
                    try
                    {
                        saveFile.OverwritePrompt = true;
                        savedialogue(package, saveFile);
                    }

                    catch (Exception e)
                    {
                        if (result != false)
                        {
                            errormessage = e.Message;
                            TaskDialog.Show(e.Message, "Error saving the file!");
                            break;
                        }

                    }
                } while (result != false && errormessage != null);

            }
            
        }
        private bool IsParameterEmpty(Parameter param)
        {
            if (param == null) return true;

            // Detect Yes/No (Boolean)
            if (param.Definition.GetDataType() == SpecTypeId.Boolean.YesNo)
            {
                switch (param.AsValueString())
                {
                    case "Yes": 
                        return false;
                    case "No": 
                        return false;
                    case null: 
                        return true;
                    default:
                        return true;
                }
                
                
            }

            switch (param.StorageType)
            {
                case StorageType.String:
                    return string.IsNullOrWhiteSpace(param.AsString());

                case StorageType.Integer:
                    return param.AsInteger() == 0;

                case StorageType.Double:
                    return param.AsDouble() == 0;

                case StorageType.ElementId:
                    return param.AsElementId() == ElementId.InvalidElementId;

                default:
                    return true;
            }

        }

        private void savedialogue(ExcelPackage package, SaveFileDialog saveFile)
        {
            FileInfo filename = new FileInfo(saveFile.FileName);
            package.SaveAs(filename);
            Process.Start(Path.Combine(filename.Directory.ToString(), filename.ToString()));
        }


    }
}
