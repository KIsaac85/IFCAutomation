using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;



namespace IFC_Parameter_Automation
{
    public class ExcelExport
    {
        public Document _doc { get; set; }
        private ExcelPackage package { get; set; }
        public IList<Parameter> elementParams { get; set; }

        public void ExportParametersToExcel(Document doc, IList<BuiltInCategory> categories)
        {
            ExcelPackage.License.SetNonCommercialPersonal("abc");

           
            _doc = doc;
            using (package = new ExcelPackage())
            {
                foreach (var bic in categories)
                {
                    ExportCategorySheet(bic);
                }

                SaveExcel(package);
            }
            
        }
        private void ExportCategorySheet(BuiltInCategory bic)
        {
            var elements = new FilteredElementCollector(_doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .ToElements();

            if (!elements.Any())
                return;

            string sheetName = bic.ToString().Replace("OST_", "");
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

            worksheet.Cells[1, 1].Value = "Element ID";
            worksheet.Cells[1, 2].Value = "Element Name";
            worksheet.Cells[1, 3].Value = "Parameter Name";
            worksheet.Cells[1, 4].Value = "Parameter Value";

            int row = 2;
            ElementId previousId = null;
            int mergeStartRow = row;

            foreach (Element element in elements)
            {
                var elementParams = element.Parameters
                    .Cast<Parameter>()
                    .Where(p => p.Definition.GetGroupTypeId() == GroupTypeId.Ifc)
                    .ToList();

                foreach (Parameter param in elementParams)
                {
                    if (!IsParameterEmpty(param))
                        continue;

                    if (previousId != null && previousId != element.Id)
                    {
                        int mergeEnd = row - 1;
                        if (mergeEnd > mergeStartRow)
                        {
                            worksheet.Cells[mergeStartRow, 1, mergeEnd, 1].Merge = true;
                            worksheet.Cells[mergeStartRow, 2, mergeEnd, 2].Merge = true;
                        }
                        mergeStartRow = row;
                    }

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

            worksheet.Cells.AutoFitColumns();
            worksheet.Cells.Style.HorizontalAlignment =
                OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment =
                OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

            worksheet.Cells.Style.WrapText = true;


            var range = worksheet.Cells[
                worksheet.Dimension.Start.Row,
                worksheet.Dimension.Start.Column,
                worksheet.Dimension.End.Row,
                worksheet.Dimension.End.Column];

            range.Style.Border.Top.Style =
            range.Style.Border.Bottom.Style =
            range.Style.Border.Left.Style =
            range.Style.Border.Right.Style =
                OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            
        }


        private bool IsParameterEmpty(Parameter param)
        {
            if (param == null)
                return true;
            if (param.Definition.Name.Equals("Export to IFC", StringComparison.OrdinalIgnoreCase))
                return false;
            // YES / NO parameters
            if (param.Definition.GetDataType() == SpecTypeId.Boolean.YesNo)
            {
                switch (param.AsValueString())
                {
                    case "No":
                        return false;
                    case "Yes": 
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
                    switch (param.AsValueString())
                    {
                        case "By Type":
                            return false;
                        case "0":
                            return false;
                        default:
                            return true;
                    }

                case StorageType.Double:

                    return !param.HasValue;

                case StorageType.ElementId:
                    return param.AsElementId() == ElementId.InvalidElementId;

                default:
                    return true;
            }
        }






        private void SaveExcel(ExcelPackage package)
        {
            SaveFileDialog saveFile = new SaveFileDialog
            {
                FileName = "IFC_Parameters",
                DefaultExt = ".xlsx",
                Filter = "Excel Sheet (*.xlsx)|*.xlsx"
            };

            if (saveFile.ShowDialog() != true)
                return;

            try
            {
                FileInfo file = new FileInfo(saveFile.FileName);

                if (file.Exists)
                    file.Delete(); 

                package.SaveAs(file);
                Process.Start(file.FullName);
            }
            catch (IOException)
            {
                TaskDialog.Show(
                    "Excel Export Error",
                    "The Excel file is currently open.\n\n" +
                    "Please close the file and try again."
                );
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Failed", ex.Message);
            }
        }


    }
}
