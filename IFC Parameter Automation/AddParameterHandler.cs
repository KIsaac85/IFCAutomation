// AddParameterHandler.cs
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IFC_Parameter_Automation
{
    public class AddParameterHandler : IExternalEventHandler
    {
        // The list of parameters to add (set by the WPF window before Raise)
        public List<ParamInfo> ParamList { get; set; } = new List<ParamInfo>();

        // Keep reference to the document we want to modify (set by UI)
        public Document Doc { get; set; }

        // Path to a persistent shared parameters file - keep it stable so GUIDs remain
        private string GetSharedParamPath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "IFC_Parameter_Automation");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "SharedParameters.txt");
        }

        public void Execute(UIApplication uiApp)
        {
            if (Doc == null)
            {
                TaskDialog.Show("AddParameterHandler", "Document is null. Cannot add parameters.");
                return;
            }

            try
            {
                string sharedParamPath = GetSharedParamPath();

                if (!File.Exists(sharedParamPath))
                {
                    using (FileStream fs = File.Create(sharedParamPath)) { }
                }

                // Tell Revit which shared parameter file to use
                uiApp.Application.SharedParametersFilename = sharedParamPath;

                DefinitionFile defFile = uiApp.Application.OpenSharedParameterFile();
                if (defFile == null)
                {
                    TaskDialog.Show("Error", "Could not open Shared Parameter file.");
                    return;
                }

                // Ensure group exists
                DefinitionGroup group = defFile.Groups.Cast<DefinitionGroup>()
                    .FirstOrDefault(g => g.Name == "Add IFC Params")
                    ?? defFile.Groups.Create("Add IFC Params");

                BindingMap bindingMap = Doc.ParameterBindings;

                using (Transaction t = new Transaction(Doc, "Add IFC Parameters (batch)"))
                {
                    t.Start();

                    foreach (var p in ParamList)
                    {
                        try
                        {
                            // Check category allows bound parameters
                            if (!p.Category.AllowsBoundParameters)
                            {
                                TaskDialog.Show("Info", $"Category {p.Category.Name} does not allow bound parameters.");
                                continue;
                            }

                            // Determine ForgeTypeId
                            ForgeTypeId forgeTypeId = SpecTypeId.String.Text;
                            if (p.DataType.Equals("Boolean", StringComparison.OrdinalIgnoreCase))
                                forgeTypeId = SpecTypeId.Boolean.YesNo;
                            else if (p.DataType.Equals("Real", StringComparison.OrdinalIgnoreCase))
                                forgeTypeId = SpecTypeId.Number;

                            // If the definition already exists in the group (same name), reuse it:
                            Definition def = group.Definitions.Cast<Definition>()
                                .FirstOrDefault(d => d.Name.Equals(p.Code, StringComparison.OrdinalIgnoreCase));

                            if (def == null)
                            {
                                ExternalDefinitionCreationOptions opts =
                                    new ExternalDefinitionCreationOptions(p.Code, forgeTypeId)
                                    {
                                        Description = p.Description
                                    };
                                def = group.Definitions.Create(opts);
                            }

                            // Build category set & binding
                            CategorySet cs = uiApp.Application.Create.NewCategorySet();
                            cs.Insert(p.Category);
                            InstanceBinding binding = uiApp.Application.Create.NewInstanceBinding(cs);

                            bool existsInMap = false;
                            

                            var it = bindingMap.ForwardIterator();
                            it.Reset();

                            while (it.MoveNext())
                            {
                                Definition defExisting = it.Key as Definition;
                                Binding bindingExisting = it.Current as Binding;

                                if (defExisting != null &&
                                    defExisting.Name.Equals(p.Code, StringComparison.OrdinalIgnoreCase) &&
                                    bindingExisting is InstanceBinding)
                                {
                                    existsInMap = true;
                                    break;
                                }
                            }




                            if (existsInMap)
                            {
                                // Parameter binding already present - skip or optionally ReInsert carefully
                                continue;
                            }

                            bool inserted = bindingMap.Insert(def, binding, GroupTypeId.Ifc);

                            // Optionally, if you want to set a default value on the current element(s), you can do that here.
                            if (!inserted)
                            {
                                // If Insert failed, notify (but don't ReInsert blindly)
                                // You may want to log more details in real code
                                //TaskDialog.Show("Info", $"Parameter '{p.Code}' could not be inserted.");
                            }
                        }
                        catch (Exception exInner)
                        {
                            TaskDialog.Show("Error (param loop)", exInner.ToString());
                        }
                    }

                    t.Commit();
                }

                // Clear list after operation
                ParamList.Clear();

                TaskDialog.Show("Success", "Parameters processed. Check Categories / Project Parameters.");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("AddParameterHandler Error", ex.ToString());
            }
        }

        public string GetName()
        {
            return "Add IFC Parameters Handler";
        }
    }
}
        