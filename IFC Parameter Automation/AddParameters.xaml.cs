    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.UI.Selection;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Runtime.Serialization;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;
    using System.Windows.Documents;
    using System.Windows.Input;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using System.Windows.Navigation;
    using System.Windows.Shapes;


    namespace IFC_Parameter_Automation
    {
        /// <summary>
        /// Interaction logic for AddParameters.xaml
        /// </summary>
        public partial class AddParameters : Window
        {

            #region Document
            private UIDocument _uidoc { get; set; }
            private Document doc { get; set; }
            #endregion


            #region Single Element Selection Members
            private Element singleElement { get; set; }
            private Reference objectReference { get; set; }
            private static SelectionFilter SingleSelectionFilter { get; set; }
            public PropertyListViewModel _viewModel { get; set; }
            public singleSelectedElement selectedElement { get; set; }
            public List<PropertyItem> selectedProperties { get; set; } = new List<PropertyItem>();
            public Category cat { get; set; }


             private AddParameterHandler _handler;
             private ExternalEvent _externalEvent;

        #endregion
        public string JsonPath { get; set; }
            public AddParameters(UIDocument uidoc)
            {
                InitializeComponent();
                _uidoc = uidoc;
                doc = uidoc.Document;


            // create ExternalEvent and handler
            _handler = new AddParameterHandler();
            _externalEvent = ExternalEvent.Create(_handler);

            SingleSelectionFilter = new SelectionFilter();
                // Assign the DataContext to your ViewModel
                _viewModel = new PropertyListViewModel(JsonPath);
                DataContext = _viewModel;

            }

            private void Pick_Object(object sender, RoutedEventArgs e)
            {
                Hide();

                try
                {
                    objectReference = _uidoc.Selection.PickObject(ObjectType.Element, SingleSelectionFilter);
                }
                catch (Exception) {  }


                if (objectReference != null)
                {
                    singleElement = doc.GetElement(objectReference.ElementId);

                    JsonPath = "";

                    if (singleElement.Category.Id.Value == (int)BuiltInCategory.OST_Windows)
                    {

                        JsonPath = @"F:\Access\semester 3\Project\IFC Parameter Automation\IFC Parameter Automation\Pset_WindowCommon.JSON";
                        selectedElement = new singleSelectedElement(singleElement);

                    }
                    else if (singleElement.Category.Id.Value == (int)BuiltInCategory.OST_Doors)
                    {

                        JsonPath = @"F:\Access\semester 3\Project\IFC Parameter Automation\IFC Parameter Automation\Pset_DoorCommon.JSON";
                        selectedElement = new singleSelectedElement(singleElement);

                    }
                else if (singleElement.Category.Id.Value == (int)BuiltInCategory.OST_Walls)
                {

                    JsonPath = @"F:\Access\semester 3\Project\IFC Parameter Automation\IFC Parameter Automation\Pset_WallCommon.JSON";
                    selectedElement = new singleSelectedElement(singleElement);

                }
                if (!string.IsNullOrEmpty(JsonPath))
                    {
                        _viewModel.LoadProperties(JsonPath);

                        // Refresh DataContext bindings
                        DataContext = null;
                        DataContext = _viewModel;
                    }
                }

                Show();
            }

            private void Cancel_Click(object sender, RoutedEventArgs e)
            {
                Close();
            }

            private void Confirm_Click(object sender, RoutedEventArgs e)
            {
                if (selectedElement == null)
                {
                    TaskDialog.Show("Error", "No element selected.");
                    return;
                }

                // Get selected properties from the ViewModel
                selectedProperties = _viewModel.Properties.Where(p => p.IsSelected).ToList();

                if (!selectedProperties.Any())
                {
                    TaskDialog.Show("Info", "No parameters selected.");
                    return;
                }
                cat = singleElement.Category;
            // Build ParamList for handler
            _handler.ParamList.Clear();
            foreach (var prop in selectedProperties)
            {
                bool alreadyExists = selectedElement.IFCparameters
                    .Any(p => p.Definition.Name.Equals(prop.Code, StringComparison.OrdinalIgnoreCase));

                if (alreadyExists)
                {
                    // you may collect these to show to user
                    continue;
                }

                _handler.ParamList.Add(new ParamInfo()
                {
                    Code = prop.Code,
                    DataType = prop.DataType,
                    Description = prop.Definition,
                    Category = cat
                });
            }

            if (!_handler.ParamList.Any())
            {
                TaskDialog.Show("Info", "All selected parameters already exist or nothing to add.");
                return;
            }

            // assign the document so handler knows which document to modify
            _handler.Doc = doc;

            // raise the external event; Revit will call handler.Execute(...) on its thread
             _externalEvent.Raise();
            // Note: Raise returns void, but external event will run soon. Do not call Revit API here.



        }
        public void AddParameter(string code, string dataType, string description, Category category)
            {
                // Create a new shared parameter file (temporary)
                string modulePath = doc.PathName;
                string tit = doc.Title;

                string sharedParamPath =
                System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                  "IFC_Parameter_Automation",
                  "SharedParameters.txt");

                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(sharedParamPath));

                if (!File.Exists(sharedParamPath))
                {
                    using (FileStream fs = File.Create(sharedParamPath)) { }
                }

                //  Tell Revit to use this file as the shared parameter file
                _uidoc.Application.Application.SharedParametersFilename = sharedParamPath;

                // Now open it safely
                DefinitionFile defFile = _uidoc.Application.Application.OpenSharedParameterFile();
                if (defFile == null)
                {
                    TaskDialog.Show("Error", "Could not open Shared Parameter file.");
                    return;
                }

                // Create or get the group
                DefinitionGroup group = defFile.Groups.Cast<DefinitionGroup>()
                    .FirstOrDefault(g => g.Name == "Add IFC Params")
                    ?? defFile.Groups.Create("Add IFC Params");

                // Choose the data type
                ForgeTypeId forgeTypeId = SpecTypeId.String.Text;
                if (dataType.Equals("Boolean", StringComparison.OrdinalIgnoreCase))
                    forgeTypeId = SpecTypeId.Boolean.YesNo;
                else if (dataType.Equals("Real", StringComparison.OrdinalIgnoreCase))
                    forgeTypeId = SpecTypeId.Number;

                ExternalDefinitionCreationOptions options = new ExternalDefinitionCreationOptions(code, forgeTypeId)
                {
                    Description = description
                };
           
                Definition newParamDef = group.Definitions.Create(options);

                // Set up category binding
                CategorySet categorySet = _uidoc.Application.Application.Create.NewCategorySet();
                categorySet.Insert(category);

                InstanceBinding binding = _uidoc.Application.Application.Create.NewInstanceBinding(categorySet);
                BindingMap bindingMap = doc.ParameterBindings;

         
                    bool inserted = bindingMap.Insert(newParamDef, binding, GroupTypeId.Ifc);
                
                    if (inserted)
                        TaskDialog.Show("Success", $"Parameter '{code}' added successfully to {category.Name}.");
                    else
                        TaskDialog.Show("Info", $"Parameter '{code}' already exists or could not be added.");
            
            }

        }
    }
