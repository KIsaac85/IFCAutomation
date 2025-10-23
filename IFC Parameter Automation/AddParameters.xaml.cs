using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
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

        #endregion
        public string JsonPath { get; set; }
        public AddParameters(UIDocument uidoc)
        {
            InitializeComponent();
            _uidoc = uidoc;
            doc = uidoc.Document;
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
            catch (Exception) { TaskDialog.Show("Error", "error"); }


            if (objectReference != null)
            {
                singleElement = doc.GetElement(objectReference.ElementId);

                JsonPath="";

                if (singleElement.Category.Id.Value == (int)BuiltInCategory.OST_Walls)
                    {

                    JsonPath = @"F:\Access\semester 3\Project\IFC Parameter Automation\Pset_WallCommon.JSON";
                }
                    else if (singleElement.Category.Id.Value == (int)BuiltInCategory.OST_StructuralColumns)
                    {

                    JsonPath = @"F:\Access\semester 3\Project\IFC Parameter Automation\Pset_DoorCommon.JSON";

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
    
    }
}
