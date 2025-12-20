using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
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
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class UserControl1 : Window
    {
        #region Document Members
        private UIDocument _uidoc { get; set; }
        private Document doc { get; set; }
        #endregion



        public UserControl1(UIDocument uidoc)
        {
            InitializeComponent();
            _uidoc = uidoc;
            doc = uidoc.Document;
           
        }

        private void Add_IFC_Parameters(object sender, RoutedEventArgs e)
        {
            Close();
            AddParameters addParametersWindow = new AddParameters(_uidoc);
            addParametersWindow.ShowDialog();
        }

        private void Check_IFC_Parameters(object sender, RoutedEventArgs e)
        {
            ExcelExport excelExport = new ExcelExport();
            excelExport.ExportParametersToExcel(doc);
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

     
    }
}
