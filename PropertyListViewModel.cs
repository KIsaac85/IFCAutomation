using System.Collections.ObjectModel;
using System.IO;
using Newtonsoft.Json;


namespace IFC_Parameter_Automation
{
    public class PropertyListViewModel
    {
        public ObservableCollection<PropertyItem> Properties { get; set; }

        public PropertyListViewModel()
        {
            LoadProperties();
        }

        private void LoadProperties()
        {
            string jsonPath = @"F:\Access\semester 3\Project\IFC Parameter Automation\Pset_DoorCommon.JSON";
            if (File.Exists(jsonPath))
            {
                string jsonText = File.ReadAllText(jsonPath);
                Properties = JsonConvert.DeserializeObject<ObservableCollection<PropertyItem>>(jsonText);
            }
            else
            {
                Properties = new ObservableCollection<PropertyItem>();
            }
        }
    }
}