using System.Collections.ObjectModel;
using System.IO;
using Newtonsoft.Json;


namespace IFC_Parameter_Automation
{
    public class PropertyListViewModel
    {
        
        public ObservableCollection<PropertyItem> Properties { get; set; }

        public PropertyListViewModel(string jsonPath)
        {
          
            
        }

        public void LoadProperties(string jsonPath)
        {
            
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