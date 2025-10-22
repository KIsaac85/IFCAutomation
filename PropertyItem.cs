using System;
using System.Collections.ObjectModel;
using System.IO;
using Newtonsoft.Json;

public ObservableCollection<PropertyItem> Properties { get; set; }



public class PropertyItem
{
    public int Id { get; set; }
    public string Code { get; set; }
    public string DictionaryURI { get; set; }
    public string DataType { get; set; }
    public string Definition { get; set; }

    // For checkbox binding:
    public bool IsSelected { get; set; }
}

public void LoadProperties()
{
    string jsonPath = @"C:\path\to\Pset_DoorCommon.json";
    string jsonText = File.ReadAllText(jsonPath);
    Properties = JsonConvert.DeserializeObject<ObservableCollection<PropertyItem>>(jsonText);
}
