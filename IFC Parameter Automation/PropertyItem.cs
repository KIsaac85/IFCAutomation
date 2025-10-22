using System;
using System.Collections.ObjectModel;
using System.IO;



namespace IFC_Parameter_Automation
{
   

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
}




