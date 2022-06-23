using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelImportExport
{
    public class MainWindowModel : INotifyPropertyChanged
    {
        private Element element;
        public Element Element
        {
            get
            {
                return element;
            }
            set
            {
                element = value;
                OnPropertyChange("Element");
            }
        }

        private Dictionary<string, dynamic> properties;
        public Dictionary<string, dynamic> Properties
        {
            get
            {
                return properties;
            }
            set
            {
                properties = value;
                OnPropertyChange("Properties");
            }
        }

        public MainWindowModel(Element element)
        {
            Element = element;

            Dictionary<string, dynamic> dict = new Dictionary<string, dynamic>();
            dict["ElementId"] = element.Id.IntegerValue.ToString();
            foreach (Parameter parameter in element.Parameters)
            {
                if (Parameters.IsSchedulable(parameter))
                {
                    string parName = parameter.Definition.Name;
                    string parValue = Parameters.GetValue(parameter);
                    dict[parName] = parValue;
                }
            }
            Properties = dict;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChange([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
            }
        }
    }
}
