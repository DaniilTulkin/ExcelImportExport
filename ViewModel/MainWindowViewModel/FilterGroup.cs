using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace ExcelImportExport
{
    public partial class MainWindowViewModel : ModelBase
    {

        private List<string> propertiesList;
        public List<string> PropertiesList
        {
            get
            {
                return propertiesList;
            }
            set
            {
                propertiesList = value;
                OnPropertyChanged("PropertiesList");
            }
        }

        private string selectedPropertyForFilter;
        public string SelectedPropertyForFilter
        {
            get
            {
                return selectedPropertyForFilter;
            }
            set
            {
                selectedPropertyForFilter = value;
                OnPropertyChanged("SelectedPropertyForFilter");

                if (!string.IsNullOrEmpty(value)) FilterIsEnabled = true;
                else FilterIsEnabled = false;
            }
        }

        private bool filterIsEnabled;
        public bool FilterIsEnabled
        {
            get
            {
                return filterIsEnabled;
            }
            set
            {
                filterIsEnabled = value;
                OnPropertyChanged("FilterIsEnabled");
            }
        }

        private List<string> operatorsList = new List<string>
        {
            "равно",
            "не равно",
            "содержит",
            "не содержит"
        };
        public List<string> OperatorsList
        {
            get
            {
                return operatorsList;
            }
            set
            {
                operatorsList = value;
                OnPropertyChanged("OperatorsList");
            }
        }

        private string selectedOperatorForFilter = "равно";
        public string SelectedOperatorForFilter
        {
            get
            {
                return selectedOperatorForFilter;
            }
            set
            {
                selectedOperatorForFilter = value;
                OnPropertyChanged("SelectedOperatorForFilter");
            }
        }

        private List<string> valuesForFilterList;
        public List<string> ValuesForFilterList
        {
            get
            {
                return valuesForFilterList;
            }
            set
            {
                valuesForFilterList = value;
                OnPropertyChanged("ValuesForFilterList");
            }
        }

        private string selectedValueForFilter = "";
        public string SelectedValueForFilter
        {
            get
            {
                return selectedValueForFilter;
            }
            set
            {
                selectedValueForFilter = value;
                OnPropertyChanged("SelectedValueForFilter");
            }
        }

        public ICommand cmbPropertyForFilterChanged => new RelayCommandWithoutParameter(OncmbPropertyForFilterChanged);
        private void OncmbPropertyForFilterChanged()
        {
            ValuesForFilterList = MainWindowModelService.PopulatePropertiesForFilter(ElementInformationList, SelectedPropertyForFilter);            
        }

        public ICommand btnFilter => new RelayCommandWithoutParameter(OnbtnFilter);
        private void OnbtnFilter()
        {
            MainWindowModelService.Filter(ElementInformationList, 
                                          SelectedPropertyForFilter, 
                                          SelectedOperatorForFilter, 
                                          SelectedValueForFilter);    
            dataGrid.Items.Refresh();
        }

        public ICommand btnDiscard => new RelayCommandWithoutParameter(OnbtnDiscard);
        private void OnbtnDiscard()
        {
            OncmbCategoryChanged();
        }
    }
}
