using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelImportExport
{
    public partial class MainWindowViewModel : ModelBase
    {
        private List<Workset> worksetsList;
        public List<Workset> WorksetsList
        {
            get
            {
                return worksetsList;
            }
            set
            {
                worksetsList = value;
                OnPropertyChanged("WorksetsList");
            }
        }

        private Workset selectedWorkset;
        public Workset SelectedWorkset
        {
            get
            {
                return selectedWorkset;
            }
            set
            {
                selectedWorkset = value;
                OnPropertyChanged("SelectedWorkset");
            }
        }

        public ICommand btnApplyWorkset => new RelayCommandWithoutParameter(OnbtnApplyWorkset);
        private void OnbtnApplyWorkset()
        {
            MainWindowModelService.SetWorksets(ElementInformationList, SelectedWorkset);
            dataGrid.Items.Refresh();
        }
    }
}
