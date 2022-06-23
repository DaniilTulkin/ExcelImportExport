using System.Collections.Generic;
using System.Windows.Input;

namespace ExcelImportExport
{
    public partial class MainWindowViewModel : ModelBase
    {
        public ICommand btnExport => new RelayCommandWithoutParameter(OnbtnExport);
        private void OnbtnExport()
        {
            Excel.WriteExcel(dataGrid);
        }

        public ICommand btnImport => new RelayCommandWithoutParameter(OnbtnImport);
        private void OnbtnImport()
        {
            List<DynamicDictionaryWrapper> ElementInformationListFromExcel = Excel.ReadExcel();
            if (ElementInformationListFromExcel != null)
                MainWindowModelService.WriteParametersToElementInformationList(ElementInformationList, ElementInformationListFromExcel);
        }
    }
}
