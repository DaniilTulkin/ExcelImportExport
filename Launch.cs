using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace ExcelImportExport.Model
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExcelImportExport : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            try
            {
                MainWindow mainWindow = new MainWindow(app);
                mainWindow?.ShowDialog();
            }
            catch (Exception ex)
            {
            }
        }

        //public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        //{
        //    try
        //    {
        //        MainWindow mainWindow = new MainWindow(commandData.Application);
        //        mainWindow?.ShowDialog();
        //        return Result.Succeeded;
        //    }
        //    catch (Exception ex)
        //    {
        //        return Result.Failed;
        //    }
        //}
        public string GetName() => nameof(ExcelImportExport);
    }
}
