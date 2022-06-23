using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Windows.Controls;

namespace ExcelImportExport
{
    public class MainWindowModelService
    {
        private UIApplication app;
        private UIDocument uidoc;
        private Document doc;
        private RevitEvent revitEvent;

        public MainWindowModelService(UIApplication app)
        {
            this.app = app;
            uidoc = app.ActiveUIDocument;
            doc = uidoc.Document;
            revitEvent = new RevitEvent();
        }

        internal List<Category> PopulateCategories()
        {
            List<Category> categoryList = new List<Category>();
            Categories categories = doc.Settings.Categories;
            foreach (Category category in categories)
            {
                if (category.CategoryType == CategoryType.Model
                    && !category.Name.ToLower().Contains(".dwg"))
                {
                    categoryList.Add(category);
                }
            }
            categoryList.Sort((x, y) => string.Compare(x.Name, y.Name));

            return categoryList;
        }

        internal List<DynamicDictionaryWrapper> PopulateElementInformation(Category selectedCategory, bool isType)
        {
            if (selectedCategory != null)
            {
                BuiltInCategory builtInCategory = (BuiltInCategory)selectedCategory.Id.IntegerValue;
                List<DynamicDictionaryWrapper> list = new List<DynamicDictionaryWrapper>();
                if (isType)
                {
                    var elementsType = new FilteredElementCollector(doc)
                                           .OfCategory(builtInCategory)
                                           .WhereElementIsElementType();
                    foreach (ElementType elementType in elementsType)
                    {
                        list.Add(new DynamicDictionaryWrapper(new MainWindowModel(elementType).Properties));
                    }
                }
                else
                {
                    var elementsInstance = new FilteredElementCollector(doc)
                                           .OfCategory(builtInCategory)
                                           .WhereElementIsNotElementType()
                                           .ToElements();
                    foreach (Element elementInstance in elementsInstance)
                    {
                        list.Add(new DynamicDictionaryWrapper(new MainWindowModel(elementInstance).Properties));
                    }
                }

                if (list.Any()) return list;
                return null;
            }
            return null;
        }

        internal void WriteParametersToElementInformationList(List<DynamicDictionaryWrapper> ElementInformationList, 
                                                                     List<DynamicDictionaryWrapper> ElementInformationListFromExcel)
        {
            foreach (DynamicDictionaryWrapper dictOld in ElementInformationList)
            {
                foreach (DynamicDictionaryWrapper dictNew in ElementInformationListFromExcel)
                {
                    if (dictOld["ElementId"].ToString() == dictNew["ElementId"].ToString())
                    {
                        foreach (string keyOld in dictOld.GetDynamicMemberNames().ToList())
                        {
                            if (dictNew[keyOld]?.ToString() != dictOld[keyOld]?.ToString())
                            {
                                dictOld[keyOld] = dictNew[keyOld];
                            }
                        }
                    }
                }
            }
        }

        internal void Filter( List<DynamicDictionaryWrapper> elementInformationList, string selectedPropertyForFilter, string selectedOperatorForFilter, string selectedValueForFilter)
        {
            var list = elementInformationList.ToList();
            foreach (DynamicDictionaryWrapper dict in list)
            {
                string value = Convert.ToString(dict[selectedPropertyForFilter]);
                switch (selectedOperatorForFilter)
                {
                    case "равно":
                        if (!(value == selectedValueForFilter)) elementInformationList.Remove(dict);
                        break;
                    case "не равно":
                        if (value == selectedValueForFilter) elementInformationList.Remove(dict);
                        break;
                    case "содержит":
                        if (!value.Contains(selectedValueForFilter)) elementInformationList.Remove(dict);
                        break;
                    case "не содержит":
                        if (value.Contains(selectedValueForFilter)) elementInformationList.Remove(dict);
                        break;
                    default:
                        break;
                }
            }
        }

        internal List<string> PopulatePropertiesForFilter(List<DynamicDictionaryWrapper> elementInformationList, string selectedPropertyForFilter)
        {
            List<string> list = new List<string>();
            foreach (DynamicDictionaryWrapper dictionary in elementInformationList)
            {
                string value = Convert.ToString(dictionary[selectedPropertyForFilter]);
                if (!list.Contains(value))
                {
                    list.Add(value);
                }
            }

            return list.OrderBy(s => s).ToList();
        }

        internal void WriteParametersToModel(List<DynamicDictionaryWrapper> familyInstanceInformationList)
        {
            using (Transaction t = new Transaction(doc, "Запись параметров"))
            {
                t.Start();

                foreach (DynamicDictionaryWrapper dictionary in familyInstanceInformationList)
                {
                    ElementId elementId = new ElementId(Convert.ToInt32(dictionary["ElementId"]));
                    Element element = doc.GetElement(elementId);

                    foreach (string key in dictionary.GetDynamicMemberNames())
                    {
                        if (key == "ElementId") continue; 
                        foreach (Parameter parameter in element.ParametersMap)
                        {
                            if (key == parameter.Definition.Name)
                            {
                                string value = dictionary[key]?.ToString();
                                if (value != null && value != Parameters.GetValue(parameter))
                                {
                                    Parameters.SetValue(parameter, value);
                                }
                            }
                        }
                    }
                }

                t.Commit();
            }
        }

        internal void AutoGeneratedColumns(DataGrid dataGrid)
        {
            dataGrid.Columns.Clear();
            var first = dataGrid.ItemsSource.Cast<object>().FirstOrDefault() as DynamicObject;
            if (first == null) return;

            var names = first.GetDynamicMemberNames().OrderBy(s => s);
            foreach (var name in names)
            {
                var column = new DataGridTextColumn
                {
                    Header = name,
                    Binding = new System.Windows.Data.Binding(name)
                };

                dataGrid.Columns.Add(column);
            }

            foreach (var item in dataGrid.Columns)
            {
                if (item.Header.ToString() == "ElementId")
                {
                    item.DisplayIndex = 0;
                    break;
                }
            }
        }

        internal List<Workset> GetWorksets()
        {
            return new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .OrderBy(w => w.Name)
                .ToList();
        }

        internal void SetWorksets(List<DynamicDictionaryWrapper> elementInformationList, Workset selectedWorkset)
        {
            if (selectedWorkset == null) return;
            try
            {
                using (Transaction t = new Transaction(doc, "Перенос в рабочий набор"))
                {
                    t.Start();

                    foreach (DynamicDictionaryWrapper dict in elementInformationList)
                    {
                        ElementId elementId = new ElementId(Convert.ToInt32(dict["ElementId"]));
                        Element element = doc.GetElement(elementId);
                        int worksetId = selectedWorkset.Id.IntegerValue;
                        element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(worksetId);
                        dict["Рабочий набор"] = Convert.ToString(worksetId);
                    }

                    t.Commit();
                }
            }
            catch (Exception)
            {
                return;
            }
        }
    }
}
