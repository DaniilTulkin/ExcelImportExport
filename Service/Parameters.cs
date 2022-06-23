using Autodesk.Revit.DB;
using System;

namespace ExcelImportExport
{
    public class Parameters
    {
        public static string GetValue(Parameter parameter)
        {
            switch (parameter.StorageType)
            {
                case StorageType.None:
                    return "";
                case StorageType.Integer:
                    return Convert.ToString(parameter.AsInteger());
                case StorageType.Double:
                    return Convert.ToString(UnitUtils.ConvertFromInternalUnits(parameter.AsDouble(), parameter.DisplayUnitType));
                case StorageType.String:
                    return parameter.AsString();
                case StorageType.ElementId:
                    if (parameter.AsElementId() != null)
                    {
                        return parameter.AsValueString();
                    }
                    return "";
                default:
                    return "";
            }
        }

        public static void SetValue(Parameter parameter, string value)
        {
            if (value != "")
            {
                switch (parameter.StorageType)
                {
                    case StorageType.None:
                        break;
                    case StorageType.Integer:
                        parameter.Set(Convert.ToInt32(value));
                        break;
                    case StorageType.Double:
                        parameter.Set(UnitUtils.ConvertToInternalUnits(Convert.ToDouble(value), parameter.DisplayUnitType));
                        break;
                    case StorageType.String:
                        parameter.Set(value);
                        break;
                    case StorageType.ElementId:
                        break;
                    default:
                        break;
                }
            }
        }

        public static bool IsSchedulable(Parameter parameter)
        {
            if (parameter.IsShared) return true;
            else
            {
                if (parameter.Definition is InternalDefinition internaldefinition)
                {
                    if (!(internaldefinition.BuiltInParameter == BuiltInParameter.INVALID)) return true;
                    return false;
                }
                return false;
            }
        }
    }
}
