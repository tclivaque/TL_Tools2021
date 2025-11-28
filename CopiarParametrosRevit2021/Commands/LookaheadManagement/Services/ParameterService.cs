using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement.Services
{
    public class ParameterService
    {
        private readonly Document _doc;

        public ParameterService(Document doc)
        {
            _doc = doc;
        }

        public string GetParameterValue(Element element, string paramName)
        {
            if (element == null) return "";

            Parameter param = element.LookupParameter(paramName);

            if (param == null)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    param = _doc.GetElement(typeId)?.LookupParameter(paramName);
                }
            }

            if (param == null || !param.HasValue)
                return "";

            string value = "";
            switch (param.StorageType)
            {
                case StorageType.String:
                    value = param.AsString();
                    break;
                case StorageType.Integer:
                    value = param.AsInteger().ToString();
                    break;
                case StorageType.Double:
                    value = param.AsDouble().ToString();
                    break;
                default:
                    value = param.AsValueString();
                    break;
            }

            return (value ?? "").Trim();
        }

        public string GetParameterValue(Element elem, BuiltInParameter bip)
        {
            if (elem == null) return "";

            // Intenta obtener el parámetro por su ID interno (más seguro que por nombre)
            Parameter param = elem.get_Parameter(bip);

            if (param == null) return "";

            // Devuelve el valor como texto (o el valor interno si es numérico)
            return param.AsString() ?? param.AsValueString() ?? "";
        }

        public bool SetParameterValue(Element element, string paramName, object value)
        {
            Parameter param = element?.LookupParameter(paramName);

            if (param != null && !param.IsReadOnly)
            {
                try
                {
                    if (value is string s)
                        param.Set(s);
                    else if (value is int i)
                        param.Set(i);
                    else
                        param.Set(value.ToString());

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }

        public XYZ GetElementCenter(Element element)
        {
            try
            {
                var bbox = element.get_BoundingBox(null);
                if (bbox != null)
                    return bbox.Min + (bbox.Max - bbox.Min) * 0.5;
            }
            catch { }

            return XYZ.Zero;
        }

        public double? GetElementWidth(Element element)
        {
            try
            {
                Element typeElement = _doc.GetElement(element.GetTypeId());

                if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Walls)
                {
                    return ((WallType)typeElement).Width;
                }

                if (element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    string widthStr = GetParameterValue(element, "ANCHO");
                    if (double.TryParse(widthStr, out double width))
                        return width;
                }
            }
            catch { }

            return null;
        }

        public List<string> GetFloorMaterialNames(Element floor)
        {
            var materialNames = new List<string>();

            if (floor is Floor f)
            {
                try
                {
                    if (_doc.GetElement(f.GetTypeId()) is FloorType floorType)
                    {
                        var compoundStructure = floorType.GetCompoundStructure();
                        if (compoundStructure != null)
                        {
                            foreach (var layer in compoundStructure.GetLayers())
                            {
                                var material = _doc.GetElement(layer.MaterialId) as Material;
                                if (material != null)
                                    materialNames.Add(material.Name.ToLower());
                            }
                        }
                    }
                }
                catch { }
            }

            return materialNames;
        }
    }
}
