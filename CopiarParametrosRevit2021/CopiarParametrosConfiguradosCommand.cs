using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class CopiarParametrosConfiguradosCommand : IExternalCommand
{
    private static string ConfigFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "TL_Tools2021",
        "parametros_copiar_config.txt");

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            // Leer parámetros configurados
            string parametrosConfig = LeerParametrosGuardados();

            if (string.IsNullOrWhiteSpace(parametrosConfig))
            {
                TaskDialog.Show("Error", "No hay parámetros configurados.\n\nUse primero 'Configurar Parámetros a Copiar'.");
                return Result.Failed;
            }

            // Separar parámetros por coma
            string[] nombresParametros = parametrosConfig.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                         .Select(p => p.Trim())
                                                         .Where(p => !string.IsNullOrEmpty(p))
                                                         .ToArray();

            if (nombresParametros.Length == 0)
            {
                TaskDialog.Show("Error", "No se encontraron parámetros válidos en la configuración.");
                return Result.Failed;
            }

            // Seleccionar elemento fuente
            Reference referenciaFuente = uidoc.Selection.PickObject(ObjectType.Element, "Selecciona el elemento fuente");
            Element elementoFuente = doc.GetElement(referenciaFuente);

            // Obtener valores de los parámetros del elemento fuente
            Dictionary<string, ParametroInfo> valoresParametros = new Dictionary<string, ParametroInfo>();

            foreach (string nombreParam in nombresParametros)
            {
                ParametroInfo info = ObtenerParametroInfo(elementoFuente, nombreParam, doc);
                if (info != null)
                {
                    valoresParametros[nombreParam] = info;
                }
            }

            if (valoresParametros.Count == 0)
            {
                message = "El elemento fuente no tiene ninguno de los parámetros configurados.";
                return Result.Failed;
            }

            // Seleccionar elementos destino
            IList<Reference> referenciasDestino = uidoc.Selection.PickObjects(ObjectType.Element, "Selecciona los elementos destino");

            using (Transaction t = new Transaction(doc, "Copiar parámetros configurados"))
            {
                t.Start();

                foreach (Reference r in referenciasDestino)
                {
                    Element elementoDestino = doc.GetElement(r);

                    foreach (var kvp in valoresParametros)
                    {
                        string nombreParam = kvp.Key;
                        ParametroInfo infoFuente = kvp.Value;

                        // Intentar encontrar y establecer el parámetro en el destino
                        EstablecerParametro(elementoDestino, nombreParam, infoFuente, doc);
                    }
                }

                t.Commit();
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }

    private string LeerParametrosGuardados()
    {
        try
        {
            if (File.Exists(ConfigFilePath))
            {
                return File.ReadAllText(ConfigFilePath).Trim();
            }
        }
        catch { }
        return string.Empty;
    }

    private ParametroInfo ObtenerParametroInfo(Element elemento, string nombreParametro, Document doc)
    {
        // Intentar obtener de ejemplar
        Parameter param = elemento.LookupParameter(nombreParametro);

        if (param != null && param.HasValue)
        {
            return new ParametroInfo
            {
                NombreParametro = nombreParametro,
                EsDeTipo = false,
                StorageType = param.StorageType,
                ValorString = param.StorageType == StorageType.String ? param.AsString() : null,
                ValorInteger = param.StorageType == StorageType.Integer ? (int?)param.AsInteger() : null,
                ValorDouble = param.StorageType == StorageType.Double ? (double?)param.AsDouble() : null,
                ValorElementId = param.StorageType == StorageType.ElementId ? param.AsElementId() : null
            };
        }

        // Intentar obtener de tipo
        ElementId typeId = elemento.GetTypeId();
        if (typeId != ElementId.InvalidElementId)
        {
            ElementType elemType = doc.GetElement(typeId) as ElementType;
            if (elemType != null)
            {
                param = elemType.LookupParameter(nombreParametro);
                if (param != null && param.HasValue)
                {
                    return new ParametroInfo
                    {
                        NombreParametro = nombreParametro,
                        EsDeTipo = true,
                        StorageType = param.StorageType,
                        ValorString = param.StorageType == StorageType.String ? param.AsString() : null,
                        ValorInteger = param.StorageType == StorageType.Integer ? (int?)param.AsInteger() : null,
                        ValorDouble = param.StorageType == StorageType.Double ? (double?)param.AsDouble() : null,
                        ValorElementId = param.StorageType == StorageType.ElementId ? param.AsElementId() : null
                    };
                }
            }
        }

        return null;
    }

    private void EstablecerParametro(Element elemento, string nombreParametro, ParametroInfo infoFuente, Document doc)
    {
        Parameter param = null;

        if (infoFuente.EsDeTipo)
        {
            // Buscar en tipo
            ElementId typeId = elemento.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                ElementType elemType = doc.GetElement(typeId) as ElementType;
                if (elemType != null)
                {
                    param = elemType.LookupParameter(nombreParametro);
                }
            }
        }
        else
        {
            // Buscar en ejemplar
            param = elemento.LookupParameter(nombreParametro);
        }

        if (param != null && !param.IsReadOnly && param.StorageType == infoFuente.StorageType)
        {
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        if (infoFuente.ValorString != null)
                            param.Set(infoFuente.ValorString);
                        break;
                    case StorageType.Integer:
                        if (infoFuente.ValorInteger.HasValue)
                            param.Set(infoFuente.ValorInteger.Value);
                        break;
                    case StorageType.Double:
                        if (infoFuente.ValorDouble.HasValue)
                            param.Set(infoFuente.ValorDouble.Value);
                        break;
                    case StorageType.ElementId:
                        if (infoFuente.ValorElementId != null)
                            param.Set(infoFuente.ValorElementId);
                        break;
                }
            }
            catch
            {
                // Continuar con el siguiente si hay error
            }
        }
    }
}

// Clase auxiliar para almacenar información del parámetro
public class ParametroInfo
{
    public string NombreParametro { get; set; }
    public bool EsDeTipo { get; set; }
    public StorageType StorageType { get; set; }
    public string ValorString { get; set; }
    public int? ValorInteger { get; set; }
    public double? ValorDouble { get; set; }
    public ElementId ValorElementId { get; set; }
}