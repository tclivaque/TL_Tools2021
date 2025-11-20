using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

[Transaction(TransactionMode.Manual)]
public class ResetOverridesCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View vistaActiva = doc.ActiveView;

        try
        {
            // Obtener todos los elementos visibles en la vista activa
            FilteredElementCollector collector = new FilteredElementCollector(doc, vistaActiva.Id)
                .WhereElementIsNotElementType();

            using (Transaction trans = new Transaction(doc, "Reset Overrides Gráficos"))
            {
                trans.Start();

                // Crear configuración de override vacía (restablece a por defecto)
                OverrideGraphicSettings overridePorDefecto = new OverrideGraphicSettings();

                foreach (Element elem in collector)
                {
                    try
                    {
                        vistaActiva.SetElementOverrides(elem.Id, overridePorDefecto);
                    }
                    catch
                    {
                        // Continuar con el siguiente elemento si hay error
                    }
                }

                trans.Commit();
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Error: {ex.Message}";
            return Result.Failed;
        }
    }
}