using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

[Transaction(TransactionMode.Manual)]
public class ResetOverridesCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // El comando normal redirige a la lógica compartida
        ExecuteLogic(commandData.Application);
        return Result.Succeeded;
    }

    // Método estático para ser usado por la Ventana (EventHandler)
    public static void ExecuteLogic(UIApplication uiApp)
    {
        UIDocument uidoc = uiApp.ActiveUIDocument;
        if (uidoc == null) return;

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
                        // Verificar si el elemento tiene geometría/categoría válida antes de limpiar
                        if (elem.Category != null)
                        {
                            vistaActiva.SetElementOverrides(elem.Id, overridePorDefecto);
                        }
                    }
                    catch
                    {
                        // Ignorar elementos que no soportan overrides
                    }
                }

                trans.Commit();
            }
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error Reset", $"Error al restablecer colores: {ex.Message}");
        }
    }
}