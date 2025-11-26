using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

public class OverrideCommandEventHandler : IExternalEventHandler
{
    public void Execute(UIApplication app)
    {
        try
        {
            // Ejecutar el comando de override
            var comando = new EjecutarOverrideCommand();

            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null)
                return;

            // Ejecutar la lógica estática del comando
            EjecutarOverrideCommand.ExecuteFromEvent(app);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Error al ejecutar override: {ex.Message}");
        }
    }

    public string GetName()
    {
        return "OverrideCommandEventHandler";
    }
}