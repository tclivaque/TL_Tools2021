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

            string mensaje = "";
            ElementSet elementos = new ElementSet();

            // Crear un wrapper para ExternalCommandData
            // Como no podemos crear ExternalCommandData directamente,
            // accedemos a lo que necesitamos (UIApplication) que ya tenemos

            // Ejecutar la lógica del comando directamente
            Document doc = uidoc.Document;
            View vistaActiva = doc.ActiveView;

            // Llamar al método público Execute del comando
            // Nota: esto requiere que el comando maneje internamente la falta de ExternalCommandData
            // Por ahora, usaremos reflexión o crearemos un método helper

            // Alternativa: crear método estático en EjecutarOverrideCommand
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
