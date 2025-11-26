using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
// CORRECCIÓN: Importar el namespace donde vive EjecutarOverrideCommand
using CopiarParametrosRevit2021.Commands.ParameterReview;

public class OverrideCommandEventHandler : IExternalEventHandler
{
    // Bandera para saber qué botón se presionó (Color o Reset)
    public bool EsReset { get; set; } = false;

    public void Execute(UIApplication app)
    {
        try
        {
            if (EsReset)
            {
                // Ejecutar lógica de Reset (Requiere actualización en ResetOverridesCommand)
                ResetOverridesCommand.ExecuteLogic(app);
            }
            else
            {
                // Ejecutar lógica de Colorear (Con Debug)
                EjecutarOverrideCommand.ExecuteFromEvent(app);
            }
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Error en el evento: {ex.Message}");
        }
    }

    public string GetName()
    {
        return "OverrideCommandEventHandler";
    }
}