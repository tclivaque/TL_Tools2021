using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public class LeyendaEventHandler : IExternalEventHandler
{
    public string ValorSeleccionado { get; set; }
    public bool MostrarTodos { get; set; }
    public View VistaActiva { get; set; }
    public Dictionary<string, List<ElementId>> ElementosPorValor { get; set; }
    public List<ElementId> ElementosSinValor { get; set; }

    public void Execute(UIApplication app)
    {
        UIDocument uidoc = app.ActiveUIDocument;
        Document doc = uidoc.Document;

        if (VistaActiva == null || ElementosPorValor == null)
            return;

        try
        {
            using (Transaction trans = new Transaction(doc, "Aislar elementos por valor"))
            {
                trans.Start();

                if (MostrarTodos)
                {
                    // Mostrar todos los elementos
                    VistaActiva.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                }
                else if (!string.IsNullOrEmpty(ValorSeleccionado))
                {
                    // Primero restablecer la vista
                    try
                    {
                        VistaActiva.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                    }
                    catch { }

                    // Luego aislar elementos con el valor seleccionado
                    List<ElementId> elementosAMostrar = new List<ElementId>();

                    if (ValorSeleccionado == "[SIN VALOR]" && ElementosSinValor != null)
                    {
                        elementosAMostrar.AddRange(ElementosSinValor);
                    }
                    else if (ElementosPorValor.ContainsKey(ValorSeleccionado))
                    {
                        elementosAMostrar.AddRange(ElementosPorValor[ValorSeleccionado]);
                    }

                    if (elementosAMostrar.Count > 0)
                    {
                        VistaActiva.IsolateElementsTemporary(elementosAMostrar);
                    }
                }

                trans.Commit();
            }
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Error al aislar elementos: {ex.Message}");
        }
    }

    public string GetName()
    {
        return "LeyendaEventHandler";
    }
}