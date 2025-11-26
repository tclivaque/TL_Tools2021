using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public class LeyendaEventHandler : IExternalEventHandler
{
    public string ValorSeleccionado { get; set; }
    public bool MostrarTodos { get; set; }
    public bool OcultarElementos { get; set; }  // Nueva propiedad
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
            using (Transaction trans = new Transaction(doc, OcultarElementos ? "Ocultar elementos por valor" : "Aislar elementos por valor"))
            {
                trans.Start();

                if (MostrarTodos)
                {
                    // Mostrar todos los elementos
                    VistaActiva.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                }
                else if (!string.IsNullOrEmpty(ValorSeleccionado))
                {
                    List<ElementId> elementosTarget = new List<ElementId>();

                    if (ValorSeleccionado == "[SIN VALOR]" && ElementosSinValor != null)
                    {
                        elementosTarget.AddRange(ElementosSinValor);
                    }
                    else if (ElementosPorValor.ContainsKey(ValorSeleccionado))
                    {
                        elementosTarget.AddRange(ElementosPorValor[ValorSeleccionado]);
                    }

                    if (elementosTarget.Count > 0)
                    {
                        if (OcultarElementos)
                        {
                            // Ocultar elementos seleccionados
                            VistaActiva.HideElementsTemporary(elementosTarget);
                        }
                        else
                        {
                            // Primero restablecer la vista
                            try
                            {
                                VistaActiva.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                            }
                            catch { }

                            // Aislar elementos seleccionados
                            VistaActiva.IsolateElementsTemporary(elementosTarget);
                        }
                    }
                }

                trans.Commit();
            }
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Error al {(OcultarElementos ? "ocultar" : "aislar")} elementos: {ex.Message}");
        }
    }

    public string GetName()
    {
        return "LeyendaEventHandler";
    }
}