using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

public class LeyendaEventHandler : IExternalEventHandler
{
    public string ValorSeleccionado { get; set; }
    public bool MostrarTodos { get; set; }
    public bool OcultarElementos { get; set; }
    public bool SeleccionarElementos { get; set; }
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
            // Caso especial: Seleccionar elementos
            if (SeleccionarElementos && !string.IsNullOrEmpty(ValorSeleccionado))
            {
                List<ElementId> elementosTarget = ObtenerElementosTarget();
                if (elementosTarget.Count > 0)
                {
                    uidoc.Selection.SetElementIds(elementosTarget);
                }
                else
                {
                    uidoc.Selection.SetElementIds(new List<ElementId>());
                }
                return;
            }

            // Casos: Aislar / Ocultar / Mostrar Todos
            using (Transaction trans = new Transaction(doc, OcultarElementos ? "Ocultar elementos" : "Aislar elementos"))
            {
                trans.Start();

                if (MostrarTodos)
                {
                    VistaActiva.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                }
                else if (!string.IsNullOrEmpty(ValorSeleccionado))
                {
                    List<ElementId> elementosTarget = ObtenerElementosTarget();

                    if (elementosTarget.Count > 0)
                    {
                        if (OcultarElementos)
                        {
                            VistaActiva.HideElementsTemporary(elementosTarget);
                        }
                        else
                        {
                            try
                            {
                                // Reseteamos antes de aislar para limpiar cualquier estado previo (como los rojos visibles)
                                VistaActiva.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
                            }
                            catch { }

                            VistaActiva.IsolateElementsTemporary(elementosTarget);
                        }
                    }
                }

                trans.Commit();
            }

            // Forzar actualización de vista para asegurar que desaparezcan fantasmas visuales
            uidoc.RefreshActiveView();
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Error al procesar la vista: {ex.Message}");
        }
    }

    private List<ElementId> ObtenerElementosTarget()
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

        return elementosTarget;
    }

    public string GetName()
    {
        return "LeyendaEventHandler";
    }
}