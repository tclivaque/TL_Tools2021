using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class TarrajeoCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            IList<Reference> referenciasElementos;

            // Verificar si hay elementos pre-seleccionados
            ICollection<ElementId> elementosSeleccionados = uidoc.Selection.GetElementIds();

            if (elementosSeleccionados.Count > 0)
            {
                // Usar elementos pre-seleccionados
                referenciasElementos = new List<Reference>();
                foreach (ElementId id in elementosSeleccionados)
                {
                    Element elem = doc.GetElement(id);
                    if (elem != null)
                    {
                        referenciasElementos.Add(new Reference(elem));
                    }
                }
            }
            else
            {
                // Pedir al usuario que seleccione elementos
                referenciasElementos = uidoc.Selection.PickObjects(ObjectType.Element,
                    "Selecciona los elementos para tarrajeo");
            }

            if (referenciasElementos.Count == 0)
            {
                message = "No se seleccionaron elementos.";
                return Result.Cancelled;
            }

            using (Transaction t = new Transaction(doc, "Tarrajeo - Asignación MET_SUP"))
            {
                t.Start();

                foreach (Reference r in referenciasElementos)
                {
                    try
                    {
                        Element elemento = doc.GetElement(r);

                        // Buscar y asignar parámetro MET_SUP
                        var pMetSup = elemento.LookupParameter("MET_SUP");
                        if (pMetSup != null && pMetSup.StorageType == StorageType.String)
                        {
                            pMetSup.Set("TR_EJECUTADO");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Continuar con el siguiente elemento
                        System.Diagnostics.Debug.WriteLine($"Error procesando elemento: {ex.Message}");
                    }
                }

                t.Commit();
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            // Manejar cualquier excepción sin mostrar mensaje al usuario
            if (ex.Message.Contains("cancelled") || ex.Message.Contains("canceled"))
            {
                return Result.Cancelled;
            }

            return Result.Failed;
        }
    }
}