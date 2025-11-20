using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class EnCursoCommand : IExternalCommand
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
                    "Selecciona los elementos para marcar EN CURSO");
            }

            if (referenciasElementos.Count == 0)
            {
                message = "No se seleccionaron elementos.";
                return Result.Cancelled;
            }

            using (Transaction t = new Transaction(doc, "EN CURSO - Asignación de valores"))
            {
                t.Start();

                foreach (Reference r in referenciasElementos)
                {
                    try
                    {
                        Element elemento = doc.GetElement(r);

                        // EN CURSO = True
                        var pEnCurso = elemento.LookupParameter("EN CURSO");
                        if (pEnCurso != null && pEnCurso.StorageType == StorageType.Integer)
                        {
                            pEnCurso.Set(1);
                        }

                        // EJECUTADO = False
                        var pEjecutado = elemento.LookupParameter("EJECUTADO");
                        if (pEjecutado != null && pEjecutado.StorageType == StorageType.Integer)
                        {
                            pEjecutado.Set(0);
                        }

                        // SEMANA DE EJECUCION = vacío
                        var pSemanaEjecucion = elemento.LookupParameter("SEMANA DE EJECUCION");
                        if (pSemanaEjecucion != null && pSemanaEjecucion.StorageType == StorageType.String)
                        {
                            pSemanaEjecucion.Set("");
                        }

                        // LOOK AHEAD - vaciar si tiene valor "EJECUTADO"
                        var pLookahead = elemento.LookupParameter("LOOK AHEAD");
                        if (pLookahead != null && pLookahead.StorageType == StorageType.String)
                        {
                            string valorLookahead = pLookahead.AsString();
                            if (!string.IsNullOrEmpty(valorLookahead) && valorLookahead == "EJECUTADO")
                            {
                                pLookahead.Set("");
                            }
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