using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class CarsCommand : IExternalCommand
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
                    "Selecciona los elementos para CARS");
            }

            if (referenciasElementos.Count == 0)
            {
                message = "No se seleccionaron elementos.";
                return Result.Cancelled;
            }

            // Calcular la semana de ejecución
            string valorSemanaEjecucion = CalcularSemanaEjecucion();

            // Valores a asignar (iguales que valorización excepto por la semana)
            bool valorEjecutado = true;
            bool valorEnCurso = false;
            bool valorRestriccion = false;

            using (Transaction t = new Transaction(doc, "CARS - Asignación de valores"))
            {
                t.Start();

                foreach (Reference r in referenciasElementos)
                {
                    try
                    {
                        Element elemento = doc.GetElement(r);

                        // Buscar y asignar parámetro EJECUTADO (booleano)
                        var pEjecutado = elemento.LookupParameter("EJECUTADO");
                        if (pEjecutado != null && pEjecutado.StorageType == StorageType.Integer)
                        {
                            pEjecutado.Set(valorEjecutado ? 1 : 0);
                        }

                        // Buscar y asignar parámetro SEMANA DE EJECUCION (texto)
                        var pSemanaEjecucion = elemento.LookupParameter("SEMANA DE EJECUCION");
                        if (pSemanaEjecucion != null && pSemanaEjecucion.StorageType == StorageType.String)
                        {
                            pSemanaEjecucion.Set(valorSemanaEjecucion);
                        }

                        // Buscar y asignar parámetro EN CURSO (booleano)
                        var pEnCurso = elemento.LookupParameter("EN CURSO");
                        if (pEnCurso != null && pEnCurso.StorageType == StorageType.Integer)
                        {
                            pEnCurso.Set(valorEnCurso ? 1 : 0);
                        }

                        // Buscar y asignar parámetro RESTRICCION (booleano)
                        var pRestriccion = elemento.LookupParameter("RESTRICCION");
                        if (pRestriccion != null && pRestriccion.StorageType == StorageType.Integer)
                        {
                            pRestriccion.Set(valorRestriccion ? 1 : 0);
                        }

                        // Buscar y asignar parámetro LOOK AHEAD (texto)
                        var pLookahead = elemento.LookupParameter("LOOK AHEAD");
                        if (pLookahead != null && pLookahead.StorageType == StorageType.String)
                        {
                            pLookahead.Set("EJECUTADO");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Continuar con el siguiente elemento sin mostrar error
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

    private string CalcularSemanaEjecucion()
    {
        // Obtener la fecha actual
        DateTime hoy = DateTime.Today;

        // Calcular cuántos días faltan para el próximo lunes
        int diasParaLunes = (7 - (int)hoy.DayOfWeek + 1) % 7;
        if (diasParaLunes == 0)
        {
            diasParaLunes = 7; // Si hoy es lunes, que el próximo sea el siguiente lunes
        }

        // Calcular la fecha del próximo lunes
        DateTime fechaEntrega = hoy.AddDays(diasParaLunes);

        // Definir la fecha del lunes de la semana 1
        DateTime inicioSemanas = new DateTime(2024, 12, 9);

        // Calcular el número de la semana según tu sistema
        int semanaEntrega = (int)((fechaEntrega - inicioSemanas).TotalDays / 7);

        // Retornar el resultado
        return "SEM " + semanaEntrega.ToString();
    }
}