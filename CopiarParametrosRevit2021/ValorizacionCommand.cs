using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class ValorizacionCommand : IExternalCommand
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
                    "Selecciona los elementos para valorización");
            }

            if (referenciasElementos.Count == 0)
            {
                message = "No se seleccionaron elementos.";
                return Result.Cancelled;
            }

            // Valores a asignar
            bool valorEjecutado = true;
            int valorSemanaValorizada = CalcularNumeroSemana();
            bool valorEnCurso = false;
            bool valorRestriccion = false;

            using (Transaction t = new Transaction(doc, "Valorización de elementos"))
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

                        // Buscar y asignar parámetro SEMANA VALORIZADA (entero)
                        var pSemanaValorizada = elemento.LookupParameter("SEMANA VALORIZADA");
                        if (pSemanaValorizada != null && pSemanaValorizada.StorageType == StorageType.Integer)
                        {
                            pSemanaValorizada.Set(valorSemanaValorizada);
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
                    }
                    catch (Exception ex)
                    {
                        // Continuar con el siguiente elemento
                        System.Diagnostics.Debug.WriteLine($"Error procesando elemento: {ex.Message}");
                    }
                }

                t.Commit();
            }

            // Sin mensaje de confirmación para mayor velocidad
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

    private int CalcularNumeroSemana()
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

        // Retornar solo el número
        return semanaEntrega;
    }
}