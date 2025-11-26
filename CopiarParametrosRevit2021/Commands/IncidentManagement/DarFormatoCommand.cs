using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

[Transaction(TransactionMode.Manual)]
public class DarFormatoCommand : IExternalCommand
{
    private const string GoogleSheetURL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQA78GFh2VfPCixaNhGZ8lroXn8bJf10b7TWTHGULtsYV2D1KWsyQNMcQpzxd0SPoL_Rrm5weLi8dBi/pub?gid=187112304&single=true&output=csv";

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        try
        {
            // 1. IMPORTAR DATOS DE GOOGLE SHEETS
            List<string> listaSdiIds;
            List<string> listaSdiTipos;

            try
            {
                var (columna1, columna34) = GoogleSheetsHelper.GetColumnsFromSheet(GoogleSheetURL, 1, 34);
                listaSdiIds = columna1;
                listaSdiTipos = columna34;

                if (listaSdiIds.Count == 0 || listaSdiTipos.Count == 0)
                {
                    TaskDialog.Show("Error", "No se pudieron obtener datos del Google Sheet.\n\nVerifique:\n- Conexión a internet\n- URL del documento\n- Permisos de acceso");
                    return Result.Failed;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error de Conexión",
                    $"No se pudo descargar la información del Google Sheet.\n\n" +
                    $"Error: {ex.Message}\n\n" +
                    $"Verifique su conexión a internet.");
                return Result.Failed;
            }

            // 2. PROCESAR ELEMENTOS
            SDIProcessor.ResultadosProcesamiento resultados;

            using (Transaction trans = new Transaction(doc, "Dar Formato SDI/Incidencias"))
            {
                trans.Start();

                try
                {
                    resultados = SDIProcessor.ProcesarElementos(doc, listaSdiIds, listaSdiTipos);
                    trans.Commit();
                }
                catch (Exception ex)
                {
                    trans.RollBack();
                    TaskDialog.Show("Error de Procesamiento",
                        $"Ocurrió un error al procesar los elementos.\n\n" +
                        $"Error: {ex.Message}");
                    return Result.Failed;
                }
            }

            // 3. MOSTRAR RESULTADOS
            try
            {
                VentanaResultadosSDI ventana = new VentanaResultadosSDI(resultados);
                ventana.ShowDialog();
            }
            catch (Exception) // <-- CORREGIDO: Se quitó la variable 'ex' que causaba la advertencia
            {
                // Si falla la ventana, al menos mostrar el resumen en TaskDialog
                TaskDialog.Show("Resultados del Procesamiento", resultados.GenerarResumen());
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Error general: {ex.Message}";
            TaskDialog.Show("Error", message);
            return Result.Failed;
        }
    }
}