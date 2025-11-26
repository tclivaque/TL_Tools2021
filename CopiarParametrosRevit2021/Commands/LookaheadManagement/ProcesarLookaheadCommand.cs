using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using CopiarParametrosRevit2021.Helpers; // Servicio unificado
using TL_Tools2021.Commands.LookaheadManagement.Services;

namespace TL_Tools2021.Commands.LookaheadManagement
{
    [Transaction(TransactionMode.Manual)]
    public class ProcesarLookaheadCommand : IExternalCommand
    {
        // --- CONFIGURACIÓN ---
        private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";
        private const string SCHEDULE_SHEET = "LOOKAHEAD";
        private const string CONFIG_SHEET = "ENTRADAS_SCRIPT_2.0 LOOKAHEAD";

        // EL ID ÚNICO PARA DEBUG
        private const string ACTIVE_ID = "282354";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Instanciar servicios
            var sheetsService = new GoogleSheetsService();
            var configReader = new ConfigReader(sheetsService, SPREADSHEET_ID, SCHEDULE_SHEET, CONFIG_SHEET);

            try
            {
                // INICIAR LOG DE DEBUG
                configReader.InitializeDebug(ACTIVE_ID);
                configReader.Log("Iniciando comando desde Revit...");

                // 1. Leer Reglas
                var rules = configReader.ReadConfigRules();
                if (rules.Count == 0)
                {
                    TaskDialog.Show("Error", "No se leyeron reglas. Revisa el Log en el Escritorio.");
                    return Result.Failed;
                }

                // 2. Leer Datos filtrando por ACTIVE_ID
                var scheduleData = configReader.ReadScheduleData(rules, ACTIVE_ID);

                if (scheduleData.Count == 0)
                {
                    configReader.Log("RESULTADO: 0 elementos encontrados.");
                    TaskDialog.Show("Aviso", $"No se encontraron actividades para el ID {ACTIVE_ID}.\nRevisa el archivo 'Lookahead_Debug_{ACTIVE_ID}.txt' en tu escritorio para ver el detalle.");
                    return Result.Succeeded;
                }

                // 3. (Aquí iría el procesamiento real en Revit con LookAheadProcessor)
                // var processor = new LookAheadProcessor(commandData.Application.ActiveUIDocument.Document);
                // processor.Process(scheduleData);

                configReader.Log($"Proceso completado. {scheduleData.Count} actividades listas para procesar.");

                TaskDialog.Show("Éxito", $"Proceso finalizado.\nActividades encontradas: {scheduleData.Count}\n\nLog generado en Escritorio: Lookahead_Debug_{ACTIVE_ID}.txt");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                configReader.Log($"\n[EXCEPCIÓN]: {ex.Message}");
                configReader.Log(ex.StackTrace);
                message = $"Error crítico: {ex.Message}";
                return Result.Failed;
            }
            finally
            {
                // SIEMPRE GUARDAR EL LOG
                configReader.SaveDebugLog();
            }
        }
    }
}