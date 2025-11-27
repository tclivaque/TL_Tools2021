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
        private const string ACTIVE_ID = "282354";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var sheetsService = new GoogleSheetsService();
            var configReader = new ConfigReader(sheetsService, SPREADSHEET_ID, SCHEDULE_SHEET, CONFIG_SHEET);

            try
            {
                // 1. Leer Reglas
                var rules = configReader.ReadConfigRules();
                if (rules.Count == 0)
                {
                    TaskDialog.Show("Error", "No se encontraron reglas de configuración.");
                    return Result.Failed;
                }

                // 2. Leer Datos filtrando por ACTIVE_ID
                var scheduleData = configReader.ReadScheduleData(rules, ACTIVE_ID);

                if (scheduleData.Count == 0)
                {
                    TaskDialog.Show("Aviso", $"No se encontraron actividades para el ID {ACTIVE_ID}.");
                    return Result.Succeeded;
                }

                // 3. (Aquí iría el procesamiento real en Revit con LookAheadProcessor)
                // var processor = new LookAheadProcessor(commandData.Application.ActiveUIDocument.Document);
                // processor.Process(scheduleData);

                TaskDialog.Show("Éxito", $"Proceso finalizado.\nActividades encontradas: {scheduleData.Count}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}