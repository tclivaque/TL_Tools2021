using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;
using CopiarParametrosRevit2021.Commands.LookaheadManagement.Services;
using CopiarParametrosRevit2021.Commands.LookaheadManagement.Models;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AsignarLookaheadCommand : IExternalCommand
    {
        // === CONFIGURACIÓN ===
        private const string SPREADSHEET_ID = "1DPSRZDrqZkCxaHQrIIaz5NSf5m3tLJcggvAx9k8x9SA";
        private const string SCHEDULE_SHEET_NAME = "LOOKAHEAD";

        // === CONSTANTES ===
        private const double ANCHO_14_METROS = 0.14;
        private const double ANCHO_19_METROS = 0.19;
        private const double TOLERANCIA_METROS = 0.01;

        private static readonly double ANCHO_14_PIES = UnitUtils.Convert(
            ANCHO_14_METROS, UnitTypeId.Meters, UnitTypeId.Feet);
        private static readonly double ANCHO_19_PIES = UnitUtils.Convert(
            ANCHO_19_METROS, UnitTypeId.Meters, UnitTypeId.Feet);
        private static readonly double TOLERANCIA_PIES = UnitUtils.Convert(
            TOLERANCIA_METROS, UnitTypeId.Meters, UnitTypeId.Feet);

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            try
            {
                // 1. Obtener Identidad del Modelo (Activo ID)
                string docTitle = doc.Title;
                string activeId = "";

                if (docTitle.Length >= 23)
                {
                    activeId = docTitle.Substring(20, 3); // Toma 3 caracteres (posiciones 20-22)
                }

                // 2. Conexión y Lectura
                var sheetsService = new GoogleSheetsService();
                var configReader = new ConfigReader(
                    sheetsService,
                    SPREADSHEET_ID,
                    SCHEDULE_SHEET_NAME);

                var configRules = configReader.ReadConfigRules();

                if (configRules == null || !configRules.Any())
                {
                    TaskDialog.Show("Error", "No se encontraron reglas en CONFIG_ACTIVIDADES.");
                    return Result.Failed;
                }

                var scheduleData = configReader.ReadScheduleData(configRules, activeId);

                if (scheduleData == null || !scheduleData.Any())
                {
                    TaskDialog.Show("Info",
                        $"No se encontraron datos para el activo '{activeId}' en la hoja LOOKAHEAD.");
                    return Result.Cancelled;
                }

                // 3. Preparación
                string discipline = GetDisciplineFromTitle(docTitle);
                var relevantRules = configRules
                    .Where(r => r.Disciplines.Contains(discipline, StringComparer.OrdinalIgnoreCase))
                    .ToList();

                if (!relevantRules.Any())
                    return Result.Cancelled;

                var collector = new ElementCollectorService(doc);
                var elementCache = collector.CollectElements(relevantRules);

                var paramService = new ParameterService(doc);
                var filterService = new FilterService(
                    doc,
                    paramService,
                    ANCHO_14_PIES,
                    ANCHO_19_PIES,
                    TOLERANCIA_PIES);

                var processor = new LookAheadProcessor(paramService, filterService, elementCache);

                // 4. Ejecución
                using (TransactionGroup tg = new TransactionGroup(doc, "Asignar Look Ahead"))
                {
                    tg.Start();

                    using (Transaction trans = new Transaction(doc, "Modificar Parámetros"))
                    {
                        trans.Start();
                        processor.AssignWeeks(scheduleData, discipline);
                        processor.ApplyExecuteAndSALogic(relevantRules);
                        trans.Commit();
                    }

                    tg.Assimilate();
                }

                TaskDialog.Show("Éxito", "Look Ahead asignado correctamente.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error Crítico", $"Ocurrió un error:\n{ex.Message}");
                return Result.Failed;
            }
        }

        private string GetDisciplineFromTitle(string title)
        {
            if (string.IsNullOrEmpty(title) || title.Length < 18)
                return "AR";

            try
            {
                string discipline = title.Substring(16, 2).ToUpper();
                return (discipline == "ES" || discipline == "SA" ||
                        discipline == "DT" || discipline == "EE") ? discipline : "AR";
            }
            catch
            {
                return "AR";
            }
        }
    }
}
