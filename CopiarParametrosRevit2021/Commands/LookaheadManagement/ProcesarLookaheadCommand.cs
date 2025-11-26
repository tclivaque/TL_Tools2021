using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;
using CopiarParametrosRevit2021.Helpers;
using TL_Tools2021.Commands.LookaheadManagement.Services;

namespace TL_Tools2021.Commands.LookaheadManagement
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ProcesarLookaheadCommand : IExternalCommand
    {
        // --- CONFIGURACIÓN ---
        private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";
        private const string SCHEDULE_SHEET = "LOOKAHEAD";
        private const string CONFIG_SHEET = "ENTRADAS_SCRIPT_2.0 LOOKAHEAD";

        // Constantes para filtros de muros
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
                // 1. Obtener disciplina del documento
                string docTitle = doc.Title;
                string discipline = GetDisciplineFromTitle(docTitle);

                // 2. Conexión y Lectura
                var sheetsService = new GoogleSheetsService();
                var configReader = new ConfigReader(sheetsService, SPREADSHEET_ID, SCHEDULE_SHEET, CONFIG_SHEET);

                // 3. Leer Reglas
                var rules = configReader.ReadConfigRules();
                if (rules.Count == 0)
                {
                    TaskDialog.Show("Error", "No se leyeron reglas de configuración.");
                    return Result.Failed;
                }

                // 4. Leer Datos del cronograma
                var scheduleData = configReader.ReadScheduleData(rules);

                if (scheduleData.Count == 0)
                {
                    TaskDialog.Show("Aviso", "No se encontraron actividades en el cronograma.");
                    return Result.Succeeded;
                }

                // 5. Preparación - Obtener solo reglas relevantes para esta disciplina
                var relevantRules = rules
                    .Where(r => r.Disciplines.Contains(discipline, StringComparer.OrdinalIgnoreCase))
                    .ToList();

                if (!relevantRules.Any())
                {
                    TaskDialog.Show("Aviso", $"No hay reglas configuradas para la disciplina '{discipline}'.");
                    return Result.Cancelled;
                }

                // 6. Recolectar elementos
                var collector = new ElementCollectorService(doc);
                var elementCache = collector.CollectElements(relevantRules);

                var paramService = new ParameterService(doc);
                var filterService = new FilterService(
                    paramService,
                    ANCHO_14_PIES,
                    ANCHO_19_PIES,
                    TOLERANCIA_PIES);

                var processor = new LookAheadProcessor(paramService, filterService, elementCache);

                // 7. Ejecución
                using (TransactionGroup tg = new TransactionGroup(doc, "Procesar Look Ahead"))
                {
                    tg.Start();

                    using (Transaction trans = new Transaction(doc, "Asignar Semanas"))
                    {
                        trans.Start();
                        processor.AssignWeeks(scheduleData, discipline);
                        processor.ApplyExecuteAndSALogic(relevantRules);
                        trans.Commit();
                    }

                    tg.Assimilate();
                }

                TaskDialog.Show("Éxito", $"Proceso finalizado.\nDisciplina: {discipline}\nActividades procesadas: {scheduleData.Count}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
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