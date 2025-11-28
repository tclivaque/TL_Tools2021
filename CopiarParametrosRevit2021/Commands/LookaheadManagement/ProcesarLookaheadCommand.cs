using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CopiarParametrosRevit2021.Commands.LookaheadManagement.Services;
using System;
using System.Linq;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ProcesarLookaheadCommand : IExternalCommand
    {
        // === CONFIGURACIÓN ===
        // Se mantiene la ID original exacta
        private const string SPREADSHEET_ID = "1DPSRZDrqZkCxaHQrIIaz5NSf5m3tLJcggvAx9k8x9SA";
        private const string SCHEDULE_SHEET_NAME = "LOOKAHEAD";
        private const string CONFIG_SHEET_NAME = "CONFIG_ACTIVIDADES";

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
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // ============ PASO 1: ASIGNAR LOOK AHEAD ============
                Result asignacionResult = EjecutarAsignacion(doc, out string asignacionMsg);

                if (asignacionResult == Result.Failed)
                {
                    message = asignacionMsg;
                    return Result.Failed;
                }

                // ============ PASO 2: ACTUALIZAR MEMBRETE ============
                Result membreteResult = EjecutarMembrete(doc, out string membreteMsg);

                // Continuar aunque falle el membrete (la asignación ya fue exitosa)
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private Result EjecutarAsignacion(Document doc, out string resultMessage)
        {
            try
            {
                // 1. Obtener Identidad del Modelo (Activo ID)
                string docTitle = doc.Title;
                string activeId = "";

                if (docTitle.Length >= 23)
                {
                    activeId = docTitle.Substring(20, 3);
                }

                // 2. Conexión y Lectura
                var sheetsService = new GoogleSheetsService();
                var configReader = new ConfigReader(
                    sheetsService,
                    SPREADSHEET_ID,
                    SCHEDULE_SHEET_NAME,
                    CONFIG_SHEET_NAME);

                var configRules = configReader.ReadConfigRules();

                if (configRules == null || !configRules.Any())
                {
                    resultMessage = "No se encontraron reglas en CONFIG_ACTIVIDADES.";
                    return Result.Failed;
                }

                var scheduleData = configReader.ReadScheduleData(configRules, activeId);

                if (scheduleData == null || !scheduleData.Any())
                {
                    resultMessage = $"No se encontraron datos para el activo '{activeId}' en la hoja LOOKAHEAD.";
                    return Result.Failed;
                }

                // 3. Preparación
                string discipline = GetDisciplineFromTitle(docTitle);
                var relevantRules = configRules
                    .Where(r => r.Disciplines.Contains(discipline, StringComparer.OrdinalIgnoreCase))
                    .ToList();

                if (!relevantRules.Any())
                {
                    resultMessage = $"No hay reglas configuradas para la disciplina '{discipline}'.";
                    return Result.Failed;
                }

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

                resultMessage = "Look Ahead asignado correctamente.";
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                resultMessage = ex.Message;
                return Result.Failed;
            }
        }

        private Result EjecutarMembrete(Document doc, out string resultMessage)
        {
            try
            {
                var membreteCommand = new MembreteLookaheadCommand();
                string msg = "";
                ElementSet elems = new ElementSet();

                // Crear un ExternalCommandData simulado
                // Como no podemos crear uno real fácilmente, ejecutaremos la lógica directamente
                Result result = ActualizarMembreteDirecto(doc, out msg);

                resultMessage = msg;
                return result;
            }
            catch (Exception ex)
            {
                resultMessage = ex.Message;
                return Result.Failed;
            }
        }

        private Result ActualizarMembreteDirecto(Document doc, out string message)
        {
            try
            {
                // Obtener información del archivo
                string fileName = doc.Title;
                if (System.IO.Path.HasExtension(fileName))
                    fileName = System.IO.Path.GetFileNameWithoutExtension(fileName);

                string[] parts = fileName.Split('-');

                if (parts.Length < 5)
                {
                    message = "El nombre del archivo no tiene el formato esperado.";
                    return Result.Failed;
                }

                string especialidad = parts[3];
                string rawSector = parts[4];
                string codigoActivo = "";

                if (rawSector.StartsWith("000") && rawSector.Length >= 6)
                    codigoActivo = rawSector.Substring(3, 3);
                else if (rawSector.StartsWith("0") && rawSector.Length >= 4)
                    codigoActivo = rawSector.Substring(1, 3);
                else
                    codigoActivo = rawSector.Length >= 3 ? rawSector.Substring(rawSector.Length - 3) : rawSector;

                // Calcular fechas
                DateTime hoy = DateTime.Today;
                int diasParaLunes = ((int)DayOfWeek.Monday - (int)hoy.DayOfWeek + 7) % 7;
                DateTime fechaEntrega = hoy.AddDays(diasParaLunes);
                DateTime inicioSemanas = new DateTime(2024, 12, 9);
                int semanaEntrega = (int)((fechaEntrega - inicioSemanas).TotalDays / 7) + 1;

                // Buscar plano
                ViewSheet targetSheet = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .FirstOrDefault(s => s.SheetNumber.Contains("-LPS-S"));

                if (targetSheet == null)
                {
                    message = "No se encontró ningún plano que contenga '-LPS-S' en su número.";
                    return Result.Failed;
                }

                FamilyInstance titleblock = new FilteredElementCollector(doc, targetSheet.Id)
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .FirstOrDefault();

                if (titleblock == null)
                {
                    message = $"El plano {targetSheet.SheetNumber} no tiene un membrete colocado.";
                    return Result.Failed;
                }

                // Actualizar
                using (Transaction t = new Transaction(doc, "Actualizar Membrete LookAhead"))
                {
                    t.Start();

                    targetSheet.Name = $"PLANO DE LOOK AHEAD PLANNING - SEMANA {semanaEntrega}";
                    string newSheetNumber = $"200114-CCC02-PL-{especialidad}-{codigoActivo}-LPS-S{semanaEntrega}";

                    if (!targetSheet.SheetNumber.Equals(newSheetNumber))
                    {
                        try { targetSheet.SheetNumber = newSheetNumber; } catch { }
                    }

                    Parameter dateParam = targetSheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE);
                    if (dateParam != null)
                        dateParam.Set(fechaEntrega.ToString("dd/MM/yyyy"));

                    SetParameter(titleblock, "NOMBRE DEL PROYECTO", "I.E. FRANCISCO BOLOGNESI CERVANTES");
                    SetParameter(titleblock, "DISCIPLINA", GetDisciplinaName(especialidad));
                    SetParameter(titleblock, "ESCALA", "INDICADA");
                    SetParameter(titleblock, "RV", "C0");
                    SetParameter(titleblock, "ESCALA GRAFICA", 0);
                    SetParameter(titleblock, "NORTE REFERENTE", 0);
                    SetParameter(titleblock, "EMITIDO PARA CONSTRUCCIÓN", 0);

                    foreach (Parameter param in titleblock.Parameters)
                    {
                        if (param.Definition.Name.StartsWith("ACTIVO "))
                        {
                            param.Set(param.Definition.Name.Contains(codigoActivo) ? 1 : 0);
                        }
                    }

                    t.Commit();
                }

                message = $"Plano: {targetSheet.SheetNumber}\nSemana: {semanaEntrega}";
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void SetParameter(Element element, string paramName, object value)
        {
            Parameter param = element.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly)
            {
                if (value is string s)
                    param.Set(s);
                else if (value is int i)
                    param.Set(i);
                else if (value is double d)
                    param.Set(d);
            }
        }

        private string GetDisciplinaName(string sigla)
        {
            switch (sigla.ToUpper())
            {
                case "AR": return "ARQUITECTURA";
                case "EE": return "INSTALACIONES ELÉCTRICAS";
                case "ES": return "ESTRUCTURAS";
                case "SA": return "INSTALACIONES SANITARIAS";
                case "ME": return "INSTALACIONES MECÁNICAS";
                default: return "ARQUITECTURA";
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