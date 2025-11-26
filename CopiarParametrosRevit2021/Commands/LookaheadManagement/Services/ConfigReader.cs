using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using TL_Tools2021.Commands.LookaheadManagement.Models;
using CopiarParametrosRevit2021.Helpers; // Asegúrate de que este namespace coincida con tu GoogleSheetsService unificado

namespace TL_Tools2021.Commands.LookaheadManagement.Services
{
    public class ConfigReader
    {
        private readonly GoogleSheetsService _sheetsService;
        private readonly string _spreadsheetId;
        private readonly string _scheduleSheetName;
        private readonly string _configSheetName;

        // --- SISTEMA DE DEBUG ---
        private readonly StringBuilder _debugLog = new StringBuilder();
        private string _debugFilePath;

        public ConfigReader(GoogleSheetsService sheetsService, string spreadsheetId,
            string scheduleSheetName, string configSheetName)
        {
            _sheetsService = sheetsService;
            _spreadsheetId = spreadsheetId;
            _scheduleSheetName = scheduleSheetName;
            _configSheetName = configSheetName;
        }

        // Configura el nombre del archivo de log con el ID
        public void InitializeDebug(string activeId)
        {
            _debugFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                $"Lookahead_Debug_{activeId}.txt");

            _debugLog.Clear();
            Log("==================================================");
            Log($" INICIO DEBUG LOOKAHEAD - {DateTime.Now}");
            Log($" ID BUSCADO: {activeId}");
            Log($" Spreadsheet: {_spreadsheetId}");
            Log("==================================================\n");
        }

        public void Log(string message)
        {
            _debugLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {message}");
        }

        public void SaveDebugLog()
        {
            try
            {
                if (!string.IsNullOrEmpty(_debugFilePath))
                    File.WriteAllText(_debugFilePath, _debugLog.ToString());
            }
            catch { /* Ignorar errores de escritura de log */ }
        }

        private string RemoveAccents(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var normalized = text.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();
            foreach (char c in normalized)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }
            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        public List<ActivityRule> ReadConfigRules()
        {
            Log("[PASO 1] LEYENDO REGLAS DE CONFIGURACIÓN...");
            var rules = new List<ActivityRule>();

            try
            {
                var values = _sheetsService.ReadData(_spreadsheetId, $"'{_configSheetName}'!A2:G");
                Log($"-> Filas encontradas en config: {values.Count}");

                int rowIndex = 2;
                foreach (var row in values)
                {
                    string keyword = GetCell(row, 0);

                    if (string.IsNullOrWhiteSpace(keyword))
                    {
                        rowIndex++;
                        continue;
                    }

                    var rule = new ActivityRule
                    {
                        ActivityKeywords = keyword.ToLower(),
                        Disciplines = GetCell(row, 1).Split(',').Select(d => d.Trim().ToUpper()).ToList(),
                        RawCategories = GetCell(row, 2),
                        FuncionFiltroEspecial = GetCell(row, 3),
                        ParametroFiltro = GetCell(row, 4),
                        KeywordsFiltro = GetCell(row, 5).Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(k => k.Trim().ToLower()).ToList(),
                        FuncionOrden = GetCell(row, 6)
                    };

                    // Parsear Categorías
                    rule.Categories = new List<BuiltInCategory>();
                    var rawCats = rule.RawCategories.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (var c in rawCats)
                    {
                        try
                        {
                            var bic = (BuiltInCategory)Enum.Parse(typeof(BuiltInCategory), c.Trim(), true);
                            rule.Categories.Add(bic);
                        }
                        catch { Log($"   [WARN] Categoría inválida '{c}' en regla '{keyword}'"); }
                    }

                    if (rule.Categories.Any())
                    {
                        rules.Add(rule);
                        Log($"   [OK] Regla cargada: '{keyword}' ({rule.Categories.Count} cats)");
                    }

                    rowIndex++;
                }
            }
            catch (Exception ex)
            {
                Log($"[ERROR] Fallo leyendo reglas: {ex.Message}");
                throw;
            }

            Log($"-> Total reglas válidas: {rules.Count}\n");
            return rules;
        }

        public List<ScheduleData> ReadScheduleData(List<ActivityRule> rules, string activeId)
        {
            Log("[PASO 2] LEYENDO CRONOGRAMA Y BUSCANDO MATCH...");
            var data = new List<ScheduleData>();

            // 1. Leer Encabezados
            var headers = _sheetsService.ReadData(_spreadsheetId, $"'{_scheduleSheetName}'!1:1");
            if (headers == null || headers.Count == 0)
            {
                Log("[ERROR FATAL] No se leyeron encabezados.");
                return data;
            }

            var headerRow = headers[0];
            int activityColumn = -1, startColumn = -1;

            // Buscar columna clave
            for (int i = 0; i < headerRow.Count; i++)
            {
                string h = headerRow[i]?.ToString()?.Trim() ?? "";
                if (h.Equals("DESCRIPCIÓN DE LA ACTIVIDAD", StringComparison.OrdinalIgnoreCase))
                {
                    activityColumn = i;
                    startColumn = i + 1;
                    Log($"-> Columna Actividad: índice {i}");
                    Log($"-> Inicio Semanas: índice {startColumn}");
                    break;
                }
            }

            if (activityColumn == -1)
            {
                Log("[ERROR FATAL] Columna 'DESCRIPCIÓN DE LA ACTIVIDAD' no encontrada.");
                throw new Exception("Columna Actividad no encontrada");
            }

            // 2. Leer Datos (Hasta columna ZZ para asegurar rango)
            var rows = _sheetsService.ReadData(_spreadsheetId, $"'{_scheduleSheetName}'!A2:ZZ");
            Log($"-> Filas de datos leídas: {rows?.Count ?? 0}");

            int rowIndex = 2;
            int matchesFound = 0;

            foreach (var row in rows)
            {
                if (row.Count <= activityColumn) { rowIndex++; continue; }

                string activityName = GetCell(row, activityColumn).ToLower();
                if (string.IsNullOrWhiteSpace(activityName)) { rowIndex++; continue; }

                // --- LOGICA DE MATCHING DE REGLA ---
                string activityNameNormalized = RemoveAccents(activityName);
                ActivityRule matchingRule = null;

                foreach (var r in rules)
                {
                    string ruleKw = RemoveAccents(r.ActivityKeywords);
                    if (activityNameNormalized.Contains(ruleKw))
                    {
                        // Priorizar la regla más específica (keyword más largo)
                        if (matchingRule == null || r.ActivityKeywords.Length > matchingRule.ActivityKeywords.Length)
                        {
                            matchingRule = r;
                        }
                    }
                }

                if (matchingRule == null)
                {
                    // Descomentar para ver todo (puede generar mucho texto)
                    // Log($"Fila {rowIndex}: '{activityName}' -> SIN REGLA"); 
                    rowIndex++;
                    continue;
                }

                // --- REGLA ENCONTRADA ---
                bool isSitio = IsSitio(matchingRule.FuncionFiltroEspecial) ||
                               activityNameNormalized.Contains("exterior") ||
                               activityNameNormalized.Contains("podotactil");

                var groups = new Dictionary<string, ScheduleGroup>();
                bool rowHasActiveId = false;

                // --- BUSCAR EL ID EN LAS SEMANAS ---
                for (int i = startColumn; i < row.Count && i < (startColumn + 30); i++)
                {
                    string cellValue = GetCell(row, i);
                    if (string.IsNullOrWhiteSpace(cellValue) || cellValue.Equals("null", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string week = GetWeek(i - startColumn);
                    if (week == null) continue;

                    string key = "";
                    bool valid = false;

                    if (isSitio)
                    {
                        // Lógica SITIO (Busca PLAT)
                        if (cellValue.ToUpper().StartsWith("PLAT"))
                        {
                            key = cellValue.Replace(" ", "").ToUpper();
                            valid = true;
                        }
                    }
                    else
                    {
                        // Lógica ACTIVO (Busca el ID 282369)
                        if (cellValue.Contains(activeId))
                        {
                            var parts = cellValue.Split('-');
                            if (parts.Length >= 3)
                            {
                                // Asume formato: SECTOR - ID - NIVEL
                                key = $"{parts[0].Trim().ToUpper()}|{parts[2].Trim().ToUpper()}";
                                valid = true;
                                rowHasActiveId = true;
                            }
                            else
                            {
                                Log($"   [WARN] Fila {rowIndex}: Celda '{cellValue}' tiene ID pero formato incorrecto.");
                            }
                        }
                    }

                    if (valid)
                    {
                        if (!groups.ContainsKey(key))
                        {
                            groups[key] = new ScheduleGroup
                            {
                                IsSitioLogic = isSitio,
                                Sector = isSitio ? key : key.Split('|')[0],
                                Nivel = isSitio ? null : key.Split('|')[1]
                            };
                        }
                        if (!groups[key].WeekCounts.ContainsKey(week)) groups[key].WeekCounts[week] = 0;
                        groups[key].WeekCounts[week]++;
                    }
                }

                // Solo si encontramos grupos válidos (match de ID) agregamos
                if (groups.Any())
                {
                    Log($"[MATCH] Fila {rowIndex}: '{activityName}' -> Regla: '{matchingRule.ActivityKeywords}'");
                    if (rowHasActiveId) Log($"       -> ¡ID {activeId} ENCONTRADO!");

                    var scheduleData = new ScheduleData
                    {
                        ActivityName = activityName,
                        MatchingRule = matchingRule
                    };
                    scheduleData.Groups.AddRange(groups.Values);
                    data.Add(scheduleData);
                    matchesFound++;
                }

                rowIndex++;
            }

            Log($"\n[FIN] Total actividades procesadas con ID {activeId}: {matchesFound}");
            return data;
        }

        private string GetCell(IList<object> row, int index)
        {
            return (index >= 0 && index < row.Count) ? row[index]?.ToString()?.Trim() ?? "" : "";
        }

        private bool IsSitio(string function)
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "FiltroPartVereda", "FiltroCajaDesague", "FiltroCajaPluvial",
                "FiltroBuzon", "FiltroPisoSitioCemento"
            }.Contains(function);
        }

        private string GetWeek(int offset)
        {
            // Ajustar mapeo de columnas a semanas
            if (offset <= 5) return "SEM 01";
            if (offset <= 11) return "SEM 02";
            if (offset <= 17) return "SEM 03";
            if (offset <= 23) return "SEM 04";
            return null;
        }
    }
}