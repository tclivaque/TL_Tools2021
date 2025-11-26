using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using TL_Tools2021.Commands.LookaheadManagement.Models;
using CopiarParametrosRevit2021.Helpers;

namespace TL_Tools2021.Commands.LookaheadManagement.Services
{
    public class ConfigReader
    {
        private readonly GoogleSheetsService _sheetsService;
        private readonly string _spreadsheetId;
        private readonly string _scheduleSheetName;
        private readonly string _configSheetName;

        public ConfigReader(GoogleSheetsService sheetsService, string spreadsheetId,
            string scheduleSheetName, string configSheetName)
        {
            _sheetsService = sheetsService;
            _spreadsheetId = spreadsheetId;
            _scheduleSheetName = scheduleSheetName;
            _configSheetName = configSheetName;
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
            var rules = new List<ActivityRule>();

            try
            {
                var values = _sheetsService.ReadData(_spreadsheetId, $"'{_configSheetName}'!A2:G");

                foreach (var row in values)
                {
                    string keyword = GetCell(row, 0);

                    if (string.IsNullOrWhiteSpace(keyword))
                        continue;

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
                        catch { /* Ignorar categorías inválidas */ }
                    }

                    if (rule.Categories.Any())
                    {
                        rules.Add(rule);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error leyendo reglas: {ex.Message}", ex);
            }

            return rules;
        }

        public List<ScheduleData> ReadScheduleData(List<ActivityRule> rules)
        {
            var data = new List<ScheduleData>();

            // 1. Leer Encabezados
            var headers = _sheetsService.ReadData(_spreadsheetId, $"'{_scheduleSheetName}'!1:1");
            if (headers == null || headers.Count == 0)
                return data;

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
                    break;
                }
            }

            if (activityColumn == -1)
                throw new Exception("Columna 'DESCRIPCIÓN DE LA ACTIVIDAD' no encontrada");

            // 2. Leer Datos (Hasta columna ZZ para asegurar rango)
            var rows = _sheetsService.ReadData(_spreadsheetId, $"'{_scheduleSheetName}'!A2:ZZ");

            foreach (var row in rows)
            {
                if (row.Count <= activityColumn) continue;

                string activityName = GetCell(row, activityColumn).ToLower();
                if (string.IsNullOrWhiteSpace(activityName)) continue;

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
                    continue;

                // --- REGLA ENCONTRADA ---
                bool isSitio = IsSitio(matchingRule.FuncionFiltroEspecial) ||
                               activityNameNormalized.Contains("exterior") ||
                               activityNameNormalized.Contains("podotactil");

                var groups = new Dictionary<string, ScheduleGroup>();

                // --- PROCESAR SEMANAS ---
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
                        // Lógica ACTIVO (Procesa cualquier celda con formato: SECTOR - ID - NIVEL)
                        var parts = cellValue.Split('-');
                        if (parts.Length >= 3)
                        {
                            // Asume formato: SECTOR - ID - NIVEL
                            key = $"{parts[0].Trim().ToUpper()}|{parts[2].Trim().ToUpper()}";
                            valid = true;
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

                // Solo si encontramos grupos válidos agregamos
                if (groups.Any())
                {
                    var scheduleData = new ScheduleData
                    {
                        ActivityName = activityName,
                        MatchingRule = matchingRule
                    };
                    scheduleData.Groups.AddRange(groups.Values);
                    data.Add(scheduleData);
                }
            }

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