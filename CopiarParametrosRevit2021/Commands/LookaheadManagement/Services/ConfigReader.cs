using Autodesk.Revit.DB;
using CopiarParametrosRevit2021.Commands.LookaheadManagement.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement.Services
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
            if (string.IsNullOrEmpty(text))
                return text;

            var normalized = text.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (char c in normalized)
            {
                var category = CharUnicodeInfo.GetUnicodeCategory(c);
                if (category != UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }

            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        public List<ActivityRule> ReadConfigRules()
        {
            var rules = new List<ActivityRule>();
            var values = _sheetsService.ReadData(_spreadsheetId, $"'{_configSheetName}'!A2:G");

            foreach (var row in values)
            {
                if (row.Count == 0 || string.IsNullOrWhiteSpace(GetCell(row, 0)))
                    continue;

                var rule = new ActivityRule
                {
                    ActivityKeywords = GetCell(row, 0).ToLower(),
                    Disciplines = GetCell(row, 1).Split(',')
                        .Select(d => d.Trim().ToUpper()).ToList(),
                    RawCategories = GetCell(row, 2),
                    FuncionFiltroEspecial = GetCell(row, 3),
                    ParametroFiltro = GetCell(row, 4),
                    KeywordsFiltro = GetCell(row, 5).Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(k => k.Trim().ToLower()).ToList(),
                    FuncionOrden = GetCell(row, 6)
                };

                rule.Categories = rule.RawCategories
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(c =>
                    {
                        try
                        {
                            return (BuiltInCategory)Enum.Parse(typeof(BuiltInCategory), c.Trim(), true);
                        }
                        catch
                        {
                            return BuiltInCategory.INVALID;
                        }
                    })
                    .Where(bic => bic != BuiltInCategory.INVALID)
                    .ToList();

                if (rule.Categories.Any())
                {
                    rules.Add(rule);
                }
            }

            return rules;
        }

        public List<ScheduleData> ReadScheduleData(List<ActivityRule> rules, string activeId)
        {
            var data = new List<ScheduleData>();
            var headers = _sheetsService.ReadData(_spreadsheetId, $"'{_scheduleSheetName}'!1:1");

            if (headers == null || headers.Count == 0)
            {
                return data;
            }

            var headerRow = headers[0];
            int activityColumn = -1, startColumn = -1;

            for (int i = 0; i < headerRow.Count; i++)
            {
                string headerValue = headerRow[i]?.ToString()?.Trim() ?? "";
                if (headerValue.Equals("DESCRIPCIÓN DE LA ACTIVIDAD", StringComparison.OrdinalIgnoreCase))
                {
                    activityColumn = i;
                    startColumn = i + 1;
                    break;
                }
            }

            if (activityColumn == -1)
            {
                throw new Exception("Columna 'DESCRIPCIÓN DE LA ACTIVIDAD' no encontrada");
            }

            var rows = _sheetsService.ReadData(_spreadsheetId, $"'{_scheduleSheetName}'!A2:ZZ");

            int rowIndex = 0;
            foreach (var row in rows)
            {
                rowIndex++;

                if (row.Count <= activityColumn)
                    continue;

                string activityName = GetCell(row, activityColumn).ToLower();
                if (string.IsNullOrWhiteSpace(activityName))
                    continue;

                string activityNameNormalized = RemoveAccents(activityName);

                var matchingRule = rules
                    .Where(r => activityNameNormalized.Contains(RemoveAccents(r.ActivityKeywords)))
                    .OrderByDescending(r => r.ActivityKeywords.Length)
                    .FirstOrDefault();

                if (matchingRule == null)
                    continue;

                var scheduleData = new ScheduleData
                {
                    ActivityName = activityName,
                    MatchingRule = matchingRule
                };

                bool isSitio = IsSitio(matchingRule.FuncionFiltroEspecial) ||
                    activityNameNormalized.Contains("exterior") ||
                    activityNameNormalized.Contains("podotactil");

                var groups = new Dictionary<string, ScheduleGroup>();

                for (int i = startColumn; i < row.Count && i < 30; i++)
                {
                    string value = GetCell(row, i);
                    if (string.IsNullOrWhiteSpace(value) ||
                        value.Equals("null", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string week = GetWeek(i - startColumn);
                    if (week == null)
                        continue;

                    string key = "";
                    bool valid = false;

                    if (isSitio)
                    {
                        if (value.ToUpper().StartsWith("PLAT"))
                        {
                            key = value.Replace(" ", "").ToUpper();
                            valid = true;
                        }
                    }
                    else
                    {
                        if (value.Contains(activeId))
                        {
                            var parts = value.Split('-');
                            if (parts.Length >= 3)
                            {
                                key = $"{parts[0].Trim().ToUpper()}|{parts[2].Trim().ToUpper()}";
                                valid = true;
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

                        if (!groups[key].WeekCounts.ContainsKey(week))
                            groups[key].WeekCounts[week] = 0;

                        groups[key].WeekCounts[week]++;
                    }
                }

                if (groups.Any())
                {
                    scheduleData.Groups.AddRange(groups.Values);
                    data.Add(scheduleData);
                }
            }

            return data;
        }

        private string GetCell(IList<object> row, int index)
        {
            return (index >= 0 && index < row.Count) ?
                row[index]?.ToString()?.Trim() ?? "" : "";
        }

        private bool IsSitio(string function)
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "FiltroPartVereda",
                "FiltroCajaDesague",
                "FiltroCajaPluvial",
                "FiltroBuzon",
                "FiltroPisoSitioCemento"
            }.Contains(function);
        }

        private string GetWeek(int offset)
        {
            if (offset <= 5) return "SEM 01";
            if (offset <= 11) return "SEM 02";
            if (offset <= 17) return "SEM 03";
            if (offset <= 23) return "SEM 04";
            return null;
        }
    }
}
