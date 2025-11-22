using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TL_Tools2021.Commands.LookaheadManagement.Models;

namespace TL_Tools2021.Commands.LookaheadManagement.Services
{
    public class ConfigReader
    {
        private readonly GoogleSheetsService _sheetsService;
        private readonly string _spreadsheetId;
        private readonly string _scheduleSheetName;
        private readonly string _configSheetName;

        // DEBUG
        private readonly StringBuilder _debugLog = new StringBuilder();
        private readonly string _debugFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ConfigReader_Debug.txt");

        public ConfigReader(GoogleSheetsService sheetsService, string spreadsheetId,
            string scheduleSheetName, string configSheetName)
        {
            _sheetsService = sheetsService;
            _spreadsheetId = spreadsheetId;
            _scheduleSheetName = scheduleSheetName;
            _configSheetName = configSheetName;

            _debugLog.AppendLine($"=== DEBUG CONFIG READER - {DateTime.Now} ===");
        }

        private void Log(string message)
        {
            _debugLog.AppendLine(message);
        }

        public void SaveDebugLog()
        {
            File.WriteAllText(_debugFilePath, _debugLog.ToString());
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
                    rules.Add(rule);
            }

            return rules;
        }

        public List<ScheduleData> ReadScheduleData(List<ActivityRule> rules, string activeId)
        {
            Log($"\n=== ReadScheduleData ===");
            Log($"ActiveId: '{activeId}'");
            Log($"Total reglas recibidas: {rules.Count}");

            var data = new List<ScheduleData>();
            var headers = _sheetsService.ReadData(_spreadsheetId, $"'{_scheduleSheetName}'!1:1");

            if (headers == null || headers.Count == 0)
            {
                Log("ERROR: headers null o vacío");
                return data;
            }

            var headerRow = headers[0];
            Log($"Headers encontrados: {headerRow.Count} columnas");

            int activityColumn = -1, startColumn = -1;

            for (int i = 0; i < headerRow.Count; i++)
            {
                string headerValue = headerRow[i]?.ToString()?.Trim() ?? "";
                if (headerValue.Equals("DESCRIPCIÓN DE LA ACTIVIDAD", StringComparison.OrdinalIgnoreCase))
                {
                    activityColumn = i;
                    startColumn = i + 1;
                    Log($"Columna 'DESCRIPCIÓN DE LA ACTIVIDAD' encontrada en índice: {i}");
                    break;
                }
            }

            if (activityColumn == -1)
            {
                Log("ERROR: Columna 'DESCRIPCIÓN DE LA ACTIVIDAD' no encontrada");
                Log($"Headers disponibles: {string.Join(", ", headerRow.Select(h => $"'{h}'"))}");
                throw new Exception("Columna 'DESCRIPCIÓN DE LA ACTIVIDAD' no encontrada");
            }

            var rows = _sheetsService.ReadData(_spreadsheetId, $"'{_scheduleSheetName}'!A2:ZZ");
            Log($"Filas leídas del LOOKAHEAD: {rows?.Count ?? 0}");

            int rowIndex = 0;
            foreach (var row in rows)
            {
                rowIndex++;

                if (row.Count <= activityColumn)
                {
                    Log($"Fila {rowIndex}: SKIP - row.Count ({row.Count}) <= activityColumn ({activityColumn})");
                    continue;
                }

                string activityName = GetCell(row, activityColumn).ToLower();
                if (string.IsNullOrWhiteSpace(activityName))
                {
                    Log($"Fila {rowIndex}: SKIP - activityName vacío");
                    continue;
                }

                Log($"\nFila {rowIndex}: Actividad = '{activityName}'");

                // Buscar regla que coincida
                var matchingRule = rules.FirstOrDefault(r => activityName.Contains(r.ActivityKeywords));
                if (matchingRule == null)
                {
                    Log($"  NO MATCH: Ninguna regla coincide. Keywords buscados:");
                    foreach (var r in rules)
                    {
                        Log($"    - '{r.ActivityKeywords}' contenido en '{activityName}'? {activityName.Contains(r.ActivityKeywords)}");
                    }
                    continue;
                }

                Log($"  MATCH con regla: '{matchingRule.ActivityKeywords}' (Disciplinas: {string.Join(",", matchingRule.Disciplines)})");

                var scheduleData = new ScheduleData
                {
                    ActivityName = activityName,
                    MatchingRule = matchingRule
                };

                bool isSitio = IsSitio(matchingRule.FuncionFiltroEspecial) ||
                    activityName.Contains("exterior") ||
                    activityName.Contains("podotactil");

                Log($"  IsSitio={isSitio} (FuncionFiltro='{matchingRule.FuncionFiltroEspecial}', exterior={activityName.Contains("exterior")}, podotactil={activityName.Contains("podotactil")})");

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

                    Log($"    Columna {i}: valor='{value}', week={week}");

                    string key = "";
                    bool valid = false;

                    if (isSitio)
                    {
                        // Lógica SITIO: Ignora ID del archivo, busca PLAT
                        if (value.ToUpper().StartsWith("PLAT"))
                        {
                            key = value.Replace(" ", "").ToUpper();
                            valid = true;
                            Log($"      SITIO: PLAT encontrado -> key='{key}'");
                        }
                        else
                        {
                            Log($"      SITIO: valor '{value}' NO empieza con PLAT");
                        }
                    }
                    else
                    {
                        // Lógica ACTIVO: Debe contener el ID del archivo
                        if (value.Contains(activeId))
                        {
                            var parts = value.Split('-');
                            if (parts.Length >= 3)
                            {
                                key = $"{parts[0].Trim().ToUpper()}|{parts[2].Trim().ToUpper()}";
                                valid = true;
                                Log($"      ACTIVO: ID '{activeId}' encontrado -> key='{key}'");
                            }
                            else
                            {
                                Log($"      ACTIVO: valor '{value}' contiene ID pero parts.Length={parts.Length} < 3");
                            }
                        }
                        else
                        {
                            Log($"      ACTIVO: valor '{value}' NO contiene activeId '{activeId}'");
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

                Log($"  Grupos generados: {groups.Count}");
                foreach (var g in groups)
                {
                    Log($"    - Key='{g.Key}', Sector='{g.Value.Sector}', Nivel='{g.Value.Nivel}', Weeks={string.Join(",", g.Value.WeekCounts.Select(w => $"{w.Key}:{w.Value}"))}");
                }

                if (groups.Any())
                {
                    scheduleData.Groups.AddRange(groups.Values);
                    data.Add(scheduleData);
                    Log($"  AGREGADO a data");
                }
                else
                {
                    Log($"  NO AGREGADO: sin grupos válidos");
                }
            }

            Log($"\n=== Fin ReadScheduleData: {data.Count} ScheduleData generados ===");
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
