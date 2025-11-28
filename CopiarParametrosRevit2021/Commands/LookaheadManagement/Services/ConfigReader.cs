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

        public ConfigReader(GoogleSheetsService sheetsService, string spreadsheetId,
            string scheduleSheetName)
        {
            _sheetsService = sheetsService;
            _spreadsheetId = spreadsheetId;
            _scheduleSheetName = scheduleSheetName;
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
            // REGLAS HARDCODEADAS - Ya no se lee de CONFIG_ACTIVIDADES
            var rules = new List<ActivityRule>
            {
                // === ARQUITECTURA - PISOS ===
                new ActivityRule
                {
                    ActivityKeywords = "porcelanato",
                    Disciplines = new List<string> { "AR" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_Floors },
                    RawCategories = "OST_Floors",
                    FuncionFiltroEspecial = "FiltroMaterialPorcelanato",
                    FuncionOrden = "Geo_YX"
                },
                new ActivityRule
                {
                    ActivityKeywords = "ceramico",
                    Disciplines = new List<string> { "AR" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_Floors },
                    RawCategories = "OST_Floors",
                    FuncionFiltroEspecial = "FiltroMaterialCeramico",
                    FuncionOrden = "Geo_YX"
                },
                new ActivityRule
                {
                    ActivityKeywords = "cemento semipulido",
                    Disciplines = new List<string> { "AR" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_Floors },
                    RawCategories = "OST_Floors",
                    FuncionFiltroEspecial = "FiltroMaterialCemento",
                    FuncionOrden = "Geo_YX"
                },
                new ActivityRule
                {
                    ActivityKeywords = "pisos de cemento semipulido",
                    Disciplines = new List<string> { "AR" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_Floors },
                    RawCategories = "OST_Floors",
                    FuncionFiltroEspecial = "FiltroPisoSitioCemento",
                    FuncionOrden = "Geo_YX"
                },

                // === ARQUITECTURA - MUROS ===
                new ActivityRule
                {
                    ActivityKeywords = "bloqueta 14",
                    Disciplines = new List<string> { "AR" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_Walls, BuiltInCategory.OST_StructuralFraming },
                    RawCategories = "OST_Walls,OST_StructuralFraming",
                    FuncionFiltroEspecial = "FiltroBloqueta14",
                    FuncionOrden = "Geo_YX"
                },
                new ActivityRule
                {
                    ActivityKeywords = "bloqueta 19",
                    Disciplines = new List<string> { "AR" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_Walls, BuiltInCategory.OST_StructuralFraming },
                    RawCategories = "OST_Walls,OST_StructuralFraming",
                    FuncionFiltroEspecial = "FiltroBloqueta19",
                    FuncionOrden = "Geo_YX"
                },

                // === ARQUITECTURA - ESCALERAS ===
                new ActivityRule
                {
                    ActivityKeywords = "escalera",
                    Disciplines = new List<string> { "AR" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_Stairs, BuiltInCategory.OST_Floors },
                    RawCategories = "OST_Stairs,OST_Floors",
                    FuncionFiltroEspecial = "FiltroEscaleraAmbiente",
                    FuncionOrden = "Vertical_Z_Desc"
                },

                // === ARQUITECTURA - PODOTACTILES ===
                new ActivityRule
                {
                    ActivityKeywords = "podotactil",
                    Disciplines = new List<string> { "AR" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_GenericModel },
                    RawCategories = "OST_GenericModel",
                    FuncionFiltroEspecial = "FiltroPodotactil",
                    FuncionOrden = "Geo_YX"
                },

                // === ELÉCTRICAS - LUMINARIAS ===
                new ActivityRule
                {
                    ActivityKeywords = "luminaria",
                    Disciplines = new List<string> { "EE" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_LightingFixtures },
                    RawCategories = "OST_ElectricalFixtures,OST_LightingFixtures",
                    FuncionFiltroEspecial = "FiltroLuminarias",
                    FuncionOrden = "Geo_YX"
                },

                // === ELÉCTRICAS - TABLEROS ===
                new ActivityRule
                {
                    ActivityKeywords = "tablero",
                    Disciplines = new List<string> { "EE" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_ElectricalEquipment },
                    RawCategories = "OST_ElectricalEquipment",
                    FuncionFiltroEspecial = "FiltroTableros",
                    FuncionOrden = "Geo_YX"
                },

                // === ELÉCTRICAS/COMUNICACIONES - BUZONES ===
                new ActivityRule
                {
                    ActivityKeywords = "buzon",
                    Disciplines = new List<string> { "DT", "EE" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_ElectricalEquipment },
                    RawCategories = "OST_ElectricalEquipment",
                    FuncionFiltroEspecial = "FiltroBuzon",
                    FuncionOrden = "Geo_YX"
                },

                // === SANITARIAS - CAJAS ===
                new ActivityRule
                {
                    ActivityKeywords = "caja",
                    Disciplines = new List<string> { "SA" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_PlumbingFixtures },
                    RawCategories = "OST_PlumbingFixtures",
                    FuncionFiltroEspecial = "FiltroCajaDesague",
                    FuncionOrden = "Geo_YX"
                },
                new ActivityRule
                {
                    ActivityKeywords = "pluvial",
                    Disciplines = new List<string> { "SA" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_PlumbingFixtures },
                    RawCategories = "OST_PlumbingFixtures",
                    FuncionFiltroEspecial = "FiltroCajaPluvial",
                    FuncionOrden = "Geo_YX"
                },

                // === ESTRUCTURAS - VEREDAS ===
                new ActivityRule
                {
                    ActivityKeywords = "vereda",
                    Disciplines = new List<string> { "ES" },
                    Categories = new List<BuiltInCategory> { BuiltInCategory.OST_Parts },
                    RawCategories = "OST_Parts",
                    FuncionFiltroEspecial = "FiltroPartVereda",
                    FuncionOrden = "Geo_YX"
                }
            };

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
