using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using TL_Tools2021.Commands.LookaheadManagement.Models;

namespace TL_Tools2021.Commands.LookaheadManagement.Services
{
    public class LookAheadProcessor
    {
        private readonly ParameterService _paramService;
        private readonly FilterService _filterService;
        private readonly Dictionary<BuiltInCategory, List<Element>> _cache;
        private readonly HashSet<ElementId> _assignedIds = new HashSet<ElementId>();

        // DEBUG: IDs específicos a rastrear
        private readonly HashSet<int> _debugIds = new HashSet<int> { 24161487, 13886589 };
        private readonly StringBuilder _debugLog = new StringBuilder();
        private readonly string _debugFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "LookAhead_Debug.txt");

        public LookAheadProcessor(ParameterService paramService, FilterService filterService,
            Dictionary<BuiltInCategory, List<Element>> cache)
        {
            _paramService = paramService;
            _filterService = filterService;
            _cache = cache;

            _debugLog.AppendLine($"=== DEBUG LOOKAHEAD - {DateTime.Now} ===");
            _debugLog.AppendLine($"IDs a rastrear: {string.Join(", ", _debugIds)}");
            _debugLog.AppendLine();
        }

        private void Log(string message)
        {
            _debugLog.AppendLine(message);
        }

        private void LogElement(Element elem, string context)
        {
            if (_debugIds.Contains(elem.Id.IntegerValue))
            {
                Log($"[ID:{elem.Id.IntegerValue}] {context}");
            }
        }

        public void SaveDebugLog()
        {
            File.WriteAllText(_debugFilePath, _debugLog.ToString());
        }

        public void AssignWeeks(List<ScheduleData> data, string discipline)
        {
            Log($"--- AssignWeeks para disciplina: {discipline} ---");

            var tasks = data.Where(t => t.MatchingRule.Disciplines
                .Contains(discipline, StringComparer.OrdinalIgnoreCase)).ToList();

            Log($"Tareas encontradas para {discipline}: {tasks.Count}");

            foreach (var task in tasks)
            {
                var rule = task.MatchingRule;
                Log($"\n>> Procesando actividad: {task.ActivityName}");
                Log($"   Keywords: {rule.ActivityKeywords}");
                Log($"   Categorías: {rule.RawCategories}");
                Log($"   FuncionFiltroEspecial: {rule.FuncionFiltroEspecial}");
                Log($"   ParametroFiltro: {rule.ParametroFiltro}");
                Log($"   KeywordsFiltro: {string.Join(",", rule.KeywordsFiltro)}");

                var candidates = rule.Categories
                    .Where(c => _cache.ContainsKey(c))
                    .SelectMany(c => _cache[c])
                    .ToList();

                Log($"   Candidatos encontrados: {candidates.Count}");

                // Verificar si los IDs de debug están en candidatos
                foreach (var debugId in _debugIds)
                {
                    var found = candidates.FirstOrDefault(c => c.Id.IntegerValue == debugId);
                    if (found != null)
                    {
                        Log($"   [ID:{debugId}] ENCONTRADO en candidatos - Categoría: {found.Category?.Name}");
                    }
                }

                if (!candidates.Any())
                    continue;

                foreach (var group in task.Groups)
                {
                    var filtered = new List<Element>();
                    string groupSectorClean = StringUtils.Clean(group.Sector);
                    string groupNivelClean = StringUtils.Clean(group.Nivel);

                    Log($"\n   >> Grupo: Sector={group.Sector} ({groupSectorClean}), Nivel={group.Nivel} ({groupNivelClean}), IsSitio={group.IsSitioLogic}");
                    Log($"      WeekCounts: {string.Join(", ", group.WeekCounts.Select(w => $"{w.Key}={w.Value}"))}");

                    foreach (var elem in candidates)
                    {
                        bool isDebugElement = _debugIds.Contains(elem.Id.IntegerValue);

                        if (_assignedIds.Contains(elem.Id))
                        {
                            if (isDebugElement) LogElement(elem, "SKIP: Ya asignado en esta ejecución");
                            continue;
                        }

                        string ejecutado = _paramService.GetParameterValue(elem, "EJECUTADO");
                        if (ejecutado == "1")
                        {
                            if (isDebugElement) LogElement(elem, $"SKIP: EJECUTADO=1");
                            continue;
                        }

                        bool filterResult = _filterService.ApplyFilter(elem, rule);
                        if (!filterResult)
                        {
                            if (isDebugElement)
                            {
                                string assemblyDesc = _paramService.GetParameterValue(elem, "Assembly Description");
                                LogElement(elem, $"SKIP: Filtro NO pasó. AssemblyDesc='{assemblyDesc}'");
                            }
                            continue;
                        }

                        if (isDebugElement) LogElement(elem, "OK: Filtro pasó");

                        string elemSector = _paramService.GetParameterValue(elem, "SECTOR");
                        string elemSectorClean = StringUtils.Clean(elemSector);
                        bool match = false;

                        if (group.IsSitioLogic)
                        {
                            match = (elemSectorClean == groupSectorClean);
                            if (isDebugElement)
                            {
                                LogElement(elem, $"SITIO: ElemSector='{elemSector}' ({elemSectorClean}) vs GroupSector='{groupSectorClean}' => Match={match}");
                            }
                        }
                        else
                        {
                            string elemNivel = _paramService.GetParameterValue(elem, "NIVEL");
                            string elemNivelClean = StringUtils.Clean(elemNivel);
                            match = (elemSectorClean == groupSectorClean && elemNivelClean == groupNivelClean);
                            if (isDebugElement)
                            {
                                LogElement(elem, $"ACTIVO: Sector='{elemSector}'({elemSectorClean}) vs '{groupSectorClean}', Nivel='{elemNivel}'({elemNivelClean}) vs '{groupNivelClean}' => Match={match}");
                            }
                        }

                        if (match)
                        {
                            filtered.Add(elem);
                            if (isDebugElement) LogElement(elem, "AGREGADO a filtered");
                        }
                    }

                    Log($"      Elementos filtrados: {filtered.Count}");

                    if (!filtered.Any())
                        continue;

                    var ordered = OrderElements(filtered, rule.FuncionOrden);
                    var distributed = DistributeElements(ordered, group.WeekCounts);

                    foreach (var pair in distributed)
                    {
                        bool isDebugElement = _debugIds.Contains(pair.Item1.Id.IntegerValue);
                        if (_paramService.SetParameterValue(pair.Item1, "LOOK AHEAD", pair.Item2))
                        {
                            _assignedIds.Add(pair.Item1.Id);
                            if (isDebugElement) LogElement(pair.Item1, $"ASIGNADO: LOOK AHEAD = {pair.Item2}");
                        }
                    }
                }
            }
        }

        public void ApplyExecuteAndSALogic(List<ActivityRule> rules)
        {
            Log($"\n=== ApplyExecuteAndSALogic ===");
            Log($"Total reglas: {rules.Count}");
            Log($"IDs ya asignados: {_assignedIds.Count}");

            var allElements = rules
                .SelectMany(r => r.Categories)
                .Distinct()
                .Where(c => _cache.ContainsKey(c))
                .SelectMany(c => _cache[c])
                .Distinct(new ElementIdComparer())
                .ToList();

            Log($"Total elementos a evaluar: {allElements.Count}");

            // Verificar si los IDs de debug están en allElements
            foreach (var debugId in _debugIds)
            {
                var found = allElements.FirstOrDefault(e => e.Id.IntegerValue == debugId);
                if (found != null)
                {
                    Log($"[ID:{debugId}] ENCONTRADO en allElements - Categoría: {found.Category?.Name}");
                }
                else
                {
                    Log($"[ID:{debugId}] NO ENCONTRADO en allElements");
                }
            }

            foreach (var elem in allElements)
            {
                bool isDebugElement = _debugIds.Contains(elem.Id.IntegerValue);

                // Verificar relevancia con cada regla
                ActivityRule matchingRule = null;
                foreach (var r in rules)
                {
                    bool categoryMatch = r.Categories.Contains((BuiltInCategory)elem.Category.Id.IntegerValue);
                    bool filterMatch = _filterService.ApplyFilter(elem, r);

                    if (isDebugElement)
                    {
                        LogElement(elem, $"Evaluando regla '{r.ActivityKeywords}': CategoryMatch={categoryMatch}, FilterMatch={filterMatch}");
                    }

                    if (categoryMatch && filterMatch)
                    {
                        matchingRule = r;
                        break;
                    }
                }

                bool relevant = matchingRule != null;

                if (!relevant)
                {
                    if (isDebugElement) LogElement(elem, "NO RELEVANTE: ninguna regla coincide");
                    continue;
                }

                if (isDebugElement) LogElement(elem, $"RELEVANTE: coincide con '{matchingRule.ActivityKeywords}'");

                string ejecutado = _paramService.GetParameterValue(elem, "EJECUTADO");
                bool wasAssigned = _assignedIds.Contains(elem.Id);

                if (isDebugElement)
                {
                    LogElement(elem, $"EJECUTADO='{ejecutado}', WasAssigned={wasAssigned}");
                }

                if (ejecutado == "1")
                {
                    _paramService.SetParameterValue(elem, "LOOK AHEAD", "EJECUTADO");
                    _paramService.SetParameterValue(elem, "EN CURSO", 0);
                    if (isDebugElement) LogElement(elem, "SET: LOOK AHEAD = EJECUTADO");
                }
                else if (!wasAssigned)
                {
                    // Siempre asignar "SA" si no fue asignado en esta ejecución
                    // Esto sobrescribe valores anteriores (SEM XX) que ya no están en la programación actual
                    _paramService.SetParameterValue(elem, "LOOK AHEAD", "SA");
                    if (isDebugElement) LogElement(elem, "SET: LOOK AHEAD = SA (no fue asignado en esta ejecución)");
                }
                else
                {
                    if (isDebugElement) LogElement(elem, "NO CAMBIO: ya fue asignado previamente");
                }
            }

            Log($"\n=== Fin ApplyExecuteAndSALogic ===");
        }

        private List<Element> OrderElements(List<Element> elements, string orderFunction)
        {
            if (orderFunction == "Geo_YX")
            {
                return elements
                    .OrderBy(e => _paramService.GetElementCenter(e).Y)
                    .ThenByDescending(e => _paramService.GetElementCenter(e).X)
                    .ToList();
            }

            if (orderFunction == "Vertical_Z_Desc")
            {
                return elements
                    .OrderByDescending(e => _paramService.GetElementCenter(e).Z)
                    .ToList();
            }

            return elements;
        }

        private List<Tuple<Element, string>> DistributeElements(
            List<Element> elements, Dictionary<string, int> weekCounts)
        {
            var result = new List<Tuple<Element, string>>();
            int totalElements = elements.Count;

            if (totalElements == 0)
                return result;

            var weeks = weekCounts.Keys
                .OrderBy(w => int.Parse(Regex.Match(w, @"\d+").Value))
                .ToList();

            if (totalElements == 1)
            {
                result.Add(new Tuple<Element, string>(elements[0], weeks.Last()));
                return result;
            }

            double total = weekCounts.Values.Sum();
            int assigned = 0;

            for (int i = 0; i < weeks.Count; i++)
            {
                string week = weeks[i];
                double share = weekCounts[week];

                int num = (i == weeks.Count - 1) ?
                    (totalElements - assigned) :
                    (int)Math.Round((totalElements * share) / total);

                if (assigned + num > totalElements)
                    num = totalElements - assigned;

                for (int j = 0; j < num; j++)
                {
                    if (assigned < totalElements)
                        result.Add(new Tuple<Element, string>(elements[assigned++], week));
                }
            }

            return result;
        }

        private class ElementIdComparer : IEqualityComparer<Element>
        {
            public bool Equals(Element x, Element y)
            {
                return x.Id == y.Id;
            }

            public int GetHashCode(Element obj)
            {
                return obj.Id.GetHashCode();
            }
        }
    }
}
