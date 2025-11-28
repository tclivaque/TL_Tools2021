using Autodesk.Revit.DB;
using CopiarParametrosRevit2021.Commands.LookaheadManagement.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement.Services
{
    public class LookAheadProcessor
    {
        private readonly ParameterService _paramService;
        private readonly FilterService _filterService;
        private readonly Dictionary<BuiltInCategory, List<Element>> _cache;
        private readonly HashSet<ElementId> _assignedIds = new HashSet<ElementId>();

        public LookAheadProcessor(ParameterService paramService, FilterService filterService,
            Dictionary<BuiltInCategory, List<Element>> cache)
        {
            _paramService = paramService;
            _filterService = filterService;
            _cache = cache;
        }

        public void AssignWeeks(List<ScheduleData> data, string discipline)
        {
            var tasks = data.Where(t => t.MatchingRule.Disciplines
                .Contains(discipline, StringComparer.OrdinalIgnoreCase)).ToList();

            foreach (var task in tasks)
            {
                var rule = task.MatchingRule;

                var candidates = rule.Categories
                    .Where(c => _cache.ContainsKey(c))
                    .SelectMany(c => _cache[c])
                    .ToList();

                if (!candidates.Any())
                    continue;

                foreach (var group in task.Groups)
                {
                    var filtered = new List<Element>();
                    string groupSectorClean = StringUtils.Clean(group.Sector);
                    string groupNivelClean = StringUtils.Clean(group.Nivel);

                    foreach (var elem in candidates)
                    {
                        // Skip si ya fue asignado
                        if (_assignedIds.Contains(elem.Id))
                            continue;

                        // Skip si ya está ejecutado
                        string ejecutado = _paramService.GetParameterValue(elem, "EJECUTADO");
                        if (ejecutado == "1")
                            continue;

                        // Aplicar filtro
                        bool filterResult = _filterService.ApplyFilter(elem, rule);
                        if (!filterResult)
                            continue;

                        // Verificar SECTOR y NIVEL
                        string elemSector = _paramService.GetParameterValue(elem, "SECTOR");
                        string elemSectorClean = StringUtils.Clean(elemSector);
                        bool match = false;

                        if (group.IsSitioLogic)
                        {
                            match = (elemSectorClean == groupSectorClean);
                        }
                        else
                        {
                            string elemNivel = _paramService.GetParameterValue(elem, "NIVEL");
                            string elemNivelClean = StringUtils.Clean(elemNivel);
                            match = (elemSectorClean == groupSectorClean && elemNivelClean == groupNivelClean);
                        }

                        if (match)
                        {
                            filtered.Add(elem);
                        }
                    }

                    if (!filtered.Any())
                        continue;

                    // Ordenar elementos
                    var ordered = OrderElements(filtered, rule.FuncionOrden);

                    // Distribuir elementos
                    var distributed = DistributeElements(ordered, group.WeekCounts);

                    // Asignar LOOK AHEAD
                    foreach (var pair in distributed)
                    {
                        if (_paramService.SetParameterValue(pair.Item1, "LOOK AHEAD", pair.Item2))
                        {
                            _assignedIds.Add(pair.Item1.Id);
                        }
                    }
                }
            }
        }

        public void ApplyExecuteAndSALogic(List<ActivityRule> rules)
        {
            var allElements = rules
                .SelectMany(r => r.Categories)
                .Distinct()
                .Where(c => _cache.ContainsKey(c))
                .SelectMany(c => _cache[c])
                .Distinct(new ElementIdComparer())
                .ToList();

            foreach (var elem in allElements)
            {
                // Verificar si es relevante (cumple con alguna regla)
                ActivityRule matchingRule = null;
                foreach (var r in rules)
                {
                    bool categoryMatch = r.Categories.Contains((BuiltInCategory)elem.Category.Id.IntegerValue);
                    bool filterMatch = _filterService.ApplyFilter(elem, r);

                    if (categoryMatch && filterMatch)
                    {
                        matchingRule = r;
                        break;
                    }
                }

                // Solo procesar elementos relevantes
                if (matchingRule == null)
                    continue;

                string ejecutado = _paramService.GetParameterValue(elem, "EJECUTADO");
                bool wasAssigned = _assignedIds.Contains(elem.Id);

                if (ejecutado == "1")
                {
                    // Si está ejecutado, marcar como "EJECUTADO"
                    _paramService.SetParameterValue(elem, "LOOK AHEAD", "EJECUTADO");
                    _paramService.SetParameterValue(elem, "EN CURSO", 0);
                }
                else if (!wasAssigned)
                {
                    // Si no fue asignado en esta ejecución, marcar como "SA"
                    _paramService.SetParameterValue(elem, "LOOK AHEAD", "SA");
                }
                // Si fue asignado, no hacer nada (ya tiene su semana)
            }
        }

        private List<Element> OrderElements(List<Element> elements, string orderFunction)
        {
            if (orderFunction == "Geo_YX")
            {
                // Ordenar por Y ascendente, luego X descendente
                return elements
                    .OrderBy(e => _paramService.GetElementCenter(e).Y)
                    .ThenByDescending(e => _paramService.GetElementCenter(e).X)
                    .ToList();
            }

            if (orderFunction == "Vertical_Z_Desc")
            {
                // Ordenar por Z descendente (de arriba hacia abajo)
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

            // Ordenar semanas
            var weeks = weekCounts.Keys
                .OrderBy(w => int.Parse(Regex.Match(w, @"\d+").Value))
                .ToList();

            // CASO ESPECIAL: Si solo hay 1 elemento, asignar a la última semana
            if (totalElements == 1)
            {
                result.Add(new Tuple<Element, string>(elements[0], weeks.Last()));
                return result;
            }

            // CASO NORMAL: Distribución proporcional
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
