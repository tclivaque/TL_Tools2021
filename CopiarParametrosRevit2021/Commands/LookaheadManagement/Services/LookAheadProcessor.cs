using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
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

                        if (_assignedIds.Contains(elem.Id))
                        {
                            continue;
                        }

                        string ejecutado = _paramService.GetParameterValue(elem, "EJECUTADO");
                        if (ejecutado == "1")
                        {
                            continue;
                        }

                        bool filterResult = _filterService.ApplyFilter(elem, rule);
                        if (!filterResult)
                        {
                            continue;
                        }


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

                    var ordered = OrderElements(filtered, rule.FuncionOrden);
                    var distributed = DistributeElements(ordered, group.WeekCounts);

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

                // Verificar relevancia con cada regla
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

                bool relevant = matchingRule != null;

                if (!relevant)
                {
                    continue;
                }


                string ejecutado = _paramService.GetParameterValue(elem, "EJECUTADO");
                bool wasAssigned = _assignedIds.Contains(elem.Id);


                if (ejecutado == "1")
                {
                    _paramService.SetParameterValue(elem, "LOOK AHEAD", "EJECUTADO");
                    _paramService.SetParameterValue(elem, "EN CURSO", 0);
                }
                else if (!wasAssigned)
                {
                    // Siempre asignar "SA" si no fue asignado en esta ejecución
                    // Esto sobrescribe valores anteriores (SEM XX) que ya no están en la programación actual
                    _paramService.SetParameterValue(elem, "LOOK AHEAD", "SA");
                }
                else
                {
                }
            }

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
