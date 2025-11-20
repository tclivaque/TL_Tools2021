using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using TL_Tools2021.Commands.LookaheadManagement.Models;

namespace TL_Tools2021.Commands.LookaheadManagement.Services
{
    public class ElementCollectorService
    {
        private readonly Document _doc;

        public ElementCollectorService(Document doc)
        {
            _doc = doc;
        }

        public Dictionary<BuiltInCategory, List<Element>> CollectElements(List<ActivityRule> rules)
        {
            var cache = new Dictionary<BuiltInCategory, List<Element>>();
            var categories = rules.SelectMany(r => r.Categories).Distinct();

            foreach (var category in categories)
            {
                if (!cache.ContainsKey(category))
                {
                    cache[category] = new FilteredElementCollector(_doc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType()
                        .ToList();
                }
            }

            return cache;
        }
    }
}
