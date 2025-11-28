using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
// IMPORTANTE: Referencia a los nuevos modelos
using CopiarParametrosRevit2021.Commands.LookaheadManagement.Models;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement.Services
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
            var uniqueCategories = rules
                .SelectMany(r => r.Categories)
                .Distinct()
                .ToList();

            foreach (var cat in uniqueCategories)
            {
                if (cat == BuiltInCategory.INVALID) continue;

                var elements = new FilteredElementCollector(_doc)
                    .OfCategory(cat)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();

                if (elements.Any())
                {
                    cache[cat] = elements;
                }
            }

            return cache;
        }
    }
}