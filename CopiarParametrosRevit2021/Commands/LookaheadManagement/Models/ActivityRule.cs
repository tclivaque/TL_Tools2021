using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement.Models
{
    public class ActivityRule
    {
        public string ActivityKeywords { get; set; }
        public List<string> Disciplines { get; set; }
        public string RawCategories { get; set; }
        public List<BuiltInCategory> Categories { get; set; }
        public string FuncionFiltroEspecial { get; set; }
        public string ParametroFiltro { get; set; }
        public List<string> KeywordsFiltro { get; set; }
        public string FuncionOrden { get; set; }

        public ActivityRule()
        {
            Disciplines = new List<string>();
            Categories = new List<BuiltInCategory>();
            KeywordsFiltro = new List<string>();
        }
    }
}