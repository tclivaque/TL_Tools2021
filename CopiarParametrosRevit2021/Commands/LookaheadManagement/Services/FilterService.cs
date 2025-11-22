using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using TL_Tools2021.Commands.LookaheadManagement.Models;

namespace TL_Tools2021.Commands.LookaheadManagement.Services
{
    public class FilterService
    {
        private readonly ParameterService _paramService;
        private readonly double _ancho14;
        private readonly double _ancho19;
        private readonly double _tolerancia;

        private readonly HashSet<string> _kwPorcelanato = new HashSet<string>
        {
            "porcelanato", "gres", "vinilico", "vinílico"
        };

        private readonly HashSet<string> _kwCeramico = new HashSet<string>
        {
            "ceramico", "cerámico"
        };

        private readonly HashSet<string> _kwCemento = new HashSet<string>
        {
            "semipulido"
        };

        private readonly HashSet<string> _kwSitioCemento = new HashSet<string>
        {
            "piso de concreto semipulido",
            "piso de cemento pulido",
            "piso de cemento semipulido"
        };

        private readonly HashSet<string> _kwVereda = new HashSet<string>
        {
            "piso", "rampa", "uñas", "solado"
        };

        private readonly HashSet<string> _kwDesague = new HashSet<string>
        {
            "desague", "desagüe"
        };

        private readonly HashSet<string> _kwPluvial = new HashSet<string>
        {
            "pluvial"
        };

        private readonly HashSet<string> _kwBuzon = new HashSet<string>
        {
            "buzon", "buzón"
        };

        private readonly HashSet<string> _kwEscalera = new HashSet<string>
        {
            "escalera"
        };

        public FilterService(ParameterService paramService, double ancho14, double ancho19, double tolerancia)
        {
            _paramService = paramService;
            _ancho14 = ancho14;
            _ancho19 = ancho19;
            _tolerancia = tolerancia;
        }

        public bool ApplyFilter(Element element, ActivityRule rule)
        {
            if (string.IsNullOrWhiteSpace(rule.FuncionFiltroEspecial))
            {
                if (string.IsNullOrWhiteSpace(rule.ParametroFiltro) || !rule.KeywordsFiltro.Any())
                    return true;

                return RunGenericFilter(element, rule.ParametroFiltro, rule.KeywordsFiltro);
            }

            switch (rule.FuncionFiltroEspecial)
            {
                case "FiltroBloqueta14":
                    return RunBloquetaFilter(element, _ancho14);

                case "FiltroBloqueta19":
                    return RunBloquetaFilter(element, _ancho19);

                case "FiltroMaterialPorcelanato":
                    return RunMaterialFilter(element, _kwPorcelanato);

                case "FiltroMaterialCeramico":
                    return RunMaterialFilter(element, _kwCeramico);

                case "FiltroMaterialCemento":
                    return RunMaterialFilter(element, _kwCemento);

                case "FiltroPisoSitioCemento":
                    return RunDescriptionFilter(element, "Assembly Description", _kwSitioCemento);

                case "FiltroDefault":
                    return true;

                case "FiltroEscaleraAmbiente":
                    return RunEscaleraFilter(element);

                case "FiltroPartVereda":
                    return RunDescriptionFilter(element, "DESCRIPTION PARTS", _kwVereda);

                case "FiltroCajaDesague":
                    return RunDescriptionFilter(element, "Assembly Description", _kwDesague);

                case "FiltroCajaPluvial":
                    return RunDescriptionFilter(element, "Assembly Description", _kwPluvial);

                case "FiltroBuzon":
                    return RunDescriptionFilter(element, "Assembly Description", _kwBuzon);
            }

            return false;
        }

        private bool RunEscaleraFilter(Element element)
        {
            int categoryId = element.Category.Id.IntegerValue;

            if (categoryId == (int)BuiltInCategory.OST_Stairs)
                return true;

            if (categoryId == (int)BuiltInCategory.OST_Floors)
                return RunDescriptionFilter(element, "AMBIENTE", _kwEscalera);

            return false;
        }

        private bool RunBloquetaFilter(Element element, double targetWidth)
        {
            string description = _paramService.GetParameterValue(element, "Assembly Description").ToLower();

            if (!description.Contains("bloqueta") && !description.Contains("dintel"))
                return false;

            double? width = _paramService.GetElementWidth(element);
            return width != null && Math.Abs(width.Value - targetWidth) <= _tolerancia;
        }

        private bool RunMaterialFilter(Element element, HashSet<string> keywords)
        {
            return _paramService.GetFloorMaterialNames(element)
                .Any(name => keywords.Any(kw => name.Contains(kw)));
        }

        private bool RunDescriptionFilter(Element element, string paramName, HashSet<string> keywords)
        {
            string value = _paramService.GetParameterValue(element, paramName).ToLower();
            return keywords.Any(kw => value.Contains(kw));
        }

        private bool RunGenericFilter(Element element, string paramName, List<string> keywords)
        {
            string value = _paramService.GetParameterValue(element, paramName).ToLower();
            return !string.IsNullOrEmpty(value) && keywords.Any(kw => value.Contains(kw));
        }
    }
}
