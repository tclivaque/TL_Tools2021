using Autodesk.Revit.DB;
using CopiarParametrosRevit2021.Commands.LookaheadManagement.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CopiarParametrosRevit2021.Commands.LookaheadManagement.Services
{
    public class FilterService
    {
        private readonly Document _doc;
        private readonly ParameterService _paramService;
        private readonly double _ancho14;
        private readonly double _ancho19;
        private readonly double _tolerancia;

        // Keywords según Python original
        private static readonly string[] KeywordsPorcelanato = { "porcelanato", "gres", "vinilico", "vinílico" };
        private static readonly string[] KeywordsCeramico = { "ceramico", "cerámico" };
        private static readonly string[] KeywordsCemento = { "semipulido" };
        private static readonly string[] KeywordsSitioCemento = {
            "piso de concreto semipulido",
            "piso de cemento pulido",
            "piso de cemento semipulido"
        };
        private static readonly string[] KeywordsDesague = { "desague", "desagüe" };
        private static readonly string[] KeywordsPluvial = { "pluvial" };
        private static readonly string[] KeywordsBuzon = { "buzon", "buzón" };
        private static readonly string[] KeywordsPartVereda = { "piso", "rampa", "uñas", "solado" };

        public FilterService(Document doc, ParameterService paramService, double ancho14, double ancho19, double tolerancia)
        {
            _doc = doc;
            _paramService = paramService;
            _ancho14 = ancho14;
            _ancho19 = ancho19;
            _tolerancia = tolerancia;
        }

        public bool ApplyFilter(Element elem, ActivityRule rule)
        {
            if (string.IsNullOrEmpty(rule.FuncionFiltroEspecial))
                return true;

            return ApplySpecialFilter(elem, rule.FuncionFiltroEspecial);
        }

        private bool ApplySpecialFilter(Element elem, string functionName)
        {
            switch (functionName)
            {
                // === FILTROS DE PISOS - MATERIALES ===
                case "FiltroMaterialPorcelanato":
                    return FilterFloorByMaterial(elem, KeywordsPorcelanato);

                case "FiltroMaterialCeramico":
                    return FilterFloorByMaterial(elem, KeywordsCeramico);

                case "FiltroMaterialCemento":
                    return FilterFloorByMaterial(elem, KeywordsCemento);

                case "FiltroPisoSitioCemento":
                    return FilterPisoSitioCemento(elem);

                // === FILTROS DE BLOQUETAS ===
                case "FiltroBloqueta14":
                    return FilterBloqueta(elem, _ancho14);

                case "FiltroBloqueta19":
                    return FilterBloqueta(elem, _ancho19);

                // === FILTRO DE ESCALERA ===
                case "FiltroEscaleraAmbiente":
                    // Solo categoría, sin keywords
                    return elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Stairs ||
                           elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors;

                // === ACTIVIDADES NUEVAS HARDCODED ===
                case "FiltroLuminarias":
                    return FilterByAssemblyDescription(elem, "il-");

                case "FiltroPodotactil":
                    return FilterByAssemblyDescription(elem, "piso de baldosa podotactil");

                case "FiltroTableros":
                    return FilterByAssemblyDescription(elem, "te-");

                // === FILTROS MEP SITIO ===
                case "FiltroCajaDesague":
                    return FilterByAssemblyDescriptionMultiple(elem, KeywordsDesague);

                case "FiltroCajaPluvial":
                    return FilterByAssemblyDescriptionMultiple(elem, KeywordsPluvial);

                case "FiltroBuzon":
                    return FilterByAssemblyDescriptionMultiple(elem, KeywordsBuzon);

                // === FILTROS ESTRUCTURAS ===
                case "FiltroPartVereda":
                    return FilterPartVereda(elem);

                default:
                    return true;
            }
        }

        // ========== FILTROS DE PISOS POR MATERIALES ==========
        private bool FilterFloorByMaterial(Element elem, string[] keywords)
        {
            if (elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Floors)
                return false;

            var materialNames = _paramService.GetFloorMaterialNames(elem);

            foreach (var matName in materialNames)
            {
                foreach (var keyword in keywords)
                {
                    if (matName.ToLower().Contains(keyword.ToLower()))
                        return true;
                }
            }

            return false;
        }

        // ========== FILTRO PISO SITIO CEMENTO ==========
        private bool FilterPisoSitioCemento(Element elem)
        {
            if (elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Floors)
                return false;

            // Busca en Assembly Description del TIPO
            Element tipo = _doc.GetElement(elem.GetTypeId());
            if (tipo == null)
                return false;

            Parameter p_assembly = tipo.LookupParameter("Assembly Description");
            if (p_assembly == null)
                return false;

            string desc = (p_assembly.AsString() ?? "").ToLower();

            foreach (var keyword in KeywordsSitioCemento)
            {
                if (desc.Contains(keyword.ToLower()))
                    return true;
            }

            return false;
        }

        // ========== FILTRO BLOQUETA ==========
        private bool FilterBloqueta(Element elem, double anchoTarget)
        {
            // Solo Walls y StructuralFraming
            if (elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Walls &&
                elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_StructuralFraming)
                return false;

            // Verificar Assembly Description del TIPO
            Element tipo = _doc.GetElement(elem.GetTypeId());
            if (tipo == null)
                return false;

            Parameter p_assembly = tipo.LookupParameter("Assembly Description");
            if (p_assembly == null)
                return false;

            string desc = (p_assembly.AsString() ?? "").ToLower();

            // Debe contener "bloqueta" o "dintel"
            if (!desc.Contains("bloqueta") && !desc.Contains("dintel"))
                return false;

            // Verificar ancho
            double? ancho = _paramService.GetElementWidth(elem);
            if (!ancho.HasValue)
                return false;

            return Math.Abs(ancho.Value - anchoTarget) <= _tolerancia;
        }

        // ========== FILTROS POR ASSEMBLY DESCRIPTION ==========
        private bool FilterByAssemblyDescription(Element elem, string keyword)
        {
            Element tipo = _doc.GetElement(elem.GetTypeId());
            if (tipo == null)
                return false;

            Parameter p_assembly = tipo.LookupParameter("Assembly Description");
            if (p_assembly == null)
                return false;

            string desc = (p_assembly.AsString() ?? "").ToLower();
            return desc.Contains(keyword.ToLower());
        }

        private bool FilterByAssemblyDescriptionMultiple(Element elem, string[] keywords)
        {
            Element tipo = _doc.GetElement(elem.GetTypeId());
            if (tipo == null)
                return false;

            Parameter p_assembly = tipo.LookupParameter("Assembly Description");
            if (p_assembly == null)
                return false;

            string desc = (p_assembly.AsString() ?? "").ToLower();

            foreach (var keyword in keywords)
            {
                if (desc.Contains(keyword.ToLower()))
                    return true;
            }

            return false;
        }

        // ========== FILTRO PART VEREDA (ESTRUCTURAS) ==========
        private bool FilterPartVereda(Element elem)
        {
            if (elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Parts)
                return false;

            Parameter p_desc = elem.LookupParameter("DESCRIPTION PARTS");
            if (p_desc == null)
                return false;

            string desc = (p_desc.AsString() ?? "").ToLower();

            foreach (var keyword in KeywordsPartVereda)
            {
                if (desc.Contains(keyword.ToLower()))
                    return true;
            }

            return false;
        }
    }
}
