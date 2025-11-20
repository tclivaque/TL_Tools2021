using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

public class SDIProcessor
{
    // Categorías a procesar
    private static readonly BuiltInCategory[] Categorias = new[]
    {
        // ARQUITECTURA Y ESTRUCTURAS
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_Roofs,
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_Ceilings,
        BuiltInCategory.OST_Stairs,
        BuiltInCategory.OST_StairsRailing,
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_Casework,
        
        // SANITARIAS
        BuiltInCategory.OST_PlumbingFixtures,
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_PipeCurves,
        BuiltInCategory.OST_PipeAccessory,
        
        // ELÉCTRICAS Y COMUNICACIONES
        BuiltInCategory.OST_ConduitFitting,
        BuiltInCategory.OST_Conduit,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_FlexPipeCurves,
        BuiltInCategory.OST_DataDevices,
        BuiltInCategory.OST_SecurityDevices,
        BuiltInCategory.OST_FireAlarmDevices,
        BuiltInCategory.OST_CommunicationDevices,
        BuiltInCategory.OST_NurseCallDevices,
        BuiltInCategory.OST_LightingFixtures,
        
        // MECÁNICAS
        BuiltInCategory.OST_DuctCurves,
        BuiltInCategory.OST_DuctFitting
    };

    public class ResultadosProcesamiento
    {
        public List<ElementoAntesDespues> ElementosSDI { get; set; } = new List<ElementoAntesDespues>();
        public List<ElementoAntesDespues> ElementosINCIDENCIA { get; set; } = new List<ElementoAntesDespues>();
        public List<Element> ElementosNoModificados { get; set; } = new List<Element>();
        public List<ElementoAntesDespues> ElementosYaCorrectos_SDI { get; set; } = new List<ElementoAntesDespues>();
        public List<ElementoAntesDespues> ElementosYaCorrectos_INC { get; set; } = new List<ElementoAntesDespues>();
        public List<Element> ElementosSDI_NoEncontrados { get; set; } = new List<Element>();
        public List<Element> ElementosConflictoSDI_INC { get; set; } = new List<Element>();
        public List<Element> ElementosINC_FormatoInvalido { get; set; } = new List<Element>();
        public List<string> ElementosConError { get; set; } = new List<string>();

        public string GenerarResumen()
        {
            int total = ElementosSDI.Count + ElementosINCIDENCIA.Count + ElementosNoModificados.Count +
                       ElementosYaCorrectos_SDI.Count + ElementosYaCorrectos_INC.Count +
                       ElementosSDI_NoEncontrados.Count + ElementosConflictoSDI_INC.Count +
                       ElementosINC_FormatoInvalido.Count;

            return $"=== RESUMEN DE ESTANDARIZACIÓN ===\n" +
                   $"Total elementos procesados: {total}\n" +
                   $"SDI estandarizados: {ElementosSDI.Count}\n" +
                   $"Incidencias estandarizadas: {ElementosINCIDENCIA.Count}\n" +
                   $"SDI ya correctos (actualizados): {ElementosYaCorrectos_SDI.Count}\n" +
                   $"Incidencias ya correctas (actualizadas): {ElementosYaCorrectos_INC.Count}\n" +
                   $"SDI no encontrados en diccionario: {ElementosSDI_NoEncontrados.Count}\n" +
                   $"Conflictos SDI+INC: {ElementosConflictoSDI_INC.Count}\n" +
                   $"Incidencias formato inválido: {ElementosINC_FormatoInvalido.Count}\n" +
                   $"No modificados: {ElementosNoModificados.Count}\n" +
                   $"Errores: {ElementosConError.Count}";
        }
    }

    public class ElementoAntesDespues
    {
        public Element Elemento { get; set; }
        public string NumeroSDI_Antes { get; set; }
        public string NumeroSDI_Despues { get; set; }
        public string TipoMod_Antes { get; set; }
        public string TipoMod_Despues { get; set; }
        public bool Modificado_Antes { get; set; }
        public bool Modificado_Despues { get; set; }
    }

    public static ResultadosProcesamiento ProcesarElementos(Document doc, List<string> listaSdiIds, List<string> listaSdiTipos)
    {
        var resultados = new ResultadosProcesamiento();

        // Crear diccionario SDI
        Dictionary<string, string> diccionarioSDI = CrearDiccionarioSDI(listaSdiIds, listaSdiTipos);

        // Obtener elementos por categorías
        List<Element> elementos = ObtenerElementosPorCategorias(doc);

        // Procesar cada elemento
        foreach (Element elemento in elementos)
        {
            try
            {
                ProcesarElemento(elemento, diccionarioSDI, resultados);
            }
            catch (Exception ex)
            {
                resultados.ElementosConError.Add($"{elemento.Id.IntegerValue}: {ex.Message}");
            }
        }

        return resultados;
    }

    private static void ProcesarElemento(Element elemento, Dictionary<string, string> diccionarioSDI, ResultadosProcesamiento resultados)
    {
        Parameter paramNumeroSDI = elemento.LookupParameter("NUMERO DE SDI");
        Parameter paramTipoMod = elemento.LookupParameter("TIPO DE MODIFICACION");
        Parameter paramModificado = elemento.LookupParameter("MODIFICADO");

        // Guardar valores ANTES
        string numeroSDI_Antes = paramNumeroSDI?.AsString() ?? "";
        string tipoMod_Antes = paramTipoMod?.AsString() ?? "";
        int modificado_Antes = paramModificado?.AsInteger() ?? 0;

        // Pre-verificación: detectar conflictos
        bool tieneSdiValido = TieneSDIValido(paramNumeroSDI);
        bool tieneIncValida = TieneIncidenciaValida(paramTipoMod);

        if (tieneSdiValido && tieneIncValida)
        {
            resultados.ElementosConflictoSDI_INC.Add(elemento);
            return;
        }

        // Procesar SDI
        if (tieneSdiValido && paramNumeroSDI != null && !paramNumeroSDI.IsReadOnly)
        {
            string valorSDI = paramNumeroSDI.AsString();

            if (VerificarFormatoSDICorrecto(valorSDI))
            {
                string sdiFormateado = valorSDI.Trim();

                if (diccionarioSDI.ContainsKey(sdiFormateado))
                {
                    string tipoMod = diccionarioSDI[sdiFormateado];

                    if (paramTipoMod != null && !paramTipoMod.IsReadOnly)
                        paramTipoMod.Set(tipoMod);

                    if (paramModificado != null && !paramModificado.IsReadOnly)
                        paramModificado.Set(1);

                    resultados.ElementosYaCorrectos_SDI.Add(new ElementoAntesDespues
                    {
                        Elemento = elemento,
                        NumeroSDI_Antes = numeroSDI_Antes,
                        NumeroSDI_Despues = sdiFormateado,
                        TipoMod_Antes = tipoMod_Antes,
                        TipoMod_Despues = tipoMod,
                        Modificado_Antes = modificado_Antes == 1,
                        Modificado_Despues = true
                    });
                    return;
                }
                else
                {
                    resultados.ElementosSDI_NoEncontrados.Add(elemento);
                    return;
                }
            }

            string numeroSDI = ExtraerNumeroSDI(valorSDI);
            if (numeroSDI != null)
            {
                string sdiFormateado = numeroSDI.PadLeft(3, '0');

                if (diccionarioSDI.ContainsKey(sdiFormateado))
                {
                    string tipoMod = diccionarioSDI[sdiFormateado];

                    paramNumeroSDI.Set(sdiFormateado);

                    if (paramTipoMod != null && !paramTipoMod.IsReadOnly)
                        paramTipoMod.Set(tipoMod);

                    if (paramModificado != null && !paramModificado.IsReadOnly)
                        paramModificado.Set(1);

                    resultados.ElementosSDI.Add(new ElementoAntesDespues
                    {
                        Elemento = elemento,
                        NumeroSDI_Antes = numeroSDI_Antes,
                        NumeroSDI_Despues = sdiFormateado,
                        TipoMod_Antes = tipoMod_Antes,
                        TipoMod_Despues = tipoMod,
                        Modificado_Antes = modificado_Antes == 1,
                        Modificado_Despues = true
                    });
                    return;
                }
                else
                {
                    resultados.ElementosSDI_NoEncontrados.Add(elemento);
                    return;
                }
            }
        }

        // Procesar INCIDENCIA
        if (tieneIncValida && paramTipoMod != null && !paramTipoMod.IsReadOnly)
        {
            string valorTipoMod = paramTipoMod.AsString();

            if (!string.IsNullOrEmpty(valorTipoMod))
            {
                var (incidencias, tieneInvalidos) = ExtraerIncidenciasValidas(valorTipoMod);

                if (tieneInvalidos)
                {
                    resultados.ElementosINC_FormatoInvalido.Add(elemento);
                    return;
                }

                if (incidencias.Count > 0)
                {
                    bool yaCorrectoInc = VerificarFormatoIncidenciaCorrecto(valorTipoMod);
                    string tipoModFormateado = FormatearIncidencias(incidencias);

                    // CAMBIO CRÍTICO: NUMERO DE SDI debe ser "N/A" para incidencias
                    string numeroSDI_NA = "N/A";

                    if (!yaCorrectoInc)
                        paramTipoMod.Set(tipoModFormateado);
                    else
                    {
                        // Si ya está correcto pero sin tilde, corregir la tilde
                        paramTipoMod.Set(tipoModFormateado);
                    }

                    // NUMERO DE SDI = "N/A" (no el número de incidencias)
                    if (paramNumeroSDI != null && !paramNumeroSDI.IsReadOnly)
                        paramNumeroSDI.Set(numeroSDI_NA);

                    if (paramModificado != null && !paramModificado.IsReadOnly)
                        paramModificado.Set(1);

                    if (yaCorrectoInc)
                    {
                        resultados.ElementosYaCorrectos_INC.Add(new ElementoAntesDespues
                        {
                            Elemento = elemento,
                            NumeroSDI_Antes = numeroSDI_Antes,
                            NumeroSDI_Despues = numeroSDI_NA,
                            TipoMod_Antes = tipoMod_Antes,
                            TipoMod_Despues = tipoModFormateado,
                            Modificado_Antes = modificado_Antes == 1,
                            Modificado_Despues = true
                        });
                    }
                    else
                    {
                        resultados.ElementosINCIDENCIA.Add(new ElementoAntesDespues
                        {
                            Elemento = elemento,
                            NumeroSDI_Antes = numeroSDI_Antes,
                            NumeroSDI_Despues = numeroSDI_NA,
                            TipoMod_Antes = tipoMod_Antes,
                            TipoMod_Despues = tipoModFormateado,
                            Modificado_Antes = modificado_Antes == 1,
                            Modificado_Despues = true
                        });
                    }
                    return;
                }
            }
        }

        // No es SDI ni INCIDENCIA
        if (paramModificado != null && !paramModificado.IsReadOnly)
            paramModificado.Set(0);

        resultados.ElementosNoModificados.Add(elemento);
    }

    // ============= FUNCIONES AUXILIARES =============

    private static Dictionary<string, string> CrearDiccionarioSDI(List<string> listaIds, List<string> listaTipos)
    {
        var diccionario = new Dictionary<string, string>();

        int count = Math.Min(listaIds.Count, listaTipos.Count);

        for (int i = 0; i < count; i++)
        {
            string sdi = ExtraerSDIDeEntrada(listaIds[i]);
            string tipo = ExtraerTipoModificacion(listaTipos[i]);

            if (sdi != null && tipo != null)
            {
                diccionario[sdi] = tipo;
            }
        }

        return diccionario;
    }

    private static string ExtraerSDIDeEntrada(string valorEntrada)
    {
        if (string.IsNullOrEmpty(valorEntrada))
            return null;

        Match match = Regex.Match(valorEntrada, @"\b\d{3}\b");
        return match.Success ? match.Value : null;
    }

    private static string ExtraerTipoModificacion(string valorEntrada)
    {
        if (string.IsNullOrEmpty(valorEntrada))
            return null;

        string valorUpper = valorEntrada.ToUpper();

        if (valorUpper.Contains("NO NORMATIVO"))
            return "NO NORMATIVO";
        else if (valorUpper.Contains("NORMATIVO"))
            return "NORMATIVO";
        else if (valorUpper.Contains("PEIP"))
            return "SOLICITADO POR PEIP";

        return null;
    }

    private static bool VerificarFormatoSDICorrecto(string valor)
    {
        if (string.IsNullOrEmpty(valor))
            return false;

        string valorLimpio = valor.Trim();
        return Regex.IsMatch(valorLimpio, @"^\d{3}$");
    }

    private static bool VerificarFormatoIncidenciaCorrecto(string valor)
    {
        if (string.IsNullOrEmpty(valor))
            return false;

        // Acepta con o sin tilde: ACLARACION o ACLARACIÓN
        string patron = @"^ACLARACI[OÓ]N - INC \d{4}( - \d{4})*$";
        return Regex.IsMatch(valor.Trim(), patron);
    }

    private static string ExtraerNumeroSDI(string valor)
    {
        if (string.IsNullOrEmpty(valor))
            return null;

        MatchCollection matches = Regex.Matches(valor, @"\d+");

        foreach (Match match in matches)
        {
            string num = match.Value;
            if (num.Length >= 1 && num.Length <= 3)
                return num;
        }

        return null;
    }

    private static (List<string> validos, bool tieneInvalidos) ExtraerIncidenciasValidas(string valor)
    {
        if (string.IsNullOrEmpty(valor))
            return (new List<string>(), false);

        MatchCollection todosNumeros = Regex.Matches(valor, @"\d+");

        List<string> validos = new List<string>();
        bool invalidos = false;

        foreach (Match match in todosNumeros)
        {
            string num = match.Value;
            if (num.Length == 4)
                validos.Add(num);
            else
                invalidos = true;
        }

        return (validos, invalidos);
    }

    private static string FormatearIncidencias(List<string> listaIncidencias)
    {
        if (listaIncidencias == null || listaIncidencias.Count == 0)
            return null;

        // CAMBIO: Usar ACLARACIÓN con tilde
        return "ACLARACIÓN - INC " + string.Join(" - ", listaIncidencias);
    }

    private static bool TieneSDIValido(Parameter param)
    {
        if (param == null || param.IsReadOnly)
            return false;

        string valor = param.AsString();
        if (string.IsNullOrEmpty(valor))
            return false;

        // No considerar "N/A" como SDI válido
        if (valor.Trim().Equals("N/A", StringComparison.OrdinalIgnoreCase))
            return false;

        return VerificarFormatoSDICorrecto(valor) || ExtraerNumeroSDI(valor) != null;
    }

    private static bool TieneIncidenciaValida(Parameter param)
    {
        if (param == null || param.IsReadOnly)
            return false;

        string valor = param.AsString();
        if (string.IsNullOrEmpty(valor))
            return false;

        var (incidencias, tieneInvalidos) = ExtraerIncidenciasValidas(valor);
        return incidencias.Count > 0 && !tieneInvalidos;
    }

    private static List<Element> ObtenerElementosPorCategorias(Document doc)
    {
        List<Element> elementosTotales = new List<Element>();

        foreach (BuiltInCategory cat in Categorias)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfCategory(cat)
                .WhereElementIsNotElementType();

            elementosTotales.AddRange(collector.ToElements());
        }

        return elementosTotales;
    }
}