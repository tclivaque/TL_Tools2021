using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
// AGREGADO: Referencia al namespace donde pusimos el servicio
using CopiarParametrosRevit2021.Helpers;

[Transaction(TransactionMode.Manual)]
public class EjecutarOverrideCommand : IExternalCommand
{
    // Configuración de Google Sheets
    private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";
    private const string SHEET_NAME = "ENTRADAS_PLUGIN_6.0_AUDITORÍA";

    private static string ConfigFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "TL_Tools2021",
        "parametro_config.txt");

    // Método estático para llamar desde el Event Handler
    public static void ExecuteFromEvent(UIApplication uiApp)
    {
        UIDocument uidoc = uiApp.ActiveUIDocument;
        if (uidoc == null) return;

        Document doc = uidoc.Document;
        View vistaActiva = doc.ActiveView;

        try
        {
            ExecuteLogic(uiApp, doc, vistaActiva);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error", $"Error al aplicar override: {ex.Message}");
        }
    }

    // Método estándar de IExternalCommand
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View vistaActiva = doc.ActiveView;

        try
        {
            ExecuteLogic(commandData.Application, doc, vistaActiva);
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Error: {ex.Message}";
            return Result.Failed;
        }
    }

    private static void ExecuteLogic(UIApplication uiApp, Document doc, View vistaActiva)
    {
        // Variables para debug
        string debugFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            $"AplicarColores_Debug_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
        List<string> debugLog = new List<string>();

        debugLog.Add("==================================================");
        debugLog.Add($"DEBUG APLICAR COLORES - {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
        debugLog.Add($"Spreadsheet: {SPREADSHEET_ID}");
        debugLog.Add($"Hoja: {SHEET_NAME}");
        debugLog.Add($"Vista: {vistaActiva.Name}");
        debugLog.Add("==================================================");
        debugLog.Add("");

        // 1. Leer parámetro configurado
        debugLog.Add("[PASO 1] LEYENDO PARÁMETRO CONFIGURADO...");
        string nombreParametro = LeerParametroGuardado();
        if (string.IsNullOrWhiteSpace(nombreParametro))
        {
            throw new Exception("No se encontró un parámetro configurado.");
        }
        debugLog.Add($"-> Parámetro configurado: '{nombreParametro}'");
        debugLog.Add("");

        // 2. Obtener lista blanca de categorías desde Google Sheets
        debugLog.Add("[PASO 2] CONECTANDO A GOOGLE SHEETS...");
        List<string> categoriasMetradas = new List<string>();
        try
        {
            var sheetsService = new GoogleSheetsService();
            categoriasMetradas = sheetsService.ObtenerCategoriasDesdeSheet(SPREADSHEET_ID, SHEET_NAME)
                                              .Select(c => c.ToUpper())
                                              .ToList();
        }
        catch (Exception ex)
        {
            debugLog.Add($"-> ERROR: {ex.Message}");
            File.WriteAllLines(debugFilePath, debugLog);
            throw new Exception($"Error conectando con Google Sheets: {ex.Message}");
        }

        if (categoriasMetradas.Count == 0)
        {
            debugLog.Add("-> ERROR: No se encontraron categorías");
            File.WriteAllLines(debugFilePath, debugLog);
            throw new Exception($"No se encontraron categorías en la hoja '{SHEET_NAME}' bajo la clave 'CATEGORIAS'.");
        }

        debugLog.Add($"-> Conexión exitosa. Categorías leídas: {categoriasMetradas.Count}");
        debugLog.Add("-> Categorías metradas:");
        foreach (var cat in categoriasMetradas.OrderBy(c => c))
        {
            debugLog.Add($"  - {cat}");
        }
        debugLog.Add("");

        // 3. Recolectar elementos
        debugLog.Add("[PASO 3] RECOLECTANDO ELEMENTOS DE LA VISTA...");
        FilteredElementCollector collector = new FilteredElementCollector(doc, vistaActiva.Id)
            .WhereElementIsNotElementType();

        Dictionary<string, List<ElementId>> elementosPorValor = new Dictionary<string, List<ElementId>>();
        List<ElementId> elementosSinValor = new List<ElementId>();
        List<ElementId> elementosNoMetrados = new List<ElementId>();
        Dictionary<string, Color> coloresPorValor = new Dictionary<string, Color>();

        // 4. Clasificación de elementos
        debugLog.Add("[PASO 4] CLASIFICANDO ELEMENTOS...");
        int elementosProcesados = 0;
        foreach (Element elem in collector)
        {
            if (elem.Category == null) continue;
            if (elem.Category.CategoryType != CategoryType.Model) continue;

            string catName = elem.Category.Name;
            if (catName == "Cameras" || catName == "Piping Systems") continue;

            if (elem.IsHidden(vistaActiva)) continue;

            elementosProcesados++;

            // FIX: Obtener el BuiltInCategory para comparar con Google Sheets
            string catBuiltIn = ((BuiltInCategory)elem.Category.Id.IntegerValue).ToString().ToUpper();
            bool esCategoriaMetrada = categoriasMetradas.Contains(catBuiltIn);

            // Debug detallado para los primeros 5 elementos
            if (elementosProcesados <= 5)
            {
                debugLog.Add("");
                debugLog.Add($"ELEMENTO {elem.Id.IntegerValue}:");
                debugLog.Add($"  -> Nombre: {(elem as FamilyInstance)?.Name ?? elem.Name}");
                debugLog.Add($"  -> Categoría (nombre): {catName}");
                debugLog.Add($"  -> Categoría (BuiltIn): {catBuiltIn}");
                debugLog.Add($"  -> ¿Está en lista de categorías metradas? {esCategoriaMetrada}");
            }

            if (!esCategoriaMetrada)
            {
                if (elementosProcesados <= 5)
                {
                    debugLog.Add($"  -> [CLASIFICADO] GRUPO C - No metrada (se pintará NEGRO)");
                }
                elementosNoMetrados.Add(elem.Id);
                continue;
            }

            string valorParametro = ObtenerValorParametro(elem, doc, nombreParametro);

            if (!string.IsNullOrWhiteSpace(valorParametro))
            {
                if (elementosProcesados <= 5)
                {
                    debugLog.Add($"  -> Valor del parámetro '{nombreParametro}': {valorParametro}");
                    debugLog.Add($"  -> [CLASIFICADO] GRUPO A - Con valor (se pintará COLOR)");
                }
                if (!elementosPorValor.ContainsKey(valorParametro))
                {
                    elementosPorValor[valorParametro] = new List<ElementId>();
                }
                elementosPorValor[valorParametro].Add(elem.Id);
            }
            else
            {
                if (elementosProcesados <= 5)
                {
                    debugLog.Add($"  -> Valor del parámetro '{nombreParametro}': (vacío)");
                    debugLog.Add($"  -> [CLASIFICADO] GRUPO B - Sin valor (se pintará ROJO)");
                }
                elementosSinValor.Add(elem.Id);
            }
        }

        debugLog.Add("");
        debugLog.Add($"-> Total elementos procesados: {elementosProcesados}");
        debugLog.Add($"-> GRUPO A (con valor): {elementosPorValor.Sum(kvp => kvp.Value.Count)} elementos");
        debugLog.Add($"-> GRUPO B (sin valor): {elementosSinValor.Count} elementos");
        debugLog.Add($"-> GRUPO C (no metrada): {elementosNoMetrados.Count} elementos");
        debugLog.Add("");

        // 5. Aplicar Overrides
        debugLog.Add("[PASO 5] APLICANDO COLORES...");
        using (Transaction trans = new Transaction(doc, "Aplicar Colores Auditoría"))
        {
            trans.Start();

            // A) GRUPO A (Valores encontrados)
            var valoresOrdenados = elementosPorValor.Keys.OrderBy(v => v).ToList();
            if (valoresOrdenados.Count > 0)
            {
                debugLog.Add($"-> Aplicando colores a GRUPO A (con valores): {valoresOrdenados.Count} valores únicos");
            }
            for (int i = 0; i < valoresOrdenados.Count; i++)
            {
                string valor = valoresOrdenados[i];
                List<ElementId> ids = elementosPorValor[valor];
                Color color = GenerarColorPorIndice(i);

                coloresPorValor[valor] = color;
                AplicarOverrideColor(vistaActiva, ids, color, doc);

                if (i < 5) // Log de los primeros 5 valores
                {
                    debugLog.Add($"   {i + 1}. Valor '{valor}': {ids.Count} elementos -> RGB({color.Red},{color.Green},{color.Blue})");
                }
            }

            // B) GRUPO B (Sin valor -> Rojo)
            if (elementosSinValor.Count > 0)
            {
                debugLog.Add($"-> Aplicando ROJO a GRUPO B (sin valor): {elementosSinValor.Count} elementos");
                Color colorRojo = new Color(255, 0, 0);
                AplicarOverrideColor(vistaActiva, elementosSinValor, colorRojo, doc);
            }

            // C) GRUPO C (No metrada -> Negro)
            if (elementosNoMetrados.Count > 0)
            {
                debugLog.Add($"-> Aplicando NEGRO a GRUPO C (no metrada): {elementosNoMetrados.Count} elementos");
                Color colorNegro = new Color(0, 0, 0);
                AplicarOverrideColor(vistaActiva, elementosNoMetrados, colorNegro, doc);

                string keyNegro = "CATEGORÍA NO METRADA";
                elementosPorValor[keyNegro] = elementosNoMetrados;
                coloresPorValor[keyNegro] = colorNegro;
            }

            trans.Commit();
            debugLog.Add("-> Transacción completada exitosamente.");
        }

        debugLog.Add("[PROCESO COMPLETADO]");
        debugLog.Add($"Archivo de debug guardado en: {debugFilePath}");
        debugLog.Add("");

        // Guardar archivo de debug
        try
        {
            File.WriteAllLines(debugFilePath, debugLog);
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Error al guardar debug", $"No se pudo guardar el archivo de debug: {ex.Message}");
        }

        // 6. Mostrar Leyenda
        VentanaLeyenda ventana = VentanaLeyenda.ObtenerInstancia();
        ventana.InicializarEventHandler(uiApp);
        ventana.ActualizarLeyenda(elementosPorValor, elementosSinValor, coloresPorValor, vistaActiva);

        if (!ventana.IsVisible)
        {
            ventana.Show();
        }
    }

    private static void AplicarOverrideColor(View view, List<ElementId> ids, Color colorFondo, Document doc)
    {
        Color colorLinea = new Color(
            (byte)((colorFondo.Red + 1) * 0.5),
            (byte)((colorFondo.Green + 1) * 0.5),
            (byte)((colorFondo.Blue + 1) * 0.5)
        );

        OverrideGraphicSettings ogs = new OverrideGraphicSettings();
        ElementId solidPatternId = GetSolidFillPatternId(doc);

        if (solidPatternId != ElementId.InvalidElementId)
        {
            ogs.SetSurfaceForegroundPatternId(solidPatternId);
            ogs.SetSurfaceForegroundPatternColor(colorFondo);
            ogs.SetSurfaceForegroundPatternVisible(true);

            ogs.SetCutForegroundPatternId(solidPatternId);
            ogs.SetCutForegroundPatternColor(colorFondo);
            ogs.SetCutForegroundPatternVisible(true);
        }

        ogs.SetProjectionLineColor(colorLinea);
        ogs.SetCutLineColor(colorLinea);

        foreach (ElementId id in ids)
        {
            view.SetElementOverrides(id, ogs);
        }
    }

    private static string ObtenerValorParametro(Element elem, Document doc, string paramName)
    {
        Parameter param = elem.LookupParameter(paramName);
        if (param == null)
        {
            ElementId typeId = elem.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                ElementType type = doc.GetElement(typeId) as ElementType;
                param = type?.LookupParameter(paramName);
            }
        }

        return (param != null && param.HasValue) ? param.AsValueString() ?? param.AsString() : null;
    }

    private static string LeerParametroGuardado()
    {
        try
        {
            if (File.Exists(ConfigFilePath)) return File.ReadAllText(ConfigFilePath).Trim();
        }
        catch { }
        return string.Empty;
    }

    private static Color GenerarColorPorIndice(int indice)
    {
        Color[] paleta = new Color[]
        {
            new Color(255, 255, 0), new Color(28, 230, 255), new Color(255, 52, 255), new Color(255, 74, 70),
            new Color(0, 137, 65), new Color(0, 111, 166), new Color(163, 0, 89), new Color(0, 0, 166),
            new Color(99, 255, 172), new Color(183, 151, 98), new Color(0, 77, 67), new Color(143, 176, 255),
            new Color(255, 145, 63), new Color(200, 161, 161), new Color(121, 0, 215)
        };
        return paleta[indice % paleta.Length];
    }

    private static ElementId GetSolidFillPatternId(Document doc)
    {
        return new FilteredElementCollector(doc)
            .OfClass(typeof(FillPatternElement))
            .Cast<FillPatternElement>()
            .FirstOrDefault(x => x.GetFillPattern().IsSolidFill)?.Id ?? ElementId.InvalidElementId;
    }
}