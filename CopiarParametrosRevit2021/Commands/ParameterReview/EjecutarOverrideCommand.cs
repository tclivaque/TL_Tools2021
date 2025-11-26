using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
// AGREGADO: Referencia al namespace donde pusimos el servicio
using CopiarParametrosRevit2021.Helpers;

[Transaction(TransactionMode.Manual)]
public class EjecutarOverrideCommand : IExternalCommand
{
    // Configuración de Google Sheets
    private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";
    private const string SHEET_NAME = "ENTRADAS_PLUGIN_6.0_AUDITORÍA";

    // ID para debug específico
    private const string DEBUG_ELEMENT_ID = "282354";

    private static string ConfigFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "TL_Tools2021",
        "parametro_config.txt");

    // Variables para debug
    private static StringBuilder _debugLog = new StringBuilder();
    private static string _debugFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
        $"AplicarColores_Debug_{DEBUG_ELEMENT_ID}.txt");

    private static void Log(string message)
    {
        _debugLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {message}");
    }

    private static void SaveDebugLog()
    {
        try
        {
            File.WriteAllText(_debugFilePath, _debugLog.ToString());
        }
        catch { /* Ignorar errores de escritura */ }
    }

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
            Log($"\n[EXCEPCIÓN]: {ex.Message}");
            Log(ex.StackTrace);
            TaskDialog.Show("Error", $"Error al aplicar override: {ex.Message}");
        }
        finally
        {
            SaveDebugLog();
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
            Log($"\n[EXCEPCIÓN]: {ex.Message}");
            Log(ex.StackTrace);
            message = $"Error: {ex.Message}";
            return Result.Failed;
        }
        finally
        {
            SaveDebugLog();
        }
    }

    private static void ExecuteLogic(UIApplication uiApp, Document doc, View vistaActiva)
    {
        // Inicializar debug
        _debugLog.Clear();
        Log("==================================================");
        Log($" DEBUG APLICAR COLORES - {DateTime.Now}");
        Log($" ID DE ELEMENTO A RASTREAR: {DEBUG_ELEMENT_ID}");
        Log($" Spreadsheet: {SPREADSHEET_ID}");
        Log($" Hoja: {SHEET_NAME}");
        Log($" Vista: {vistaActiva.Name}");
        Log("==================================================\n");

        // 1. Leer parámetro configurado
        Log("[PASO 1] LEYENDO PARÁMETRO CONFIGURADO...");
        string nombreParametro = LeerParametroGuardado();
        if (string.IsNullOrWhiteSpace(nombreParametro))
        {
            Log("[ERROR] No se encontró parámetro configurado.");
            throw new Exception("No se encontró un parámetro configurado.");
        }
        Log($"-> Parámetro configurado: '{nombreParametro}'\n");

        // 2. Obtener lista blanca de categorías desde Google Sheets
        Log("[PASO 2] CONECTANDO A GOOGLE SHEETS...");
        List<string> categoriasMetradas = new List<string>();
        try
        {
            var sheetsService = new GoogleSheetsService();
            categoriasMetradas = sheetsService.ObtenerCategoriasDesdeSheet(SPREADSHEET_ID, SHEET_NAME)
                                              .Select(c => c.ToUpper())
                                              .ToList();
            Log($"-> Conexión exitosa. Categorías leídas: {categoriasMetradas.Count}");
            Log("-> Categorías metradas:");
            foreach (var cat in categoriasMetradas)
            {
                Log($"   - {cat}");
            }
        }
        catch (Exception ex)
        {
            Log($"[ERROR] Fallo conexión Google Sheets: {ex.Message}");
            throw new Exception($"Error conectando con Google Sheets: {ex.Message}");
        }

        if (categoriasMetradas.Count == 0)
        {
            Log($"[ERROR] No se encontraron categorías en la hoja.");
            throw new Exception($"No se encontraron categorías en la hoja '{SHEET_NAME}' bajo la clave 'CATEGORIAS'.");
        }
        Log("");

        // 3. Recolectar elementos
        Log("[PASO 3] RECOLECTANDO ELEMENTOS DE LA VISTA...");
        FilteredElementCollector collector = new FilteredElementCollector(doc, vistaActiva.Id)
            .WhereElementIsNotElementType();

        Dictionary<string, List<ElementId>> elementosPorValor = new Dictionary<string, List<ElementId>>();
        List<ElementId> elementosSinValor = new List<ElementId>();
        List<ElementId> elementosNoMetrados = new List<ElementId>();
        Dictionary<string, Color> coloresPorValor = new Dictionary<string, Color>();

        bool debugElementFound = false;
        int totalElementos = 0;

        // 4. Clasificación de elementos
        Log("[PASO 4] CLASIFICANDO ELEMENTOS...");
        foreach (Element elem in collector)
        {
            totalElementos++;
            string elemIdStr = elem.Id.IntegerValue.ToString();
            bool isDebugElement = elemIdStr == DEBUG_ELEMENT_ID;

            if (isDebugElement)
            {
                debugElementFound = true;
                Log($"\n>>> ELEMENTO {DEBUG_ELEMENT_ID} ENCONTRADO <<<");
                Log($"   -> Nombre: {elem.Name}");
            }

            if (elem.Category == null)
            {
                if (isDebugElement) Log($"   -> [DESCARTADO] No tiene categoría.");
                continue;
            }

            if (isDebugElement) Log($"   -> Categoría: {elem.Category.Name}");

            if (elem.Category.CategoryType != CategoryType.Model)
            {
                if (isDebugElement) Log($"   -> [DESCARTADO] No es CategoryType.Model (es {elem.Category.CategoryType})");
                continue;
            }

            string catName = elem.Category.Name;
            if (catName == "Cameras" || catName == "Piping Systems")
            {
                if (isDebugElement) Log($"   -> [DESCARTADO] Categoría excluida: {catName}");
                continue;
            }

            if (elem.IsHidden(vistaActiva))
            {
                if (isDebugElement) Log($"   -> [DESCARTADO] Elemento está oculto en la vista.");
                continue;
            }

            string catNameUpper = catName.ToUpper();
            bool esCategoriaMetrada = categoriasMetradas.Contains(catNameUpper);

            if (isDebugElement)
            {
                Log($"   -> Categoría en mayúsculas: '{catNameUpper}'");
                Log($"   -> ¿Está en lista de categorías metradas? {esCategoriaMetrada}");
            }

            if (!esCategoriaMetrada)
            {
                elementosNoMetrados.Add(elem.Id);
                if (isDebugElement) Log($"   -> [CLASIFICADO] GRUPO C - No metrada (se pintará NEGRO)");
                continue;
            }

            string valorParametro = ObtenerValorParametro(elem, doc, nombreParametro);

            if (isDebugElement)
            {
                Log($"   -> Valor del parámetro '{nombreParametro}': '{valorParametro ?? "(vacío)"}'");
            }

            if (!string.IsNullOrWhiteSpace(valorParametro))
            {
                if (!elementosPorValor.ContainsKey(valorParametro))
                {
                    elementosPorValor[valorParametro] = new List<ElementId>();
                }
                elementosPorValor[valorParametro].Add(elem.Id);
                if (isDebugElement) Log($"   -> [CLASIFICADO] GRUPO A - Con valor: '{valorParametro}'");
            }
            else
            {
                elementosSinValor.Add(elem.Id);
                if (isDebugElement) Log($"   -> [CLASIFICADO] GRUPO B - Sin valor (se pintará ROJO)");
            }
        }

        Log($"\n-> Total elementos procesados: {totalElementos}");
        Log($"-> Elemento {DEBUG_ELEMENT_ID} encontrado: {(debugElementFound ? "SÍ" : "NO")}");
        if (!debugElementFound)
        {
            Log($"   [ADVERTENCIA] El elemento {DEBUG_ELEMENT_ID} NO fue encontrado en la vista actual.");
        }
        Log($"-> GRUPO A (con valor): {elementosPorValor.Sum(kvp => kvp.Value.Count)} elementos");
        Log($"-> GRUPO B (sin valor): {elementosSinValor.Count} elementos");
        Log($"-> GRUPO C (no metrada): {elementosNoMetrados.Count} elementos\n");

        // 5. Aplicar Overrides
        Log("[PASO 5] APLICANDO COLORES...");
        using (Transaction trans = new Transaction(doc, "Aplicar Colores Auditoría"))
        {
            trans.Start();

            // A) GRUPO A (Valores encontrados)
            Log("-> Aplicando colores a GRUPO A (con valores):");
            var valoresOrdenados = elementosPorValor.Keys.OrderBy(v => v).ToList();
            for (int i = 0; i < valoresOrdenados.Count; i++)
            {
                string valor = valoresOrdenados[i];
                List<ElementId> ids = elementosPorValor[valor];
                Color color = GenerarColorPorIndice(i);

                coloresPorValor[valor] = color;

                bool containsDebugElement = ids.Any(id => id.IntegerValue.ToString() == DEBUG_ELEMENT_ID);
                if (containsDebugElement)
                {
                    Log($"   >>> ELEMENTO {DEBUG_ELEMENT_ID} - Aplicando color RGB({color.Red},{color.Green},{color.Blue}) para valor '{valor}'");
                }

                AplicarOverrideColor(vistaActiva, ids, color, doc);
            }

            // B) GRUPO B (Sin valor -> Rojo)
            if (elementosSinValor.Count > 0)
            {
                Log($"-> Aplicando ROJO a GRUPO B (sin valor): {elementosSinValor.Count} elementos");
                Color colorRojo = new Color(255, 0, 0);

                bool containsDebugElement = elementosSinValor.Any(id => id.IntegerValue.ToString() == DEBUG_ELEMENT_ID);
                if (containsDebugElement)
                {
                    Log($"   >>> ELEMENTO {DEBUG_ELEMENT_ID} - Aplicando color ROJO");
                }

                AplicarOverrideColor(vistaActiva, elementosSinValor, colorRojo, doc);
            }

            // C) GRUPO C (No metrada -> Negro)
            if (elementosNoMetrados.Count > 0)
            {
                Log($"-> Aplicando NEGRO a GRUPO C (no metrada): {elementosNoMetrados.Count} elementos");
                Color colorNegro = new Color(0, 0, 0);

                bool containsDebugElement = elementosNoMetrados.Any(id => id.IntegerValue.ToString() == DEBUG_ELEMENT_ID);
                if (containsDebugElement)
                {
                    Log($"   >>> ELEMENTO {DEBUG_ELEMENT_ID} - Aplicando color NEGRO");
                }

                AplicarOverrideColor(vistaActiva, elementosNoMetrados, colorNegro, doc);

                string keyNegro = "CATEGORÍA NO METRADA";
                elementosPorValor[keyNegro] = elementosNoMetrados;
                coloresPorValor[keyNegro] = colorNegro;
            }

            trans.Commit();
            Log("-> Transacción completada exitosamente.");
        }

        Log("\n[PROCESO COMPLETADO]");
        Log($"Archivo de debug guardado en: {_debugFilePath}\n");

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