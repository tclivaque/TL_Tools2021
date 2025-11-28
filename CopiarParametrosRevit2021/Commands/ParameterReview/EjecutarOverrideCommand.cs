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
        // 1. Leer parámetro configurado
        string nombreParametro = LeerParametroGuardado();
        if (string.IsNullOrWhiteSpace(nombreParametro))
        {
            throw new Exception("No se encontró un parámetro configurado.");
        }

        // 2. Obtener lista blanca de categorías desde Google Sheets
        List<string> categoriasMetradas = new List<string>();
        try
        {
            // Usar explícitamente el helper general del proyecto, no el del Lookahead
            var sheetsService = new CopiarParametrosRevit2021.Helpers.GoogleSheetsService();
            categoriasMetradas = sheetsService.ObtenerCategoriasDesdeSheet(SPREADSHEET_ID, SHEET_NAME)
                                              .Select(c => c.ToUpper())
                                              .ToList();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error conectando con Google Sheets: {ex.Message}");
        }

        if (categoriasMetradas.Count == 0)
        {
            throw new Exception($"No se encontraron categorías en la hoja '{SHEET_NAME}' bajo la clave 'CATEGORIAS'.");
        }

        // 3. Recolectar elementos
        FilteredElementCollector collector = new FilteredElementCollector(doc, vistaActiva.Id)
            .WhereElementIsNotElementType();

        Dictionary<string, List<ElementId>> elementosPorValor = new Dictionary<string, List<ElementId>>();
        List<ElementId> elementosSinValor = new List<ElementId>();
        List<ElementId> elementosNoMetrados = new List<ElementId>();
        Dictionary<string, Color> coloresPorValor = new Dictionary<string, Color>();

        // 4. Clasificación de elementos
        foreach (Element elem in collector)
        {
            if (elem.Category == null) continue;
            if (elem.Category.CategoryType != CategoryType.Model) continue;

            string catName = elem.Category.Name;
            if (catName == "Cameras" || catName == "Piping Systems") continue;

            if (elem.IsHidden(vistaActiva)) continue;

            string catBuiltIn = ((BuiltInCategory)elem.Category.Id.IntegerValue).ToString().ToUpper();
            bool esCategoriaMetrada = categoriasMetradas.Contains(catBuiltIn);

            if (!esCategoriaMetrada)
            {
                elementosNoMetrados.Add(elem.Id);
                continue;
            }

            string valorParametro = ObtenerValorParametro(elem, doc, nombreParametro);

            if (!string.IsNullOrWhiteSpace(valorParametro))
            {
                if (!elementosPorValor.ContainsKey(valorParametro))
                {
                    elementosPorValor[valorParametro] = new List<ElementId>();
                }
                elementosPorValor[valorParametro].Add(elem.Id);
            }
            else
            {
                elementosSinValor.Add(elem.Id);
            }
        }

        // 5. Aplicar Overrides
        using (Transaction trans = new Transaction(doc, "Aplicar Colores Auditoría"))
        {
            trans.Start();

            // A) GRUPO A (Valores encontrados)
            var valoresOrdenados = elementosPorValor.Keys.OrderBy(v => v).ToList();
            for (int i = 0; i < valoresOrdenados.Count; i++)
            {
                string valor = valoresOrdenados[i];
                List<ElementId> ids = elementosPorValor[valor];
                Color color = GenerarColorPorIndice(i);

                coloresPorValor[valor] = color;
                AplicarOverrideColor(vistaActiva, ids, color, doc);
            }

            // B) GRUPO B (Sin valor -> Rojo)
            if (elementosSinValor.Count > 0)
            {
                Color colorRojo = new Color(255, 0, 0);
                AplicarOverrideColor(vistaActiva, elementosSinValor, colorRojo, doc);
            }

            // C) GRUPO C (No metrada -> Negro)
            if (elementosNoMetrados.Count > 0)
            {
                Color colorNegro = new Color(0, 0, 0);
                AplicarOverrideColor(vistaActiva, elementosNoMetrados, colorNegro, doc);

                string keyNegro = "CATEGORÍA NO METRADA";
                elementosPorValor[keyNegro] = elementosNoMetrados;
                coloresPorValor[keyNegro] = colorNegro;
            }

            trans.Commit();
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