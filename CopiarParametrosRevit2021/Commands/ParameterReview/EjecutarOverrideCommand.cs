using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class EjecutarOverrideCommand : IExternalCommand
{
    private static string ConfigFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "TL_Tools2021",
        "parametro_config.txt");

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        View vistaActiva = doc.ActiveView;

        try
        {
            // Leer parámetro configurado
            string nombreParametro = LeerParametroGuardado();

            if (string.IsNullOrWhiteSpace(nombreParametro))
            {
                return Result.Failed;
            }

            // Obtener todos los elementos visibles en la vista activa
            FilteredElementCollector collector = new FilteredElementCollector(doc, vistaActiva.Id)
                .WhereElementIsNotElementType();

            // Filtrar elementos según criterios
            List<Element> elementosFiltrados = new List<Element>();

            foreach (Element elem in collector)
            {
                try
                {
                    // 1. Verificar que tenga categoría
                    if (elem.Category == null)
                        continue;

                    // 2. Excluir categorías específicas
                    string categoryName = elem.Category.Name;
                    if (categoryName == "Cameras" || categoryName == "Piping Systems")
                        continue;

                    // 3. Solo categorías 3D (excluir anotaciones)
                    if (elem.Category.CategoryType != CategoryType.Model)
                        continue;

                    // 4. Verificar visibilidad
                    if (!elem.IsHidden(vistaActiva) && vistaActiva.CanCategoryBeHidden(elem.Category.Id))
                    {
                        elementosFiltrados.Add(elem);
                    }
                    else if (!vistaActiva.CanCategoryBeHidden(elem.Category.Id))
                    {
                        // Elementos de categorías que no se pueden ocultar (siempre visibles)
                        elementosFiltrados.Add(elem);
                    }
                }
                catch
                {
                    // Si hay error verificando, omitir el elemento
                    continue;
                }
            }

            Dictionary<string, List<ElementId>> elementosPorValor = new Dictionary<string, List<ElementId>>();
            List<ElementId> elementosSinValor = new List<ElementId>();
            Dictionary<string, Color> coloresPorValor = new Dictionary<string, Color>();

            foreach (Element elem in elementosFiltrados)
            {
                bool tieneValorValido = false;

                // Intentar obtener el parámetro (primero de ejemplar, luego de tipo)
                Parameter param = elem.LookupParameter(nombreParametro);

                if (param == null)
                {
                    // Intentar con el tipo del elemento
                    ElementId typeId = elem.GetTypeId();
                    if (typeId != ElementId.InvalidElementId)
                    {
                        ElementType elemType = doc.GetElement(typeId) as ElementType;
                        if (elemType != null)
                        {
                            param = elemType.LookupParameter(nombreParametro);
                        }
                    }
                }

                // Si encontró el parámetro, verificar si tiene valor válido
                if (param != null)
                {
                    string valorParametro = ObtenerValorComoString(param);

                    if (!string.IsNullOrWhiteSpace(valorParametro))
                    {
                        // Tiene valor válido
                        if (!elementosPorValor.ContainsKey(valorParametro))
                        {
                            elementosPorValor[valorParametro] = new List<ElementId>();
                        }
                        elementosPorValor[valorParametro].Add(elem.Id);
                        tieneValorValido = true;
                    }
                }

                // Si no tiene valor válido, agregarlo a elementos sin valor
                if (!tieneValorValido)
                {
                    elementosSinValor.Add(elem.Id);
                }
            }

            using (Transaction trans = new Transaction(doc, "Aplicar Override por Parámetro"))
            {
                trans.Start();

                // Ordenar valores alfabéticamente y asignar colores secuencialmente
                var valoresOrdenados = elementosPorValor.Keys.OrderBy(v => v).ToList();

                for (int i = 0; i < valoresOrdenados.Count; i++)
                {
                    string valor = valoresOrdenados[i];
                    List<ElementId> elementIds = elementosPorValor[valor];

                    // Generar color según posición alfabética
                    Color colorFondo = GenerarColorPorIndice(i);

                    // Guardar color para la leyenda
                    coloresPorValor[valor] = colorFondo;

                    // Calcular color para líneas (más oscuro - cada componente RGB / 2 redondeado hacia arriba)
                    Color colorLineas = new Color(
                        (byte)((colorFondo.Red + 1) * 9 / 10),
                        (byte)((colorFondo.Green + 1) * 9 / 10),
                        (byte)((colorFondo.Blue + 1) * 9 / 10)
                    );

                    // Crear configuración de override
                    OverrideGraphicSettings overrideSettings = new OverrideGraphicSettings();

                    // Obtener patrones
                    ElementId solidFillPatternId = GetSolidFillPatternId(doc);
                    ElementId solidLinePatternId = GetSolidLinePatternId(doc);

                    // Configurar projection lines - color más oscuro que el fondo, weight 16
                    overrideSettings.SetProjectionLineColor(colorLineas);
                    overrideSettings.SetProjectionLineWeight(16);
                    if (solidLinePatternId != null && solidLinePatternId != ElementId.InvalidElementId)
                    {
                        overrideSettings.SetProjectionLinePatternId(solidLinePatternId);
                    }

                    // Configurar cut lines - mismo color y patrón que projection lines
                    overrideSettings.SetCutLineColor(colorLineas);
                    overrideSettings.SetCutLineWeight(16);
                    if (solidLinePatternId != null && solidLinePatternId != ElementId.InvalidElementId)
                    {
                        overrideSettings.SetCutLinePatternId(solidLinePatternId);
                    }

                    // Configurar surface foreground pattern - color generado original
                    if (solidFillPatternId != null && solidFillPatternId != ElementId.InvalidElementId)
                    {
                        overrideSettings.SetSurfaceForegroundPatternId(solidFillPatternId);
                        overrideSettings.SetSurfaceForegroundPatternColor(colorFondo);
                        overrideSettings.SetSurfaceForegroundPatternVisible(true);
                    }

                    // Configurar cut patterns - mismo patrón y color que surface
                    if (solidFillPatternId != null && solidFillPatternId != ElementId.InvalidElementId)
                    {
                        overrideSettings.SetCutForegroundPatternId(solidFillPatternId);
                        overrideSettings.SetCutForegroundPatternColor(colorFondo);
                        overrideSettings.SetCutForegroundPatternVisible(true);
                    }

                    // Aplicar override a cada elemento
                    foreach (ElementId elemId in elementIds)
                    {
                        vistaActiva.SetElementOverrides(elemId, overrideSettings);
                    }
                }

                // Aplicar color rojo a elementos sin valor
                if (elementosSinValor.Count > 0)
                {
                    Color colorRojoFondo = new Color(255, 0, 0);
                    Color colorRojoLineas = new Color(127, 0, 0); // 255/2 redondeado

                    OverrideGraphicSettings overrideRojo = new OverrideGraphicSettings();

                    // Obtener patrones
                    ElementId solidFillPatternId = GetSolidFillPatternId(doc);
                    ElementId solidLinePatternId = GetSolidLinePatternId(doc);

                    // Configurar projection lines
                    overrideRojo.SetProjectionLineColor(colorRojoLineas);
                    overrideRojo.SetProjectionLineWeight(16);
                    if (solidLinePatternId != null && solidLinePatternId != ElementId.InvalidElementId)
                    {
                        overrideRojo.SetProjectionLinePatternId(solidLinePatternId);
                    }

                    // Configurar cut lines
                    overrideRojo.SetCutLineColor(colorRojoLineas);
                    overrideRojo.SetCutLineWeight(16);
                    if (solidLinePatternId != null && solidLinePatternId != ElementId.InvalidElementId)
                    {
                        overrideRojo.SetCutLinePatternId(solidLinePatternId);
                    }

                    // Configurar surface foreground pattern
                    if (solidFillPatternId != null && solidFillPatternId != ElementId.InvalidElementId)
                    {
                        overrideRojo.SetSurfaceForegroundPatternId(solidFillPatternId);
                        overrideRojo.SetSurfaceForegroundPatternColor(colorRojoFondo);
                        overrideRojo.SetSurfaceForegroundPatternVisible(true);
                    }

                    // Configurar cut patterns
                    if (solidFillPatternId != null && solidFillPatternId != ElementId.InvalidElementId)
                    {
                        overrideRojo.SetCutForegroundPatternId(solidFillPatternId);
                        overrideRojo.SetCutForegroundPatternColor(colorRojoFondo);
                        overrideRojo.SetCutForegroundPatternVisible(true);
                    }

                    // Aplicar a todos los elementos sin valor
                    foreach (ElementId elemId in elementosSinValor)
                    {
                        vistaActiva.SetElementOverrides(elemId, overrideRojo);
                    }
                }

                trans.Commit();
            }

            // Mostrar ventana de leyenda (modeless)
            VentanaLeyenda ventana = VentanaLeyenda.ObtenerInstancia();
            ventana.InicializarEventHandler(commandData.Application);
            ventana.ActualizarLeyenda(elementosPorValor, elementosSinValor, coloresPorValor, vistaActiva);

            if (!ventana.IsVisible)
            {
                ventana.Show();
            }

            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            message = $"Error: {ex.Message}";
            return Result.Failed;
        }
    }

    private string LeerParametroGuardado()
    {
        try
        {
            if (File.Exists(ConfigFilePath))
            {
                return File.ReadAllText(ConfigFilePath).Trim();
            }
        }
        catch { }
        return string.Empty;
    }

    private string ObtenerValorComoString(Parameter param)
    {
        if (!param.HasValue)
            return string.Empty;

        switch (param.StorageType)
        {
            case StorageType.String:
                string strVal = param.AsString();
                return strVal ?? string.Empty;
            case StorageType.Integer:
                return param.AsInteger().ToString();
            case StorageType.Double:
                return param.AsDouble().ToString();
            case StorageType.ElementId:
                ElementId elemId = param.AsElementId();
                return (elemId != null && elemId != ElementId.InvalidElementId) ? elemId.IntegerValue.ToString() : string.Empty;
            default:
                return string.Empty;
        }
    }

    private Color GenerarColorPorIndice(int indice)
    {
        // Paleta de 80 colores de BIMcollab
        Color[] paletaBIMcollab = new Color[]
        {
            new Color(255, 255, 0),    // 1
            new Color(28, 230, 255),   // 2
            new Color(255, 52, 255),   // 3
            new Color(255, 74, 70),    // 4
            new Color(0, 137, 65),     // 5
            new Color(0, 111, 166),    // 6
            new Color(163, 0, 89),     // 7
            new Color(255, 219, 229),  // 8
            new Color(122, 73, 0),     // 9
            new Color(0, 0, 166),      // 10
            new Color(99, 255, 172),   // 11
            new Color(183, 151, 98),   // 12
            new Color(0, 77, 67),      // 13
            new Color(143, 176, 255),  // 14
            new Color(153, 125, 135),  // 15
            new Color(90, 0, 7),       // 16
            new Color(128, 150, 147),  // 17
            new Color(254, 255, 230),  // 18
            new Color(27, 68, 0),      // 19
            new Color(79, 198, 1),     // 20
            new Color(59, 93, 255),    // 21
            new Color(74, 59, 83),     // 22
            new Color(255, 47, 128),   // 23
            new Color(97, 97, 90),     // 24
            new Color(186, 9, 0),      // 25
            new Color(107, 121, 0),    // 26
            new Color(0, 194, 160),    // 27
            new Color(255, 170, 146),  // 28
            new Color(255, 144, 201),  // 29
            new Color(185, 3, 170),    // 30
            new Color(209, 97, 0),     // 31
            new Color(221, 239, 255),  // 32
            new Color(0, 0, 53),       // 33
            new Color(123, 79, 75),    // 34
            new Color(161, 194, 153),  // 35
            new Color(48, 0, 24),      // 36
            new Color(10, 166, 216),   // 37
            new Color(1, 51, 73),      // 38
            new Color(0, 132, 111),    // 39
            new Color(55, 33, 1),      // 40
            new Color(255, 181, 0),    // 41
            new Color(194, 255, 237),  // 42
            new Color(160, 121, 191),  // 43
            new Color(204, 7, 68),     // 44
            new Color(192, 185, 178),  // 45
            new Color(194, 255, 153),  // 46
            new Color(0, 30, 9),       // 47
            new Color(0, 72, 156),     // 48
            new Color(111, 0, 98),     // 49
            new Color(12, 189, 102),   // 50
            new Color(238, 195, 255),  // 51
            new Color(69, 109, 117),   // 52
            new Color(183, 123, 104),  // 53
            new Color(122, 135, 161),  // 54
            new Color(120, 141, 102),  // 55
            new Color(136, 85, 120),   // 56
            new Color(250, 208, 159),  // 57
            new Color(255, 138, 154),  // 58
            new Color(209, 87, 160),   // 59
            new Color(190, 196, 89),   // 60
            new Color(69, 102, 72),    // 61
            new Color(0, 134, 237),    // 62
            new Color(136, 111, 76),   // 63
            new Color(52, 54, 45),     // 64
            new Color(180, 168, 189),  // 65
            new Color(0, 166, 170),    // 66
            new Color(69, 44, 44),     // 67
            new Color(99, 99, 117),    // 68
            new Color(163, 200, 201),  // 69
            new Color(255, 145, 63),   // 70
            new Color(147, 138, 129),  // 71
            new Color(87, 83, 41),     // 72
            new Color(0, 254, 207),    // 73
            new Color(176, 91, 111),   // 74
            new Color(140, 208, 255),  // 75
            new Color(59, 151, 0),     // 76
            new Color(4, 247, 87),     // 77
            new Color(200, 161, 161),  // 78
            new Color(30, 110, 0),     // 79
            new Color(121, 0, 215)     // 80
        };

        // Usar módulo para reciclar colores si hay más de 80 valores
        int indiceColor = indice % paletaBIMcollab.Length;
        return paletaBIMcollab[indiceColor];
    }

    private ElementId GetSolidFillPatternId(Document doc)
    {
        // Obtener patrón sólido usando FilteredElementCollector
        try
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            FillPatternElement solidPattern = collector.OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(x => x.GetFillPattern().IsSolidFill);

            if (solidPattern != null)
                return solidPattern.Id;
        }
        catch { }

        return ElementId.InvalidElementId;
    }

    private ElementId GetSolidLinePatternId(Document doc)
    {
        // Usar el método estático GetSolidPatternId como en el código Python
        try
        {
            return LinePatternElement.GetSolidPatternId();
        }
        catch
        {
            // Fallback al ElementId del patrón sólido conocido
            return new ElementId(-3000010);
        }
    }
}