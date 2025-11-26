using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System.Reflection;

// 1. CORRECCIÓN: Namespace añadido para que Revit encuentre la clase
namespace CopiarParametrosRevit2021.Commands.ParameterReview
{
    [Transaction(TransactionMode.Manual)]
    public class EjecutarOverrideCommand : IExternalCommand
    {
        // Ruta para el archivo de configuración local (nombre del parámetro a buscar)
        private static string ConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "TL_Tools2021",
            "parametro_config.txt");

        // Ruta para el LOG DE DEBUG en el Escritorio
        private static string DebugLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "Debug_AplicarColores.txt");

        // --- CONFIGURACIÓN GOOGLE SHEETS ---
        private const string SPREADSHEET_ID = "14bYBONt68lfM-sx6iIJxkYExXS0u7sdgijEScL3Ed3Y";
        private const string SHEET_NAME = "ENTRADAS_PLUGIN_6.0_AUDITORÍA";
        // Buscamos "CATEGORIAS" en Col A, leemos Col B.
        // Asumiremos un rango amplio para buscar
        private const string READ_RANGE = "'ENTRADAS_PLUGIN_6.0_AUDITORÍA'!A:B";

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
                TaskDialog.Show("Error", $"Error crítico: {ex.Message}\nRevisar log en Escritorio.");
                LogDebug($"ERROR CRITICO EN EVENTO: {ex.ToString()}");
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                ExecuteLogic(commandData.Application, commandData.Application.ActiveUIDocument.Document, commandData.Application.ActiveUIDocument.Document.ActiveView);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                LogDebug($"ERROR CRITICO EN EXECUTE: {ex.ToString()}");
                return Result.Failed;
            }
        }

        private static void ExecuteLogic(UIApplication uiApp, Document doc, View vistaActiva)
        {
            // 1. INICIAR LOG
            File.WriteAllText(DebugLogPath, $"--- INICIO DEBUG APLICAR COLORES: {DateTime.Now} ---\n");
            LogDebug($"Vista Activa: {vistaActiva.Name}");
            LogDebug($"Documento: {doc.Title}");

            // 2. LEER PARÁMETRO DE CONFIGURACIÓN LOCAL
            string nombreParametro = LeerParametroGuardado();
            LogDebug($"Parámetro seleccionado (Local Config): '{nombreParametro}'");

            if (string.IsNullOrWhiteSpace(nombreParametro))
            {
                LogDebug("ERROR: No se encontró nombre de parámetro configurado.");
                throw new Exception("No se encontró un parámetro configurado.");
            }

            // 3. OBTENER LISTA BLANCA DE GOOGLE SHEETS
            HashSet<string> categoriasPermitidas = ObtenerCategoriasDesdeSheets();
            LogDebug($"Total Categorías Permitidas (desde Excel): {categoriasPermitidas.Count}");
            LogDebug($"Ejemplos permitidos: {string.Join(", ", categoriasPermitidas.Take(5))}");

            // 4. COLECTAR ELEMENTOS DE LA VISTA
            FilteredElementCollector collector = new FilteredElementCollector(doc, vistaActiva.Id)
                .WhereElementIsNotElementType();

            LogDebug($"Total elementos en vista (bruto): {collector.GetElementCount()}");

            // Listas para procesar
            Dictionary<string, List<ElementId>> grupoA_ConValor = new Dictionary<string, List<ElementId>>();
            List<ElementId> grupoB_SinValor = new List<ElementId>();
            List<ElementId> grupoC_NoMetrada = new List<ElementId>();

            int elementosProcesados = 0;

            foreach (Element elem in collector)
            {
                if (elem.Category == null) continue;
                if (elem.Category.CategoryType != CategoryType.Model) continue; // Solo modelo 3D
                if (elem.IsHidden(vistaActiva)) continue; // Ignorar ocultos

                elementosProcesados++;
                string catRevit = elem.Category.Name.ToUpper().Trim();

                // --- DEBUG DETALLADO PARA LOS PRIMEROS 5 ELEMENTOS ---
                if (elementosProcesados <= 5)
                {
                    LogDebug($"Check Elem ID: {elem.Id} | Cat Revit: '{catRevit}' | ¿Está en lista?: {categoriasPermitidas.Contains(catRevit)}");
                }

                // LOGICA DE CLASIFICACIÓN
                if (!categoriasPermitidas.Contains(catRevit))
                {
                    // GRUPO C: Categoría NO está en el Excel -> NEGRO
                    grupoC_NoMetrada.Add(elem.Id);
                    continue;
                }

                // Si llegamos aquí, la categoría SÍ está en el Excel. Evaluamos el parámetro.
                bool tieneValor = false;
                string valorHallado = "";

                Parameter param = elem.LookupParameter(nombreParametro);
                if (param == null)
                {
                    // Intentar buscar en Tipo
                    ElementId typeId = elem.GetTypeId();
                    if (typeId != ElementId.InvalidElementId)
                    {
                        ElementType type = doc.GetElement(typeId) as ElementType;
                        param = type?.LookupParameter(nombreParametro);
                    }
                }

                if (param != null)
                {
                    valorHallado = ObtenerValorComoString(param);
                    if (!string.IsNullOrWhiteSpace(valorHallado))
                    {
                        if (!grupoA_ConValor.ContainsKey(valorHallado))
                            grupoA_ConValor[valorHallado] = new List<ElementId>();

                        grupoA_ConValor[valorHallado].Add(elem.Id);
                        tieneValor = true;
                    }
                }

                if (!tieneValor)
                {
                    // GRUPO B: Categoría está, pero parámetro vacío -> ROJO
                    grupoB_SinValor.Add(elem.Id);
                }
            }

            LogDebug($"--- RESUMEN CLASIFICACIÓN ---");
            LogDebug($"Grupo A (Colores/Match): {grupoA_ConValor.Values.Sum(x => x.Count)} elementos.");
            LogDebug($"Grupo B (Rojo/Sin Valor): {grupoB_SinValor.Count} elementos.");
            LogDebug($"Grupo C (Negro/No Listado): {grupoC_NoMetrada.Count} elementos.");

            // 5. APLICAR OVERRIDES (TRANSACCIÓN)
            using (Transaction trans = new Transaction(doc, "Aplicar Colores Debug"))
            {
                trans.Start();

                // Patrones
                ElementId solidFill = GetSolidFillPatternId(doc);

                // GRUPO A: Colores Paleta
                int colorIndex = 0;
                var valoresOrdenados = grupoA_ConValor.Keys.OrderBy(v => v).ToList();
                Dictionary<string, Color> coloresLeyenda = new Dictionary<string, Color>();

                foreach (string val in valoresOrdenados)
                {
                    Color colorFondo = GenerarColorPorIndice(colorIndex++);
                    coloresLeyenda[val] = colorFondo;
                    Color colorLinea = OscurecerColor(colorFondo);

                    OverrideGraphicSettings ogs = CrearOverride(colorFondo, colorLinea, solidFill);

                    foreach (ElementId id in grupoA_ConValor[val])
                    {
                        vistaActiva.SetElementOverrides(id, ogs);
                    }
                }

                // GRUPO B: Rojo
                if (grupoB_SinValor.Count > 0)
                {
                    Color rojo = new Color(255, 0, 0);
                    OverrideGraphicSettings ogsRojo = CrearOverride(rojo, new Color(127, 0, 0), solidFill);
                    foreach (ElementId id in grupoB_SinValor)
                        vistaActiva.SetElementOverrides(id, ogsRojo);
                }

                // GRUPO C: Negro
                if (grupoC_NoMetrada.Count > 0)
                {
                    Color negro = new Color(0, 0, 0); // Negro puro
                    Color grisOscuro = new Color(50, 50, 50); // Líneas gris oscuro para que se note algo
                    OverrideGraphicSettings ogsNegro = CrearOverride(negro, grisOscuro, solidFill);
                    foreach (ElementId id in grupoC_NoMetrada)
                        vistaActiva.SetElementOverrides(id, ogsNegro);
                }

                trans.Commit();
            }

            // Actualizar Leyenda UI (Pasar datos a la ventana si está abierta)
            // Nota: Aquí simplificamos para no depender de VentanaLeyenda en este snippet,
            // pero deberías mantener tu llamada original a VentanaLeyenda.ObtenerInstancia()...
            try
            {
                VentanaLeyenda ventana = VentanaLeyenda.ObtenerInstancia();
                // Si tienes un método para inicializar, úsalo. Si no, asegúrate de que VentanaLeyenda sea accesible.
                // ventana.InicializarEventHandler(uiApp); 
                ventana.ActualizarLeyenda(grupoA_ConValor, grupoB_SinValor, coloresLeyenda, vistaActiva);
                if (!ventana.IsVisible) ventana.Show();
            }
            catch (Exception ex)
            {
                LogDebug($"Error actualizando leyenda UI: {ex.Message}");
            }

            LogDebug("--- FIN PROCESO EXITOSO ---");
        }

        // --- MÉTODOS AUXILIARES ---

        private static void LogDebug(string linea)
        {
            try { File.AppendAllText(DebugLogPath, $"{DateTime.Now:HH:mm:ss} | {linea}\n"); } catch { }
        }

        private static HashSet<string> ObtenerCategoriasDesdeSheets()
        {
            HashSet<string> categorias = new HashSet<string>();
            try
            {
                LogDebug("Iniciando conexión a Google Sheets...");

                // Buscar credenciales (ajusta la ruta si es diferente en tu estructura final)
                string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string credPath = Path.Combine(assemblyFolder, "revitsheetsintegration-89c34b39c2ae.json");

                LogDebug($"Buscando credenciales en: {credPath}");
                if (!File.Exists(credPath))
                {
                    LogDebug("ERROR: No existe archivo de credenciales JSON.");
                    return categorias;
                }

                GoogleCredential credential;
                using (var stream = new FileStream(credPath, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(new[] { SheetsService.Scope.Spreadsheets });
                }

                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "RevitPluginAudit"
                });

                LogDebug("Conexión establecida. Descargando datos...");

                // Leer columnas A y B
                SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SPREADSHEET_ID, READ_RANGE);

                ValueRange response = request.Execute();
                IList<IList<object>> values = response.Values;

                if (values != null && values.Count > 0)
                {
                    bool listaEncontrada = false;
                    foreach (var row in values)
                    {
                        if (row.Count == 0) continue;

                        string colA = row[0].ToString().ToUpper().Trim();

                        // Buscamos el encabezado donde empieza la lista
                        if (colA.Contains("CATEGORIA") || colA.Contains("CATEGORIAS"))
                        {
                            listaEncontrada = true;
                            continue; // Saltar el encabezado
                        }

                        if (listaEncontrada)
                        {
                            // Si estamos debajo del encabezado, leemos la Columna B (índice 1)
                            if (row.Count > 1)
                            {
                                string catExcel = row[1].ToString().ToUpper().Trim();
                                if (!string.IsNullOrEmpty(catExcel))
                                {
                                    categorias.Add(catExcel);
                                }
                            }
                            // Si la Col A tiene texto nuevo (otro encabezado), podríamos parar, 
                            // pero por ahora asumimos que la lista sigue.
                        }
                    }
                }
                else
                {
                    LogDebug("ADVERTENCIA: Google Sheets devolvió 0 valores.");
                }
            }
            catch (Exception ex)
            {
                LogDebug($"ERROR LEYENDO GOOGLE SHEETS: {ex.Message}");
            }

            return categorias;
        }

        private static OverrideGraphicSettings CrearOverride(Color bg, Color line, ElementId patternId)
        {
            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            if (patternId != ElementId.InvalidElementId)
            {
                ogs.SetSurfaceForegroundPatternId(patternId);
                ogs.SetSurfaceForegroundPatternColor(bg);
                ogs.SetSurfaceForegroundPatternVisible(true);
                ogs.SetCutForegroundPatternId(patternId);
                ogs.SetCutForegroundPatternColor(bg);
                ogs.SetCutForegroundPatternVisible(true);
            }
            ogs.SetProjectionLineColor(line);
            ogs.SetCutLineColor(line);
            return ogs;
        }

        private static Color OscurecerColor(Color c)
        {
            return new Color((byte)(c.Red * 0.8), (byte)(c.Green * 0.8), (byte)(c.Blue * 0.8));
        }

        // ... (Mantén aquí tus métodos auxiliares existentes: LeerParametroGuardado, ObtenerValorComoString, GenerarColorPorIndice, GetSolidFillPatternId) ...
        // INCLUYE AQUÍ EL RESTO DE MÉTODOS AUXILIARES QUE YA TENÍAS EN TU ARCHIVO ORIGINAL
        // PARA AHORRAR ESPACIO NO LOS COPIÉ TODOS, PERO SON NECESARIOS.

        private static string LeerParametroGuardado()
        {
            try { return File.Exists(ConfigFilePath) ? File.ReadAllText(ConfigFilePath).Trim() : string.Empty; } catch { return string.Empty; }
        }

        private static string ObtenerValorComoString(Parameter param)
        {
            if (!param.HasValue) return string.Empty;
            switch (param.StorageType)
            {
                case StorageType.String: return param.AsString() ?? "";
                case StorageType.ElementId: return param.AsElementId().IntegerValue.ToString();
                default: return param.AsValueString() ?? ""; // AsValueString es mejor para Doubles/Integers
            }
        }

        private static Color GenerarColorPorIndice(int indice)
        {
            // Usando lógica simple para demo, puedes pegar tu paleta gigante aquí
            byte r = (byte)((indice * 50) % 255);
            byte g = (byte)((indice * 80) % 255);
            byte b = (byte)((indice * 110) % 255);
            return new Color(r, g, b);
        }

        private static ElementId GetSolidFillPatternId(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(x => x.GetFillPattern().IsSolidFill)?.Id ?? ElementId.InvalidElementId;
        }
    }
}